"""
generer_xml_import_almapro.py
─────────────────────────────────────────────────────────────────────────────
Génère un fichier XML compatible "Import depuis autre logiciel" d'AlmaPro.

DEUX MODES :
  • Mode AUTO   : lit ibdata1 (MySQL du POSTE SERVEUR) pour en extraire les
                  données du patient, puis génère le XML
  • Mode DIRECT : génère le XML directement depuis les paramètres fournis,
                  sans lire ibdata1 — utile quand ibdata1 n'est pas accessible

Le mode AUTO est tenté en premier. Si ibdata1 ne contient pas le patient
(ex : vous êtes sur le poste CLIENT dont ibdata1 est vide), le script bascule
automatiquement en mode DIRECT après vous avoir expliqué la situation.

Format XML : spécification AlmaPro v3.2+
Encodage   : iso-8859-1  (obligatoire pour AlmaPro)
Dates      : jj/mm/aaaa  (obligatoire)

Usage :
    # Mode interactif — le script pose les questions
    python generer_xml_import_almapro.py

    # Ligne de commande, ibdata1 du serveur accessible
    python generer_xml_import_almapro.py \\
        --nom TEST --prenom Test --ddn 23/11/1990 \\
        --ibdata "C:\\Program Files\\MySQL\\MySQL Server 5.0\\data\\ibdata1" \\
        --out "C:\\ImportXMLAlmaPro" --num 1

    # Ligne de commande, mode direct (pas d'ibdata1)
    python generer_xml_import_almapro.py \\
        --nom TEST --prenom Test --ddn 23/11/1990 --sexe M \\
        --direct \\
        --out "C:\\ImportXMLAlmaPro" --num 1

    # ibdata1 sur le serveur réseau
    python generer_xml_import_almapro.py \\
        --nom MEGRET --prenom Leo --ddn 23/11/1999 \\
        --ibdata "\\\\192.168.0.179\\almapro\\..\\MySQL\\data\\ibdata1" \\
        --out "C:\\ImportXMLAlmaPro" --num 2
─────────────────────────────────────────────────────────────────────────────
"""

import re
import os
import sys
import argparse
from datetime import datetime, date
from xml.dom import minidom
import xml.etree.ElementTree as ET


# ══════════════════════════════════════════════════════════════════════════════
# DÉCODAGE INNODB DATE
# ══════════════════════════════════════════════════════════════════════════════

INNODB_DATE_XOR = 0x800000

def decode_innodb_date(b1: int, b2: int, b3: int) -> str | None:
    """
    Décode 3 octets InnoDB (type DATE) → chaîne jj/mm/aaaa.
    InnoDB encode : val = (year*512 + month*32 + day) XOR 0x800000
    """
    val   = ((b1 << 16) | (b2 << 8) | b3) ^ INNODB_DATE_XOR
    day   =  val & 0x1F
    month = (val >> 5) & 0x0F
    year  =  val >> 9
    if 1900 < year < 2100 and 1 <= month <= 12 and 1 <= day <= 31:
        try:
            date(year, month, day)
            return f"{day:02d}/{month:02d}/{year}"
        except ValueError:
            pass
    return None


def find_dates_in_chunk(chunk: bytes) -> list[str]:
    seen: list[str] = []
    for i in range(len(chunk) - 2):
        d = decode_innodb_date(chunk[i], chunk[i+1], chunk[i+2])
        if d and d not in seen:
            seen.append(d)
    return seen


# ══════════════════════════════════════════════════════════════════════════════
# DIAGNOSTIC IBDATA1
# ══════════════════════════════════════════════════════════════════════════════

# Patients de démonstration connus présents dans TOUT ibdata1 AlmaPro installé
# Utilisés pour valider que l'ibdata1 chargé est bien celui du serveur avec données
DEMO_PATIENTS = [
    ("DEMO",   "18/02/1976"),   # DEMO Paul
    ("DEMO",   "01/01/1960"),   # DEMO Elisabeth (approx)
]

def diagnose_ibdata(data: bytes) -> dict:
    """
    Analyse l'ibdata1 pour déterminer s'il contient des données patients.
    Retourne un dict avec : has_patients, size_mb, patients_found
    """
    size_mb = len(data) / 1024 / 1024

    # Chercher au moins un patient DEMO connu
    patients_found = []
    for nom, ddn in DEMO_PATIENTS:
        key = (nom + " " * 50)[:50].encode("latin-1")
        pos = 0
        while True:
            p = data.find(key, pos)
            if p == -1: break
            chunk = data[p:p+150]
            if ddn in find_dates_in_chunk(chunk):
                patients_found.append(f"{nom} DDN={ddn}")
                break
            pos = p + 1

    return {
        "size_mb":       size_mb,
        "has_patients":  len(patients_found) > 0,
        "patients_found": patients_found,
    }


# ══════════════════════════════════════════════════════════════════════════════
# RECHERCHE PATIENT DANS IBDATA1
# ══════════════════════════════════════════════════════════════════════════════

def find_patient(data: bytes, nom: str, ddn: str) -> int | None:
    """
    Recherche le patient dans ibdata1 par NOM (VARCHAR(50) paddé) + DDN InnoDB.
    Pas de contrainte sur le prénom (peut être vide pour patients StudioVision).
    Retourne l'offset ou None.
    """
    nom_key = (nom.upper() + " " * 50)[:50].encode("latin-1")
    start = 0
    while True:
        pos = data.find(nom_key, start)
        if pos == -1:
            break
        chunk = data[pos:pos+150]
        if ddn in find_dates_in_chunk(chunk):
            return pos
        start = pos + 1
    return None


def extract_patient_context(data: bytes, pos: int, nom: str, prenom: str) -> dict:
    """Extrait les données disponibles dans ibdata1 pour ce patient."""
    ctx = _empty_ctx()
    full = data.decode("latin-1", errors="replace")
    nom_up    = nom.upper()
    prenom_cap = prenom.capitalize()

    # Sexe (byte dans les 120 octets après le début de l'enregistrement)
    chunk = data[pos:pos+120]
    ctx["sexe"] = "F" if (b"\x46\x00" in chunk[100:115]) else "M"

    # Journal d'audit → dates de consultation
    pat = re.compile(
        r"Le (\d{2}/\d{2}/\d{4}) à \d{2}:\d{2}:\d{2}, "
        r"\w[\w ]+ a (?:ouvert|modifié) le dossier patient de "
        + re.escape(nom_up) + r"(?:\s+" + re.escape(prenom_cap) + r")?"
    )
    seen: set[str] = set()
    for m in pat.finditer(full):
        d = m.group(1)
        if d not in seen:
            seen.add(d)
            ctx["consultations"].append({
                "date":  d,
                "titre": "Consultation (import StudioVision)",
                "texte": (f"Accès au dossier le {d} — données issues de l'import "
                          f"StudioVision vers AlmaPro."),
            })

    ctx["consultations"].sort(key=lambda x: x["date"])
    print(f"  → {len(ctx['consultations'])} consultation(s) extraite(s) du journal")
    return ctx


# ══════════════════════════════════════════════════════════════════════════════
# GÉNÉRATION XML
# ══════════════════════════════════════════════════════════════════════════════

def _empty_ctx() -> dict:
    return {
        "sexe": "M", "num_secu": "", "adresse": "", "ville": "", "cp": "",
        "tel": "", "email": "", "profession": "", "remarque": "",
        "consultations": [], "ordonnances": [], "biologie": [],
        "vaccins": [], "atcd": [], "allergies": [], "notes": [],
    }


def sub(parent: ET.Element, tag: str, text: str = "") -> ET.Element:
    el = ET.SubElement(parent, tag)
    el.text = text or ""
    return el


def generate_xml(nom: str, prenom: str, ddn: str, ctx: dict) -> str:
    """
    Génère le XML au format AlmaPro v3.2+ (iso-8859-1, jj/mm/aaaa).
    Respecte strictement l'ordre des balises indiqué dans la documentation.
    """
    root = ET.Element("Export_XML_AlmaPro")

    # ── Données administratives ───────────────────────────────────────────
    d = ET.SubElement(root, "Dossiers")
    sub(d, "Nom",                    nom.upper())
    sub(d, "Prenom",                 prenom)
    sub(d, "Sexe",                   ctx.get("sexe", "M"))
    sub(d, "Adresse",                ctx.get("adresse", ""))
    sub(d, "AdresseL2",              "")
    sub(d, "Ville",                  ctx.get("ville", ""))
    sub(d, "CodePostal",             ctx.get("cp", ""))
    sub(d, "Tel_Domicile",           ctx.get("tel", ""))
    sub(d, "Tel_Pro",                "")
    sub(d, "Tel_Port",               "")
    sub(d, "FAX",                    "")
    sub(d, "EMail",                  ctx.get("email", ""))
    sub(d, "Nom_Jeune_Fille",        "")
    sub(d, "Date_de_naissance",      ddn)
    sub(d, "Date_Contrat_Signe",     "")
    sub(d, "ContratSigne",           "FALSE")
    sub(d, "Decede",                 "FALSE")
    sub(d, "Profession_du_patient",  ctx.get("profession", ""))
    sub(d, "Num_Secu",               ctx.get("num_secu", ""))
    remarque = ctx.get("remarque") or (
        f"Dossier importé depuis StudioVision "
        f"le {datetime.now().strftime('%d/%m/%Y')}."
    )
    sub(d, "Remarque", remarque)

    # ── Consultations ─────────────────────────────────────────────────────
    consultations = ctx.get("consultations") or [{
        "date":  datetime.now().strftime("%d/%m/%Y"),
        "titre": "Import StudioVision",
        "texte": (f"Dossier importé depuis StudioVision.\n"
                  f"Patient : {nom.upper()} {prenom}\n"
                  f"Date de naissance : {ddn}\n"
                  f"Import le {datetime.now().strftime('%d/%m/%Y à %H:%M')}"),
    }]
    for c in consultations:
        bloc = ET.SubElement(root, "Dossier_Consultations")
        sub(bloc, "Date_Contact",  c.get("date", "01/01/2000"))
        if c.get("titre"):
            sub(bloc, "Titre_Contact", c["titre"])
        sub(bloc, "Texte_Contact", c.get("texte", ""))

    # ── Ordonnances ───────────────────────────────────────────────────────
    for i, o in enumerate(ctx.get("ordonnances", []), start=1):
        bloc = ET.SubElement(root, "Ordonnance")
        sub(bloc, "Date_Ordonnance",  o.get("date", "01/01/2000"))
        sub(bloc, "Texte_Ordonnance", o.get("texte", ""))
        sub(bloc, "ID_Ordonnance",    str(o.get("id", i)))

    # ── Résultats biologie ────────────────────────────────────────────────
    for b in ctx.get("biologie", []):
        bloc = ET.SubElement(root, "Resultat_Biologie")
        sub(bloc, "Date_Resultat_Biologie",   b.get("date", "01/01/2000"))
        sub(bloc, "Examen_Resultat_Biologie", b.get("examen", ""))
        sub(bloc, "Valeur_Resultat_Biologie", b.get("valeur", ""))
        sub(bloc, "Unite_Resultat_Biologie",  b.get("unite", ""))
        sub(bloc, "ID_Resultat_Biologie",     str(b.get("id", 1)))

    # ── Antécédents ───────────────────────────────────────────────────────
    for a in ctx.get("atcd", []):
        ET.SubElement(ET.SubElement(root, "ATCD"), "Texte_ATCD").text = a

    # ── Allergies ─────────────────────────────────────────────────────────
    for a in ctx.get("allergies", []):
        ET.SubElement(ET.SubElement(root, "Allergie"), "Texte_Allergie").text = a

    # ── Notes ─────────────────────────────────────────────────────────────
    for n in ctx.get("notes", []):
        ET.SubElement(ET.SubElement(root, "Note"), "Texte_Note").text = n

    # ── Vaccins ───────────────────────────────────────────────────────────
    for v in ctx.get("vaccins", []):
        bloc = ET.SubElement(root, "Vaccin")
        sub(bloc, "Date_Vaccinatio",         v.get("date", "01/01/2000"))
        sub(bloc, "Nom_Vaccin",              v.get("nom", ""))
        sub(bloc, "Information_Vaccination", v.get("info", ""))

    # ── Sérialisation iso-8859-1 ──────────────────────────────────────────
    raw   = ET.tostring(root, encoding="unicode")
    dom   = minidom.parseString(raw)
    lines = (dom.toprettyxml(indent="   ", encoding="iso-8859-1")
               .decode("iso-8859-1", errors="replace").split("\n"))
    lines[0] = '<?xml version="1.0" encoding="iso-8859-1" standalone="yes"?>'
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# ÉCRITURE DU DOSSIER DE SORTIE
# ══════════════════════════════════════════════════════════════════════════════

def write_output(output_dir: str, num: int, xml_content: str) -> str:
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(os.path.join(output_dir, "docext"), exist_ok=True)
    xml_path = os.path.join(output_dir, f"{num}.xml")
    with open(xml_path, "w", encoding="iso-8859-1", errors="replace") as f:
        f.write(xml_content)
    return xml_path


# ══════════════════════════════════════════════════════════════════════════════
# INSTRUCTIONS FINALES
# ══════════════════════════════════════════════════════════════════════════════

def print_instructions(nom: str, prenom: str, ddn: str,
                       output_dir: str, xml_path: str, mode: str):
    sep = "=" * 65
    print(f"\n{sep}")
    print(f"  FICHIER PRÊT  [{mode}]  — {nom} {prenom} ({ddn})")
    print(sep)
    print(f"""
  Fichier généré : {xml_path}

  ─────────────────────────────────────────────────────────────
  ÉTAPES D'IMPORT DANS ALMAPRO  (à faire sur le POSTE SERVEUR)
  ─────────────────────────────────────────────────────────────

  1. Copiez CE dossier sur le poste serveur :
       {output_dir}
     → placer à cet emplacement sur le serveur :
       C:\\ImportXMLAlmaPro

  2. Ouvrez AlmaPro sur le serveur.

  3. Écran sélection patient → bouton "Import"
     → choisissez "Importer des dossiers AlmaPro"

  4. Dans "Chemin des fichiers d'export AlmaPro"
     → cliquez "Choisir" et sélectionnez :
       C:\\ImportXMLAlmaPro

  5. Cochez la ligne  {nom} {prenom}

  6. Cliquez "Importer les dossiers cochés"
  ─────────────────────────────────────────────────────────────
""")


# ══════════════════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Génère un XML d'import AlmaPro (depuis ibdata1 ou en direct)."
    )
    parser.add_argument("--nom",    default=None)
    parser.add_argument("--prenom", default=None)
    parser.add_argument("--ddn",    default=None, help="jj/mm/aaaa")
    parser.add_argument("--sexe",   default="M",  choices=["M", "F"])
    parser.add_argument("--ibdata",
        default=r"C:\Program Files\MySQL\MySQL Server 5.0\data\ibdata1",
        help="Chemin vers ibdata1 du POSTE SERVEUR")
    parser.add_argument("--out",    default=r"C:\ImportXMLAlmaPro")
    parser.add_argument("--num",    type=int, default=1,
        help="Numéro du fichier XML (1.xml, 2.xml…)")
    parser.add_argument("--direct", action="store_true",
        help="Génère le XML directement sans lire ibdata1")
    args = parser.parse_args()

    print("=" * 65)
    print("  Générateur XML import AlmaPro")
    print("=" * 65)

    # Saisie interactive si arguments manquants
    nom    = args.nom    or input("\nNom du patient (MAJUSCULES) : ").strip().upper()
    prenom = args.prenom or input("Prénom                      : ").strip()
    ddn    = args.ddn    or input("Date de naissance jj/mm/aaaa: ").strip()

    try:
        datetime.strptime(ddn, "%d/%m/%Y")
    except ValueError:
        print(f"\n✗ Date invalide '{ddn}' — format attendu : jj/mm/aaaa")
        sys.exit(1)

    output_dir  = args.out
    num         = args.num
    ibdata_path = args.ibdata
    force_direct = args.direct

    print(f"\nPatient : {nom} {prenom}  DDN={ddn}")
    print(f"Sortie  : {output_dir}{os.sep}{num}.xml\n")

    ctx = None
    mode = "DIRECT"

    # ── Tentative mode AUTO ───────────────────────────────────────────────
    if not force_direct:
        print("─" * 65)
        print(f"1/5  ibdata1 → {ibdata_path}")

        if not os.path.exists(ibdata_path):
            print(f"  ⚠  Fichier introuvable.")
            print(f"     → Basculement en mode DIRECT (XML généré sans ibdata1)")
        else:
            size_mb = os.path.getsize(ibdata_path) / 1024 / 1024
            print(f"  Lecture ({size_mb:.0f} Mo)…", end=" ", flush=True)
            with open(ibdata_path, "rb") as f:
                data = f.read()
            print("OK")

            # Diagnostic : cet ibdata1 contient-il des patients ?
            print("\n2/5  Diagnostic de l'ibdata1 chargé")
            diag = diagnose_ibdata(data)
            if not diag["has_patients"]:
                print(f"  ⚠  Cet ibdata1 ({size_mb:.0f} Mo) ne contient PAS de données patients.")
                print()
                print("  EXPLICATION :")
                print("  Vous êtes probablement sur le POSTE CLIENT.")
                print("  Sur le poste client, ibdata1 est vide car les données")
                print("  patients vivent UNIQUEMENT sur le POSTE SERVEUR.")
                print()
                print("  SOLUTIONS :")
                print("  A) Exécutez ce script directement sur le POSTE SERVEUR.")
                print()
                print(f"  B) Copiez ibdata1 du serveur vers ce poste puis relancez :")
                print(f"     Serveur  : C:\\Program Files\\MySQL\\MySQL Server 5.0\\data\\ibdata1")
                print(f"     Commande : python {os.path.basename(sys.argv[0])} \\")
                print(f"                  --nom {nom} --prenom {prenom} --ddn {ddn} \\")
                print(f"                  --ibdata <chemin_ibdata1_copié>")
                print()
                print("  C) Générez le XML sans ibdata1 (données minimales) :")
                print(f"     python {os.path.basename(sys.argv[0])} \\")
                print(f"                  --nom {nom} --prenom {prenom} --ddn {ddn} \\")
                print(f"                  --sexe {args.sexe} --direct")
                print()
                rep = input("  Continuer en mode DIRECT (données minimales) ? [O/n] : ").strip().lower()
                if rep in ("", "o", "oui", "y", "yes"):
                    force_direct = True
                else:
                    print("  Abandon.")
                    sys.exit(0)
            else:
                print(f"  ✓ ibdata1 valide ({size_mb:.0f} Mo) — patients DEMO détectés")

                print("\n3/5  Recherche du patient")
                pos = find_patient(data, nom, ddn)
                if pos is None:
                    print(f"  ✗ Patient {nom} / DDN={ddn} absent de cet ibdata1.")
                    print(f"  → Basculement en mode DIRECT")
                    force_direct = True
                else:
                    print(f"  ✓ Trouvé à l'offset {pos:,}")
                    print("\n4/5  Extraction des données")
                    ctx = extract_patient_context(data, pos, nom, prenom)
                    mode = "AUTO (ibdata1)"

    # ── Mode DIRECT (fallback ou forcé) ──────────────────────────────────
    if force_direct or ctx is None:
        mode = "DIRECT"
        ctx = _empty_ctx()
        ctx["sexe"] = args.sexe
        ctx["remarque"] = (
            f"Dossier importé depuis StudioVision "
            f"le {datetime.now().strftime('%d/%m/%Y')} (mode direct)."
        )
        print(f"\n{'─'*65}")
        print("Mode DIRECT — XML généré depuis les paramètres saisis")
        print("(données minimales : nom, prénom, date de naissance, sexe)")

    # ── Génération XML ────────────────────────────────────────────────────
    step = "5" if not force_direct else "2"
    print(f"\n{step}/5  Génération du XML (iso-8859-1)")
    xml_content = generate_xml(nom, prenom, ddn, ctx)
    print("  ✓ XML généré")

    # ── Écriture ──────────────────────────────────────────────────────────
    print(f"\nÉcriture des fichiers")
    xml_path = write_output(output_dir, num, xml_content)
    print(f"  ✓ {xml_path}")
    print(f"  ✓ {os.path.join(output_dir, 'docext')}{os.sep}")

    # ── Aperçu XML ────────────────────────────────────────────────────────
    print_instructions(nom, prenom, ddn, output_dir, xml_path, mode)
    print("  Aperçu du fichier XML :")
    print("  " + "─" * 60)
    for line in xml_content.split("\n")[:55]:
        print(f"  {line}")


if __name__ == "__main__":
    main()
