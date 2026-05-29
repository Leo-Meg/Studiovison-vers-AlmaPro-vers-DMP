"""
studiovision_vers_almapro.py
═══════════════════════════════════════════════════════════════════════════════
Export d'un dossier patient de StudioVision → XML importable dans AlmaPro.
Copie le résultat vers le partage réseau du serveur AlmaPro.

ARCHITECTURE
    sv_reader.py   — lecture de PUBLIC.MDB + COM (fiche ouverte)
    alma_xml.py    — génération du XML format AlmaPro v3.2+
    ↑ ce fichier   — orchestration, CLI, copie réseau

DEUX MODES DE RECHERCHE DU PATIENT
    1. AUTO  — lit la fiche actuellement ouverte dans StudioVision via COM
               (requiert win32com + StudioVision ouvert sur le bon patient)
    2. MANUEL — recherche par nom/prénom ou code dans PUBLIC.MDB

USAGE
    # Mode auto : patient ouvert dans StudioVision
    python studiovision_vers_almapro.py --auto

    # Mode manuel — nom + prénom
    python studiovision_vers_almapro.py --nom TEST --prenom Test

    # Mode manuel — code patient StudioVision (plus fiable)
    python studiovision_vers_almapro.py --code 12345

    # Inspecter les tables disponibles dans PUBLIC.MDB
    python studiovision_vers_almapro.py --schema

PRÉREQUIS (à installer une seule fois sur le poste client)
    pip install pyodbc pywin32
    + Microsoft Access Database Engine (64 bits) :
      https://www.microsoft.com/en-us/download/details.aspx?id=54920
═══════════════════════════════════════════════════════════════════════════════
"""

from __future__ import annotations
import argparse
import os
import re
import shutil
import sys
import traceback
from datetime import datetime
from pathlib import Path

# Modules du projet
try:
    from sv_reader import (
        get_active_patient,
        get_patient_by_code,
        get_patient_consultations,
        get_patient_documents,
        get_patient_folder_path,
        list_patient_files,
        _connect, _rows, _get, _fmt_date,
        PYODBC_OK, WIN32_OK,
    )
    from alma_xml import generate_xml
except ImportError as e:
    print(f"✗ Erreur d'import : {e}")
    print("  Assurez-vous que sv_reader.py et alma_xml.py sont dans le même dossier.")
    sys.exit(1)


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION  (à adapter à votre installation)
# ══════════════════════════════════════════════════════════════════════════════

# Chemin vers PUBLIC.MDB sur le poste client
DEFAULT_MDB = r"C:\StudioVision\PUBLIC.MDB"

# Racine des dossiers photos patients StudioVision
DEFAULT_PHOTOS = r"C:\StudioVision\Photos"

# Dossier partagé sur le serveur AlmaPro où déposer les exports
# (ce chemin est directement accessible depuis le poste client via le réseau)
DEFAULT_DEST = r"\\DESKTOP-MQKMFAQ\Users\Public\Documents\import"


# ══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES
# ══════════════════════════════════════════════════════════════════════════════

def _next_num(folder: str) -> int:
    """Retourne le prochain numéro libre pour <n>.xml dans le dossier."""
    n = 1
    while os.path.exists(os.path.join(folder, f"{n}.xml")):
        n += 1
    return n


def _write_export(dest_dir: str, num: int,
                  xml_content: str,
                  file_list: list[dict]) -> tuple[str, int]:
    """
    Écrit le dossier d'export AlmaPro :
        <dest_dir>/<num>.xml
        <dest_dir>/docext/<fichier1>
        ...
    Retourne (chemin_xml, nb_fichiers_copiés).
    """
    os.makedirs(dest_dir, exist_ok=True)
    docext = os.path.join(dest_dir, "docext")
    os.makedirs(docext, exist_ok=True)

    xml_path = os.path.join(dest_dir, f"{num}.xml")
    with open(xml_path, "w", encoding="iso-8859-1", errors="replace") as f:
        f.write(xml_content)

    copied = 0
    for doc in file_list:
        dst = os.path.join(docext, doc["name"])
        if os.path.exists(dst):
            base, ext = os.path.splitext(doc["name"])
            ts  = datetime.now().strftime("%Y%m%d%H%M%S")
            dst = os.path.join(docext, f"{base}_{ts}{ext}")
        try:
            shutil.copy2(str(doc["src"]), dst)
            copied += 1
        except Exception as e:
            print(f"  ⚠  Copie échouée : {doc['name']} → {e}")

    return xml_path, copied


def _search_patient_by_name(mdb_path: str, nom: str, prenom: str) -> list[dict]:
    """
    Recherche dans PUBLIC.MDB par nom + prénom.
    Retourne une liste de patients (code + identité minimale).
    Ne liste PAS tous les patients : retourne liste vide si aucun résultat.
    """
    conn = _connect(mdb_path)
    try:
        # Requête Access avec LIKE insensible à la casse via UCase
        where_parts, params = [], []
        if nom:
            where_parts.append("UCase([NOM]) LIKE ?")
            params.append(f"%{nom.upper()}%")
        if prenom:
            where_parts.append("UCase([Prénom]) LIKE ?")
            params.append(f"%{prenom.upper()}%")
        where = ("WHERE " + " AND ".join(where_parts)) if where_parts else ""
        rows  = _rows(conn, f"SELECT * FROM Patients {where}", params)
    finally:
        conn.close()

    return [
        {
            "code":   _get(r, "Code patient"),
            "nom":    _get(r, "NOM").upper(),
            "prenom": _get(r, "Prénom", "PRENOM", "Prenom"),
            "ddn":    _fmt_date(r.get("Date de naissance") or r.get("DDN")),
        }
        for r in rows
    ]


def _inspect_schema(mdb_path: str) -> None:
    """Affiche le schéma de PUBLIC.MDB (tables + colonnes)."""
    conn = _connect(mdb_path)
    cur  = conn.cursor()
    tables = [t.table_name for t in cur.tables(tableType="TABLE")]
    print(f"\nSchéma de {mdb_path}  ({len(tables)} tables)\n")
    for t in sorted(tables):
        try:
            cur.execute(f"SELECT TOP 1 * FROM [{t}]")
            cols = [c[0] for c in cur.description]
            print(f"  {t:<35} {', '.join(cols)}")
        except Exception:
            print(f"  {t:<35} (lecture impossible)")
    conn.close()


def _build_atcd_allergies(patient: dict) -> tuple[list[str], list[str]]:
    """Transforme les champs texte ATCD / Allergies en listes."""
    atcd = []
    if patient.get("antecedents"):
        # Séparer par ponctuation courante ou saut de ligne
        for item in re.split(r"[;\n/]|(?:\r\n)", patient["antecedents"]):
            item = item.strip(" -•·")
            if len(item) > 2:
                atcd.append(item[:200])

    allergies = []
    if patient.get("allergies"):
        for item in re.split(r"[,;\n/]", patient["allergies"]):
            item = item.strip(" -•·")
            if len(item) > 2:
                allergies.append(item[:100])

    return atcd, allergies


def _print_instructions(nom: str, prenom: str, dest: str,
                         xml_path: str, n_docs: int) -> None:
    sep = "=" * 70
    print(f"\n{sep}")
    print(f"  EXPORT TERMINÉ — {nom} {prenom}")
    print(sep)
    print(f"""
  Fichier XML  : {xml_path}
  Documents    : {n_docs} fichier(s) dans {os.path.join(dest, 'docext')}

  ──────────────────────────────────────────────────────────────────
  IMPORT DANS ALMAPRO  (sur le POSTE SERVEUR)
  ──────────────────────────────────────────────────────────────────

  1. Ouvrez AlmaPro → écran de sélection patient
     → bouton "Import" → "Importer des dossiers AlmaPro"

  2. Chemin des fichiers d'export AlmaPro :
     → sélectionnez  {dest}

  3. Cochez la ligne  {nom} {prenom}

  4. Cliquez "Importer les dossiers cochés"
  ──────────────────────────────────────────────────────────────────
""")


# ══════════════════════════════════════════════════════════════════════════════
# ORCHESTRATION PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════

def run(mdb_path: str,
        photos_root: str,
        dest_dir: str,
        num: int | None = None,
        nom: str = "",
        prenom: str = "",
        code: str = "",
        auto: bool = False) -> None:

    print("=" * 70)
    print("  StudioVision → AlmaPro")
    print("=" * 70)

    # ── ÉTAPE 1 : Obtenir le code patient ─────────────────────────────────
    patient_info: dict | None = None

    # Mode AUTO : lire la fiche ouverte dans StudioVision via COM
    if auto:
        print("\n[1/5] Lecture de la fiche ouverte dans StudioVision…")
        if not WIN32_OK:
            print("  ⚠  win32com non disponible — basculement en mode manuel.")
            auto = False
        else:
            active = get_active_patient()
            if active:
                code = active["code"]
                print(f"  ✓ Fiche ouverte : {active['nom']} {active['prenom']} "
                      f"(code {code})")
            else:
                print("  ✗ Aucun patient ouvert dans StudioVision.")
                print("  → Ouvrez la fiche patient dans StudioVision puis relancez,")
                print("    ou utilisez --nom / --code pour une recherche manuelle.")
                sys.exit(1)

    # Mode MANUEL : recherche par code direct
    if not auto and code:
        print(f"\n[1/5] Recherche par code {code}…")

    # Mode MANUEL : recherche par nom + prénom
    if not auto and not code:
        print(f"\n[1/5] Recherche de '{nom} {prenom}' dans PUBLIC.MDB…")
        results = _search_patient_by_name(mdb_path, nom, prenom)

        if not results:
            print(f"  ✗ Aucun patient trouvé pour : "
                  f"NOM='{nom.upper()}' Prénom='{prenom}'")
            print("  Conseils :")
            print("  • Vérifiez l'orthographe (essayez avec seulement le nom)")
            print("  • Utilisez --code si vous connaissez le code StudioVision")
            print("  • Utilisez --schema pour vérifier les colonnes disponibles")
            sys.exit(1)

        if len(results) > 1:
            # Afficher uniquement les correspondances, pas toute la base
            print(f"  {len(results)} patient(s) correspondent :")
            for i, r in enumerate(results):
                print(f"    [{i}]  {r['nom']:<20} {r['prenom']:<15} "
                      f"DDN={r['ddn']:<12} Code={r['code']}")
            while True:
                try:
                    idx = int(input(f"\n  Choisissez [0-{len(results)-1}] : "))
                    if 0 <= idx < len(results):
                        code = results[idx]["code"]
                        break
                except (ValueError, KeyboardInterrupt):
                    pass
                print("  Choix invalide.")
        else:
            code = results[0]["code"]
            print(f"  ✓ Trouvé : {results[0]['nom']} {results[0]['prenom']} "
                  f"(code {code})")

    # ── ÉTAPE 2 : Charger la fiche complète depuis PUBLIC.MDB ─────────────
    print(f"\n[2/5] Chargement de la fiche patient (code {code})…")
    patient = get_patient_by_code(mdb_path, code)
    if not patient:
        print(f"  ✗ Patient code={code} introuvable dans PUBLIC.MDB.")
        sys.exit(1)
    print(f"  ✓ {patient['nom']} {patient['prenom']}  DDN={patient['ddn']}")

    # ── ÉTAPE 3 : Extraire consultations + documents ───────────────────────
    print("\n[3/5] Extraction des données médicales…")
    consultations = get_patient_consultations(mdb_path, code)
    print(f"  ✓ {len(consultations)} consultation(s)")

    db_docs = get_patient_documents(mdb_path, code)
    print(f"  ✓ {len(db_docs)} entrée(s) dans la table Documents")

    # Résoudre le dossier photos sur le disque
    folder = get_patient_folder_path(mdb_path, code, photos_root)
    if folder:
        file_list = list_patient_files(folder)
        print(f"  ✓ Dossier : {folder}")
        print(f"  ✓ {len(file_list)} fichier(s) image/PDF")
    else:
        file_list = []
        print(f"  ⚠  Dossier photos introuvable pour ce patient")
        print(f"     (photos_root = {photos_root})")

    # ── ÉTAPE 4 : Générer le XML ───────────────────────────────────────────
    print("\n[4/5] Génération du XML AlmaPro (iso-8859-1)…")
    atcd, allergies = _build_atcd_allergies(patient)
    xml_content = generate_xml(
        patient=patient,
        consultations=consultations,
        documents=file_list,
        atcd=atcd,
        allergies=allergies,
    )
    print("  ✓ XML généré")

    # ── ÉTAPE 5 : Écriture vers le partage réseau ──────────────────────────
    if num is None:
        num = _next_num(dest_dir)

    print(f"\n[5/5] Écriture vers {dest_dir}\\{num}.xml…")
    try:
        xml_path, n_copied = _write_export(dest_dir, num, xml_content, file_list)
        print(f"  ✓ {xml_path}")
        print(f"  ✓ {n_copied} document(s) copié(s) → docext/")
    except Exception as e:
        print(f"\n  ✗ Erreur lors de l'écriture : {e}")
        print(f"  Vérifiez que le partage réseau est accessible : {dest_dir}")
        traceback.print_exc()
        sys.exit(1)

    # Aperçu XML
    print("\n  Aperçu du XML (premières lignes) :")
    print("  " + "─" * 65)
    for line in xml_content.split("\n")[:55]:
        print(f"  {line}")

    _print_instructions(patient["nom"], patient["prenom"],
                        dest_dir, xml_path, n_copied)


# ══════════════════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE CLI
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    if not PYODBC_OK:
        print("✗ pyodbc non installé.")
        print("  pip install pyodbc pywin32")
        print("  + https://www.microsoft.com/en-us/download/details.aspx?id=54920")
        sys.exit(1)

    parser = argparse.ArgumentParser(
        description="Export StudioVision → XML importable dans AlmaPro"
    )
    parser.add_argument("--auto",   action="store_true",
        help="Lire la fiche ouverte dans StudioVision (COM)")
    parser.add_argument("--nom",    default="")
    parser.add_argument("--prenom", default="")
    parser.add_argument("--code",   default="",
        help="Code patient StudioVision (plus fiable que nom/prénom)")
    parser.add_argument("--mdb",    default=DEFAULT_MDB,
        help=f"Chemin vers PUBLIC.MDB (défaut : {DEFAULT_MDB})")
    parser.add_argument("--photos", default=DEFAULT_PHOTOS,
        help=f"Racine dossiers photos (défaut : {DEFAULT_PHOTOS})")
    parser.add_argument("--dest",   default=DEFAULT_DEST,
        help=f"Destination réseau (défaut : {DEFAULT_DEST})")
    parser.add_argument("--num",    type=int, default=None,
        help="Numéro du fichier XML (auto si absent)")
    parser.add_argument("--schema", action="store_true",
        help="Afficher le schéma de PUBLIC.MDB et quitter")
    args = parser.parse_args()

    mdb = args.mdb
    if not os.path.exists(mdb):
        print(f"✗ PUBLIC.MDB introuvable : {mdb}")
        print("  Ajustez --mdb ou DEFAULT_MDB en haut du script.")
        sys.exit(1)

    if args.schema:
        _inspect_schema(mdb)
        return

    nom    = args.nom.strip().upper()
    prenom = args.prenom.strip()
    code   = args.code.strip()

    # Saisie interactive si rien de fourni
    if not args.auto and not nom and not code:
        print("\nRecherche du patient :")
        nom    = input("  Nom    (MAJUSCULES) : ").strip().upper()
        prenom = input("  Prénom             : ").strip()
        if not nom:
            code = input("  Code StudioVision  : ").strip()

    run(
        mdb_path=mdb,
        photos_root=args.photos,
        dest_dir=args.dest,
        num=args.num,
        nom=nom,
        prenom=prenom,
        code=code,
        auto=args.auto,
    )


if __name__ == "__main__":
    main()
