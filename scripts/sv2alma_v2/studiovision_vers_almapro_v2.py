"""
studiovision_vers_almapro_v2.py
Version 2 - Import direct MySQL AlmaPro depuis StudioVision
Usage principal : python studiovision_vers_almapro_v2.py --auto
"""

from __future__ import annotations
import argparse, os, re, shutil, sys, traceback
from datetime import datetime

# -- Configuration --------------------------------------------------------------
DEFAULT_MDB    = r"M:\fichier\PUBLIC.MDB"
DEFAULT_PHOTOS = r"M:\PHOTOS"
MYSQL_HOST     = "192.168.0.179"
FALLBACK_XML   = r"\\DESKTOP-MQKMFAQ\Users\Public\Documents\import"

# -- Imports modules locaux -----------------------------------------------------
try:
    from sv_reader import (
        get_active_patient, get_patient_by_code,
        get_patient_consultations, get_patient_folder_path,
        list_patient_files, _connect as sv_connect,
        _rows, _get, _fmt_date, PYODBC_OK, WIN32_OK,
    )
    from alma_mysql_writer import AlmaWriter
    from alma_xml import generate_xml
except ImportError as e:
    print(f"Erreur import : {e}")
    print("Verifiez que sv_reader.py, alma_mysql_writer.py et alma_xml.py")
    print("sont dans le meme dossier.")
    sys.exit(1)

# -- Utilitaires ----------------------------------------------------------------

def parse_list(text):
    if not text: return []
    return [i.strip(" -") for i in re.split(r"[,;\n/]", text) if len(i.strip(" -")) > 2]

def search_patients(mdb, nom, prenom):
    conn = sv_connect(mdb)
    try:
        parts, params = [], []
        if nom:
            parts.append("UCase([NOM]) LIKE ?")
            params.append(f"%{nom.upper()}%")
        if prenom:
            parts.append("UCase([Pr\xe9nom]) LIKE ?")
            params.append(f"%{prenom.upper()}%")
        where = ("WHERE " + " AND ".join(parts)) if parts else ""
        rows = _rows(conn, f"SELECT * FROM Patients {where}", params)
    finally:
        conn.close()
    return [{"code": _get(r,"Code patient"),
             "nom":  _get(r,"NOM").upper(),
             "prenom": _get(r,"Pr\xe9nom","Prenom"),
             "ddn":  _fmt_date(r.get("Date de naissance") or r.get("DDN"))}
            for r in rows]

def resolve_patient(mdb, nom, prenom, code, auto):
    """Trouve le patient - priorite : COM (auto) > code > nom+prenom"""
    # Mode AUTO : lire la fiche ouverte dans StudioVision
    if auto:
        if not WIN32_OK:
            print("  win32com non disponible.")
        else:
            active = get_active_patient()
            if active:
                code = active["code"]
                print(f"  Fiche StudioVision detectee : {active['nom']} {active['prenom']} (code {code})")
            else:
                print("  Aucun patient ouvert dans StudioVision.")
                print("  Ouvrez la fiche patient dans StudioVision puis relancez.")
                return None

    if code:
        p = get_patient_by_code(mdb, code)
        if not p: print(f"  Code {code} introuvable dans PUBLIC.MDB.")
        return p

    results = search_patients(mdb, nom, prenom)
    if not results:
        print(f"  Aucun patient : NOM='{nom}' Prenom='{prenom}'")
        return None
    if len(results) == 1:
        code = results[0]["code"]
        print(f"  Trouve : {results[0]['nom']} {results[0]['prenom']} (code {code})")
    else:
        print(f"  {len(results)} patients correspondent :")
        for i, r in enumerate(results):
            print(f"    [{i}]  {r['nom']:<20} {r['prenom']:<15} DDN={r['ddn']} Code={r['code']}")
        while True:
            try:
                idx = int(input(f"  Choisissez [0-{len(results)-1}] : "))
                if 0 <= idx < len(results):
                    code = results[idx]["code"]; break
            except (ValueError, KeyboardInterrupt):
                pass
    return get_patient_by_code(mdb, code)

def fallback_xml(patient, consults, files, atcd, allergies, dest):
    os.makedirs(dest, exist_ok=True)
    docext = os.path.join(dest, "docext"); os.makedirs(docext, exist_ok=True)
    n = 1
    while os.path.exists(os.path.join(dest, f"{n}.xml")): n += 1
    xml = generate_xml(patient=patient, consultations=consults,
                       documents=files, atcd=atcd, allergies=allergies)
    xml_path = os.path.join(dest, f"{n}.xml")
    with open(xml_path, "w", encoding="iso-8859-1", errors="replace") as f:
        f.write(xml)
    copied = 0
    for doc in files:
        dst = os.path.join(docext, doc["name"])
        if not os.path.exists(dst):
            try: shutil.copy2(str(doc["src"]), dst); copied += 1
            except Exception: pass
    return xml_path, copied

# -- Point d'entree -------------------------------------------------------------

def main():
    if not PYODBC_OK:
        print("pyodbc non installe. pip install pyodbc pywin32")
        sys.exit(1)

    parser = argparse.ArgumentParser(
        description="StudioVision -> AlmaPro v2 (MySQL direct)"
    )
    parser.add_argument("--auto", action="store_true",
        help="Lire la fiche ouverte dans StudioVision (recommande)")
    parser.add_argument("--nom",    default="")
    parser.add_argument("--prenom", default="")
    parser.add_argument("--code",   default="")
    parser.add_argument("--mdb",    default=DEFAULT_MDB)
    parser.add_argument("--photos", default=DEFAULT_PHOTOS)
    parser.add_argument("--mysql-host", default=MYSQL_HOST)
    parser.add_argument("--force",  action="store_true",
        help="Importer meme si le patient existe deja")
    parser.add_argument("--test-mysql", action="store_true",
        help="Tester la connexion MySQL uniquement")
    parser.add_argument("--list",   action="store_true",
        help="Lister les derniers patients dans AlmaPro")
    parser.add_argument("--schema", default="",
        help="Schema d'une table AlmaPro (ex: adm_patient)")
    args = parser.parse_args()

    host = args.mysql_host

    # -- Commandes utilitaires --------------------------------------------------

    if args.test_mysql:
        print(f"Test connexion MySQL {host}:3306 ...")
        try:
            w = AlmaWriter(host=host)
            w.connect()
            patients = w.list_patients(5)
            print(f"  Connexion OK ! {len(patients)} patient(s) recents :")
            for p in patients:
                print(f"    {p}")
            w.disconnect()
        except Exception as e:
            print(f"  ECHEC : {e}")
        return

    if args.list:
        print(f"Derniers patients AlmaPro (MySQL {host}) :")
        try:
            w = AlmaWriter(host=host)
            w.connect()
            for p in w.list_patients(20):
                print(f"  {p}")
            w.disconnect()
        except Exception as e:
            print(f"  ECHEC : {e}")
        return

    if args.schema:
        print(f"Schema de {args.schema} :")
        try:
            w = AlmaWriter(host=host)
            w.connect()
            for c in w.get_schema(args.schema):
                print(f"  {c}")
            w.disconnect()
        except Exception as e:
            print(f"  ECHEC : {e}")
        return

    # -- Import principal -------------------------------------------------------

    print("=" * 65)
    print("  StudioVision -> AlmaPro v2  (MySQL direct)")
    print("=" * 65)

    mdb = args.mdb
    if not os.path.exists(mdb):
        print(f"PUBLIC.MDB introuvable : {mdb}")
        print("Ajustez DEFAULT_MDB en haut du script ou utilisez --mdb")
        sys.exit(1)

    nom    = args.nom.strip().upper()
    prenom = args.prenom.strip()
    code   = args.code.strip()

    # Sans --auto et sans argument : demander si on doit lancer en mode auto
    if not args.auto and not nom and not code:
        print()
        if WIN32_OK:
            print("Mode AUTO disponible (lit la fiche ouverte dans StudioVision).")
            rep = input("Utiliser le patient ouvert dans StudioVision ? [O/n] : ").strip().lower()
            if rep in ("", "o", "oui", "y"):
                args.auto = True
        if not args.auto:
            nom    = input("Nom    (MAJUSCULES) : ").strip().upper()
            prenom = input("Prenom             : ").strip()
            if not nom:
                code = input("Code StudioVision  : ").strip()

    # Etape 1 : resoudre le patient
    print("\n[1/4] Identification du patient...")
    patient = resolve_patient(mdb, nom, prenom, code, args.auto)
    if not patient:
        sys.exit(1)
    print(f"  OK : {patient['nom']} {patient['prenom']}  DDN={patient['ddn']}")

    # Etape 2 : extraire les donnees
    print("\n[2/4] Extraction des donnees StudioVision...")
    pcode     = patient.get("code","")
    consults  = get_patient_consultations(mdb, pcode) if pcode else []
    print(f"  {len(consults)} consultation(s)")
    folder    = get_patient_folder_path(mdb, pcode, args.photos) if pcode else None
    files     = list_patient_files(folder) if folder else []
    print(f"  {len(files)} document(s)" + (f" dans {folder}" if folder else " (dossier photos introuvable)"))
    atcd      = parse_list(patient.get("antecedents",""))
    allergies = parse_list(patient.get("allergies",""))

    # Etape 3 : insertion MySQL
    print(f"\n[3/4] Insertion dans MySQL AlmaPro ({host})...")

    if args.force:
        AlmaWriter._patient_exists = lambda self, p: False

    num_dossier = None
    mysql_ok    = False
    try:
        w = AlmaWriter(host=host)
        w.connect()
        num_dossier = w.insert_patient(patient, consults, atcd, allergies)
        w.disconnect()
        mysql_ok = True
    except ValueError as e:
        print(f"\n  {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n  MySQL inaccessible : {e}")
        print(f"  Basculement en mode XML (fallback)...")
        try:
            xml_path, n_copied = fallback_xml(patient, consults, files,
                                               atcd, allergies, FALLBACK_XML)
            print(f"  XML : {xml_path}  ({n_copied} docs)")
            print(f"\n  Import manuel requis dans AlmaPro :")
            print(f"  Import -> Importer dossiers AlmaPro -> {FALLBACK_XML}")
        except Exception as e2:
            print(f"  Fallback echoue aussi : {e2}")
        sys.exit(1)

    # Etape 4 : copier les documents
    n_docs = 0
    if files and mysql_ok:
        print(f"\n[4/4] Copie des documents...")
        img_dest = os.path.join(
            f"\\\\{host}\\almapro\\imageexterne",
            f"REP_{num_dossier}"
        )
        try:
            os.makedirs(img_dest, exist_ok=True)
            for doc in files:
                dst = os.path.join(img_dest, doc["name"])
                if not os.path.exists(dst):
                    shutil.copy2(str(doc["src"]), dst)
                    n_docs += 1
            print(f"  {n_docs} document(s) -> {img_dest}")
        except Exception as e:
            print(f"  Copie impossible : {e}")
            print(f"  Copiez manuellement {folder}")
            print(f"  vers {img_dest}")
    else:
        print("\n[4/4] Pas de documents a copier.")

    # R-sum-
    print(f"\n{'=' * 65}")
    print(f"  IMPORT TERMINE - {patient['nom']} {patient['prenom']}")
    print(f"{'=' * 65}")
    print(f"""
  Dossier AlmaPro : #{num_dossier}
  Consultations   : {len(consults)}
  Documents       : {n_docs}

  Le patient est visible dans AlmaPro immediatement.
  Prochaine etape : AlmaPro -> dossier #{num_dossier} -> bouton DMP
""")


if __name__ == "__main__":
    main()
