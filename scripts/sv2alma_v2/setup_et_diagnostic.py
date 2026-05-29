"""
setup_et_diagnostic.py — utilise PyMySQL (compatible MySQL 5.0)
"""
import sys, os, json, subprocess, socket

HOST = "192.168.0.179"
PORT = 3306

print("=" * 65)
print("  Diagnostic sv2alma_v2")
print("=" * 65)

# ── PyMySQL ────────────────────────────────────────────────────────
print("\n[1/4] PyMySQL ...", end=" ", flush=True)
try:
    import pymysql
    print(f"OK — {pymysql.__version__}")
except ImportError:
    print("installation...")
    subprocess.run([sys.executable, "-m", "pip", "install", "pymysql"], check=True)
    import pymysql
    print(f"OK — {pymysql.__version__}")

# ── MySQL ──────────────────────────────────────────────────────────
print(f"\n[2/4] Connexion MySQL AlmaPro ({HOST}:{PORT}) ...")
schemas = {}
mysql_ok = False
try:
    conn = pymysql.connect(
        host=HOST, port=PORT, user="root", password="",
        database="almapro", charset="latin1", connect_timeout=10,
    )
    cur = conn.cursor()
    cur.execute("SHOW TABLES")
    tables = [r[0] for r in cur.fetchall()]
    print(f"  ✓ Connecté — {len(tables)} tables")

    for table in ["adm_patient", "consu_contact",
                  "adm_administratif",
                  "adm_acces_et_les_modifications_indentite"]:
        if table in tables:
            cur.execute(f"DESCRIBE `{table}`")
            cols = [r[0] for r in cur.fetchall()]
            schemas[table] = cols
            print(f"  ✓ {table} ({len(cols)} colonnes)")
            print(f"       {', '.join(cols[:6])}{'…' if len(cols)>6 else ''}")
        else:
            print(f"  ✗ {table} — table absente")

    cur.execute("SELECT COUNT(*) FROM adm_patient")
    print(f"\n  Patients actuels : {cur.fetchone()[0]}")
    cur.close(); conn.close()
    mysql_ok = True

except pymysql.err.OperationalError as e:
    code, msg = e.args
    print(f"  ✗ [{code}] {msg}")
    if code == 2003:
        print("  → Pare-feu VM : exécutez en PowerShell admin sur la VM :")
        print("    New-NetFirewallRule -DisplayName 'MySQL 3306' `")
        print("      -Direction Inbound -Protocol TCP -LocalPort 3306 -Action Allow")
    elif code == 1045:
        print("  → Mot de passe requis. Cherchez dans :")
        print("    C:\\almapro\\almapro.ini  ou AlmaPro → Outils avancés → Droits MySQL")

# ── PUBLIC.MDB ─────────────────────────────────────────────────────
print("\n[3/4] PUBLIC.MDB StudioVision ...")
DEFAULT_MDB = r"C:\StudioVision\PUBLIC.MDB"
mdb_ok = False
mdb_path = DEFAULT_MDB
if not os.path.exists(mdb_path):
    for c in [r"C:\Program Files\StudioVision\PUBLIC.MDB",
              r"C:\Program Files (x86)\StudioVision\PUBLIC.MDB",
              r"D:\StudioVision\PUBLIC.MDB"]:
        if os.path.exists(c):
            mdb_path = c; break

if os.path.exists(mdb_path):
    try:
        import pyodbc
        for drv in ("Microsoft Access Driver (*.mdb, *.accdb)",
                    "Microsoft Access Driver (*.mdb)"):
            try:
                c = pyodbc.connect(f"DRIVER={{{drv}}};DBQ={mdb_path};", autocommit=True)
                cur = c.cursor()
                tables_sv = [t.table_name for t in cur.tables(tableType="TABLE")]
                print(f"  ✓ {mdb_path} — {len(tables_sv)} tables")
                c.close(); mdb_ok = True; break
            except: continue
        if not mdb_ok:
            print("  ✗ Pilote Access manquant")
            print("    https://www.microsoft.com/en-us/download/details.aspx?id=54920")
    except ImportError:
        print("  ✗ pyodbc manquant — pip install pyodbc")
else:
    print(f"  ✗ Introuvable : {mdb_path}")

# ── win32com ───────────────────────────────────────────────────────
print("\n[4/4] win32com (mode --auto) ...", end=" ", flush=True)
try:
    import win32com.client; print("OK"); win32_ok = True
except ImportError:
    print("absent — pip install pywin32"); win32_ok = False

# ── Bilan ──────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("  BILAN")
print("=" * 65)
for label, ok in [("MySQL AlmaPro", mysql_ok), ("PUBLIC.MDB", mdb_ok),
                   ("win32com (--auto)", win32_ok)]:
    print(f"  {'✓' if ok else '✗'}  {label}")

print()
if mysql_ok and mdb_ok:
    print("  ✓ Opérationnel. Testez :")
    print("    python studiovision_vers_almapro_v2.py --test-mysql")
    print("    python studiovision_vers_almapro_v2.py --auto")
else:
    print("  ✗ Corrigez les erreurs ci-dessus puis relancez.")

config = {
    "mysql_host": HOST, "mysql_port": PORT,
    "mdb_path": mdb_path, "mysql_ok": mysql_ok,
    "mdb_ok": mdb_ok, "schemas": schemas,
    "checked_at": __import__("datetime").datetime.now().isoformat(),
}
with open("config_diagnostic.json","w") as f:
    json.dump(config, f, indent=2, default=str)
print(f"\n  Résultats sauvegardés → config_diagnostic.json")
