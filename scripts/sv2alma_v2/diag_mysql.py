"""
diag_mysql.py - diagnostic connexion MySQL AlmaPro
Utilise PyMySQL (compatible MySQL 5.0)

Usage :
    python diag_mysql.py                         # root sans mot de passe
    python diag_mysql.py --password MonMotDePasse
    python diag_mysql.py --try-common            # essaie les mots de passe courants
"""

import sys, subprocess, socket, argparse

HOST = "192.168.0.179"
PORT = 3306
MYSQL_BIN = r"C:\Program Files\MySQL\MySQL Server 5.0\bin"

parser = argparse.ArgumentParser()
parser.add_argument("--password", default=None)
parser.add_argument("--try-common", action="store_true",
                    help="Essaie les mots de passe courants AlmaPro")
args = parser.parse_args()

print("Python :", sys.version)
print("Arch   :", "64-bit" if sys.maxsize > 2**32 else "32-bit")
print()

# -- PyMySQL ------------------------------------------------------------------
print("[1] Import pymysql ...", end=" ", flush=True)
try:
    import pymysql
    print(f"OK - {pymysql.__version__}")
except ImportError:
    print("installation...")
    subprocess.run([sys.executable, "-m", "pip", "install", "pymysql"], check=True)
    import pymysql
    print(f"OK - {pymysql.__version__}")

# -- Ping ---------------------------------------------------------------------
print(f"[2] Ping {HOST} ...", end=" ", flush=True)
r = subprocess.run(["ping","-n","1","-w","2000",HOST], capture_output=True, text=True)
if "TTL=" in r.stdout or "ttl=" in r.stdout:
    print("OK")
else:
    print("ECHEC - VM ne repond pas")
    sys.exit(1)

# -- Port TCP -----------------------------------------------------------------
print(f"[3] Port TCP {HOST}:{PORT} ...", end=" ", flush=True)
try:
    s = socket.create_connection((HOST, PORT), timeout=5); s.close(); print("OK")
except Exception as e:
    print(f"ECHEC : {e}")
    print()
    print("  Pare-feu VM bloque le port 3306.")
    print("  PowerShell admin sur la VM :")
    print("  New-NetFirewallRule -DisplayName 'MySQL' -Direction Inbound")
    print("    -Protocol TCP -LocalPort 3306 -Action Allow")
    sys.exit(1)

# -- Tentative avec mots de passe courants -------------------------------------
def try_connect(host, port, password):
    try:
        conn = pymysql.connect(
            host=host, port=port, user="root", password=password,
            database="almapro", charset="latin1", connect_timeout=5,
        )
        conn.close()
        return True
    except Exception:
        return False

if args.try_common:
    print("[4] Essai des mots de passe courants AlmaPro ...")
    candidates = [
        "",           # vide
        "almapro",    # le plus courant AlmaPro
        "AlmaPro",
        "almapromd",
        "alma",
        "Alma",
        "admin",
        "Admin",
        "root",
        "mysql",
        "MySQL",
        "password",
        "123456",
        "12345",
        "medecin",
        "docteur",
    ]
    found = None
    for pwd in candidates:
        label = f"'{pwd}'" if pwd else "(vide)"
        print(f"  Essai {label:<20}", end=" ", flush=True)
        if try_connect(HOST, PORT, pwd):
            print("FONCTIONNE !")
            found = pwd
            break
        else:
            print("echec")
    if found is not None:
        args.password = found
        print(f"\n  Mot de passe trouve : '{found}'")
        print(f"  Ajoutez dans alma_mysql_writer.py :")
        print(f'  MYSQL_PASSWORD = "{found}"')
    else:
        print("\n  Aucun mot de passe courant ne fonctionne.")
        print("  Utilisez reset_mysql_v3.ps1 sur la VM pour reinitialiser.")
        sys.exit(1)

# -- Connexion principale ------------------------------------------------------
password = args.password if args.password is not None else ""
label_pwd = f"'{password}'" if password else "(vide)"
print(f"[4] Connexion MySQL root@{HOST} password={label_pwd} ...", end=" ", flush=True)
try:
    conn = pymysql.connect(
        host=HOST, port=PORT, user="root", password=password,
        database="almapro", charset="latin1", connect_timeout=10,
    )
    print("OK !")
except pymysql.err.OperationalError as e:
    code, msg = e.args
    print(f"ECHEC [{code}]")
    print()
    if code == 1045:
        print("  Mot de passe incorrect.")
        print()
        print("  OPTIONS :")
        print()
        print("  A) Essayer les mots de passe courants :")
        print("     python diag_mysql.py --try-common")
        print()
        print("  B) Reinitialiser depuis la VM :")
        print("     1. Copier reset_mysql_v3.ps1 sur la VM")
        print("     2. PowerShell admin -> .\\reset_mysql_v3.ps1")
        print("     Attention : le daemon s'appelle mysqld-nt.exe (pas mysqld.exe)")
        print(f"     Chemin : {MYSQL_BIN}\\mysqld-nt.exe")
        print()
        print("  C) Si vous connaissez le mot de passe :")
        print("     python diag_mysql.py --password MOT_DE_PASSE")
    sys.exit(1)

# -- Schema --------------------------------------------------------------------
cur = conn.cursor()

cur.execute("SHOW TABLES")
tables = [r[0] for r in cur.fetchall()]
print(f"[5] Tables : {len(tables)} dans almapro")

# Afficher le GRANT actuel pour root depuis le poste client
print("[6] Verification droits acces depuis poste client ...")
try:
    cur.execute("SELECT Host, User, Password FROM mysql.user WHERE User='root'")
    for row in cur.fetchall():
        host_val, user_val, pwd_val = row
        has_pwd = "OUI" if pwd_val else "NON"
        print(f"  root@{host_val} - mot de passe: {has_pwd}")
except Exception as e:
    print(f"  {e}")

print("[7] DESCRIBE adm_patient :")
if "adm_patient" in tables:
    cur.execute("DESCRIBE adm_patient")
    cols = [(r[0], r[1]) for r in cur.fetchall()]
    for name, typ in cols:
        print(f"       {name:<40} {typ}")

print("[8] DESCRIBE consu_contact :")
if "consu_contact" in tables:
    cur.execute("DESCRIBE consu_contact")
    for r in cur.fetchall():
        print(f"       {r[0]:<40} {r[1]}")

cur.execute("SELECT COUNT(*) FROM adm_patient")
print(f"[9] Patients : {cur.fetchone()[0]}")

cur.close(); conn.close()

print()
print("=" * 60)
print("  TOUT OK - MySQL AlmaPro accessible")
if password:
    print(f"  Mot de passe : '{password}'")
    print()
    print("  Mettez a jour MYSQL_PASSWORD dans alma_mysql_writer.py :")
    print(f'  MYSQL_PASSWORD = "{password}"')
print("=" * 60)
