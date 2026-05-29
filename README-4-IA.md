# StudioVision → AlmaPro → DMP
## Complete Technical Handoff for Next AI Assistant

**Language:** English (code comments bilingual FR/EN)
**Project status:** Partially complete — MySQL access broken, needs AlmaPro reinstall
**Last working state:** Direct MySQL import tested successfully (72 consultations, 42 docs)

---

## 1. WHAT THIS PROJECT IS

An ophthalmology practice in France needs to migrate patient records from **StudioVision**
(an ophthalmology-specific EMR based on Microsoft Access) to **AlmaPro** (a general-practice
French EMR based on WinDev + MySQL), in order to send patient records to the national
**DMP** (Dossier Médical Partagé — French national health record system).

The DMP upload is natively supported by AlmaPro via its built-in module (requires a CPS
smart card reader). AlmaPro is the mandatory bridge between StudioVision and DMP.

```
StudioVision (PUBLIC.MDB) → [our scripts] → AlmaPro (MySQL) → [AlmaPro native] → DMP
```

---

## 2. INFRASTRUCTURE

```
┌─────────────────────────────────────────────────────┐
│  Network: 192.168.0.x                               │
│                                                     │
│  ┌──────────────────────┐  ┌──────────────────────┐ │
│  │  Client PC           │  │  VM Server           │ │
│  │  hostname: Box-      │  │  IP: 192.168.0.179   │ │
│  │  contacto            │  │  hostname: DESKTOP-  │ │
│  │                      │  │  MQKMFAQ             │ │
│  │  - StudioVision      │  │                      │ │
│  │  - AlmaPro client    │  │  - AlmaPro server    │ │
│  │  - Python scripts    │  │  - MySQL 5.0.87-nt   │ │
│  │  - M:\ = network     │  │  - C:\almapro\       │ │
│  │    drive to server   │  │  - ibdata1 (26MB)    │ │
│  └──────────────────────┘  └──────────────────────┘ │
│                                                     │
│  ┌──────────────────────┐                           │
│  │  Secretary PC        │                           │
│  │  - AlmaPro client    │                           │
│  │  - CPS card reader   │ ← DMP upload happens here│
│  └──────────────────────┘                           │
└─────────────────────────────────────────────────────┘
```

**Key paths:**
```
CLIENT:
  M:\fichier\PUBLIC.MDB         ← StudioVision database (Access MDB)
  M:\PHOTOS\                    ← Patient image/document folders

SERVER VM:
  C:\almapro\                   ← AlmaPro application root
  C:\almapro\AlmaProV4.exe      ← AlmaPro executable
  C:\almapro\config.dat         ← WinDev encrypted config (contains MySQL password)
  C:\almapro\imageexterne\      ← Patient external documents storage
  C:\Program Files\MySQL\MySQL Server 5.0\bin\mysqld-nt.exe  ← MySQL daemon
  C:\Program Files\MySQL\MySQL Server 5.0\data\ibdata1       ← InnoDB tablespace
  \\DESKTOP-MQKMFAQ\Users\Public\Documents\import\           ← XML import folder
```

---

## 3. STUDIOVISION DATABASE SCHEMA (PUBLIC.MDB)

StudioVision stores data in a Microsoft Access .MDB file accessible via ODBC.
Source of truth: OphthaView project on GitHub (Vic-Warden/OphthaView).

### Table: `Patients`
```sql
[Code patient]     INTEGER  PRIMARY KEY   -- unique patient code
NOM                TEXT(50)               -- last name UPPERCASE
Prénom             TEXT(50)               -- first name
Date de naissance  DATE
Sexe               TEXT(1)               -- M/F
NumSecu            TEXT(21)              -- French NIR
Adresse            TEXT(100)
Ville              TEXT(50)
CodePostal         TEXT(10)
Téléphone          TEXT(20)
Email              TEXT(50)
Antécédants        MEMO                  -- medical history (NOTE: typo in original)
Allergies          MEMO
Traitements        MEMO
Diagnostic OPH     MEMO                  -- ophthalmological diagnosis
Important          MEMO                  -- important notes
```

### Table: `Consultation`
```sql
[Code patient]         INTEGER   FK → Patients
[N° consultation]      INTEGER   consultation number
Date                   DATE
DOMINANTE              TEXT      chief complaint / visit reason
REFRACTION             MEMO      free-text refraction
AVscOD, AVscOG         TEXT      uncorrected visual acuity OD/OG
AVccOD, AVccOG         TEXT      corrected visual acuity OD/OG
TOD, TOG               TEXT      intraocular pressure OD/OG
Ordonnance             MEMO      prescription text
AutresPrescriptions    MEMO      other prescriptions
```

### Table: `tREFRACTION` (structured refraction measurements)
```sql
NumConsult   INTEGER   FK → Consultation.[N° consultation]
TypeRef      TEXT      "1"=subj far, "2"=subj near, "4"=cycloplegia,
                       "5"=autoref, "12"=worn glasses, "13"=worn lenses
SphOD        FLOAT     sphere right eye
CylOD        FLOAT     cylinder right eye
AxeOD        INTEGER   axis right eye
AddOD        FLOAT     addition right eye
SphOG        FLOAT     sphere left eye
CylOG        FLOAT     cylinder left eye
AxeOG        INTEGER   axis left eye
AddOG        FLOAT     addition left eye
AVDL         TEXT      distance VA right
AVGL         TEXT      distance VA left
AVDP         TEXT      near VA right
AVGP         TEXT      near VA left
```

### Table: `tKERATO` (keratometry)
```sql
NumConsult   INTEGER   FK → Consultation.[N° consultation]
K1OD         FLOAT     flat meridian right eye
K2OD         FLOAT     steep meridian right eye
KmOD         FLOAT     mean keratometry right eye
AxeOD        INTEGER   axis right eye
K1OG, K2OG, KmOG, AxeOG  -- same for left eye
```

### Table: `Documents`
```sql
[code patient]    INTEGER   FK → Patients
[Photo externe]   TEXT      relative path: \group\folder\filename.ext
Date              DATE
DESCRIPTIONS      TEXT
TEXTE             MEMO
TypeVW            TEXT      document type
```

**Patient folder resolution (critical — same as Studiovision-Autosync):**
```python
# Query to find patient folder:
rows = conn.execute(
    "SELECT TOP 1 [Photo externe] FROM Documents "
    "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL",
    (patient_code,)
)
# Result: "\05.000\5182megr.leo\filename.tif"
# Split on / or \ → parts = ["05.000", "5182megr.leo", "filename.tif"]
# Patient folder = photos_root / parts[0] / parts[1]
# Example: M:\PHOTOS\05.000\5182megr.leo\
```

### Reading the active patient via COM (Studiovision-Autosync method)
```python
import win32com.client

def get_active_patient():
    """Read patient fields from the currently open StudioVision form."""
    access = win32com.client.GetActiveObject("Access.Application")
    form = access.Screen.ActiveForm
    data = {}
    for i in range(form.Controls.Count):
        try:
            ctrl = form.Controls(i)
            name = str(ctrl.Name)
            if name in ("Code patient", "NOM", "Prénom"):
                data[name] = ctrl.Value
        except Exception:
            pass
    if "Code patient" in data:
        return {
            "code": str(data["Code patient"]).strip(),
            "nom": str(data["NOM"]).strip().upper(),
            "prenom": str(data["Prénom"]).strip()
        }
    return None
```

---

## 4. ALMAPRO DATABASE SCHEMA (MySQL 5.0)

### Critical facts about MySQL setup
- **Daemon:** `mysqld-nt.exe` (NOT `mysqld.exe`) — Windows NT variant of MySQL 5.0
- **Port:** 3306
- **Charset:** `latin1` — ALL queries and connections MUST use latin1
- **Engine:** InnoDB (shared tablespace in ibdata1)
- **Python driver:** Use **PyMySQL** — mysql-connector-python 9.x is INCOMPATIBLE with MySQL 5.0
- **Password:** Stored encrypted in `C:\almapro\config.dat` (WinDev proprietary format)

```python
# CORRECT connection:
import pymysql
conn = pymysql.connect(
    host="192.168.0.179", port=3306,
    user="root", password="RECOVERED_PASSWORD",
    database="almapro", charset="latin1",
    connect_timeout=10
)

# WRONG — will crash silently on MySQL 5.0:
import mysql.connector  # DO NOT USE — incompatible with MySQL 5.0
```

### Table: `adm_patient` (confirmed from almapro.db SQLite schema)
```sql
NUMEROPATIENT    INTEGER    PRIMARY KEY AUTO_INCREMENT
NOM              VARCHAR(50)
PRENOM           VARCHAR(50)
DATENAISS        DATE       NOT NULL   -- ⚠ NO DEFAULT — insert '1900-01-01' if unknown
SEXE             VARCHAR(1)            -- 'M' or 'F'
NUMASSURERO      INTEGER    DEFAULT 0
```

**⚠ CRITICAL:** `DATENAISS` has no default value. Any INSERT without it will fail:
```
(1364, "Field 'DATENAISS' doesn't have a default value")
```

### Table: `adm_administratif`
```sql
NumeroPatient        INTEGER    FK → adm_patient.NUMEROPATIENT
ADRESSEL1            VARCHAR(255)
ADRESSEL2            VARCHAR(255)
VILLE                VARCHAR(35)
CODEPOSTAL           VARCHAR(12)
Telephone            VARCHAR(15)
TelephonePortable    VARCHAR(15)
TelephonePro         VARCHAR(15)
FAX                  VARCHAR(15)
Email                VARCHAR(50)
CodeMedecin          INTEGER              -- internal doctor number
Commentaire          VARCHAR(255)
```

### Table: `consu_contact` (consultations)
```sql
NumeroContact        INTEGER    PRIMARY KEY AUTO_INCREMENT
NUMEROPATIENT        INTEGER    FK → adm_patient.NUMEROPATIENT
NumeroMedecin        INTEGER              -- internal doctor number
InitialesMedecin     VARCHAR(5)           -- e.g. "OM"
NOM                  VARCHAR(50)          -- consultation title
Rtf_Cat1             LONGTEXT             -- consultation text in RTF format ⚠
Efface               TINYINT  DEFAULT 0   -- soft delete flag
Date                 DATE
```

**⚠ RTF format required:** AlmaPro stores consultation text as RTF. Plain text will display
but may cause formatting issues. Minimal RTF:
```python
def to_rtf(text: str) -> str:
    escaped = text.replace("\\","\\\\").replace("{","\\{").replace("}","\\}")
    parts = []
    for ch in escaped:
        code = ord(ch)
        if code > 127:
            parts.append(f"\\'{code:02x}")
        elif ch == "\n":
            parts.append("\\par\r\n")
        else:
            parts.append(ch)
    return (
        r"{\rtf1\ansi\ansicpg1252\deff0"
        r"{\fonttbl{\f0\fswiss\fcharset0 Arial;}}"
        r"{\colortbl;}\f0\fs20 "
        + "".join(parts) + "}"
    )
```

### Doctor information (from ibdata1 analysis)
```python
DEFAULT_RPPS             = "7817136230"
DEFAULT_MEDECIN_NOM      = "MEGRET"
DEFAULT_MEDECIN_PRENOM   = "Olivier"
DEFAULT_MEDECIN_INITIALES = "OM"
DEFAULT_MEDECIN_NUM      = 1   # internal AlmaPro doctor number
```

---

## 5. IBDATA1 BINARY ANALYSIS

`ibdata1` is the InnoDB shared tablespace containing ALL AlmaPro patient data.
MySQL 5.0 with InnoDB uses a shared tablespace architecture.

### InnoDB DATE encoding
```python
INNODB_DATE_XOR = 0x800000

def decode_innodb_date(b1: int, b2: int, b3: int) -> str | None:
    """Decode 3 InnoDB bytes → 'dd/mm/yyyy' string."""
    val   = ((b1 << 16) | (b2 << 8) | b3) ^ INNODB_DATE_XOR
    day   =  val & 0x1F
    month = (val >> 5) & 0x0F
    year  =  val >> 9
    if 1900 < year < 2100 and 1 <= month <= 12 and 1 <= day <= 31:
        try:
            from datetime import date
            date(year, month, day)
            return f"{day:02d}/{month:02d}/{year}"
        except ValueError:
            pass
    return None
```

### Patient record structure in ibdata1
```
Offset +0   : NOM      VARCHAR(50) padded with spaces  → 50 bytes
Offset +50  : PRENOM   VARCHAR(50) padded with spaces  → 50 bytes (may be binary/empty)
Offset +100 : DATENAISS DATE (3 bytes InnoDB encoded)
...
```

### Finding patients by name + DDN
```python
def find_patient(data: bytes, nom: str, ddn: str) -> int | None:
    """Search ibdata1 for patient by name + birth date."""
    nom_key = (nom.upper() + " " * 50)[:50].encode("latin-1")
    start = 0
    while True:
        pos = data.find(nom_key, start)
        if pos == -1:
            break
        chunk = data[pos:pos+150]
        dates = find_dates_in_chunk(chunk)
        if ddn in dates:
            return pos  # found at this offset
        start = pos + 1
    return None
```

### Key offsets found during analysis
```
Offset 3,703,347  → DEMO PATIENTE (DDN 14/07/1932)
Offset 3,709,137  → TEST Test (DDN 23/11/1990)
Offset 4,146,984  → TEST Test (another occurrence)
Journal audit     → Contains access logs with date+doctor+patient
```

---

## 6. AlmaPro XML IMPORT FORMAT

This is the "import from other software" format, NOT the native AlmaPro-to-AlmaPro format.

**Spec document:** `C:\almapro\Importer des données d'un autre logiciel à partir d'un fichier XML compatible AlmaPro.pdf`

**Critical rules:**
- Encoding: `iso-8859-1` — MANDATORY, not UTF-8
- Dates: `dd/mm/yyyy` — MANDATORY
- Tag order: STRICT — must follow the spec exactly
- File naming: `1.xml`, `2.xml`, etc. (numeric)
- Documents: place in `docext/` subfolder alongside the XML

**⚠ IMPORTANT LIMITATION:** This import ALWAYS creates a NEW patient record.
It does NOT merge with existing records. Using it on an already-imported patient
creates duplicates. The only way to avoid duplicates is direct MySQL INSERT.

```xml
<?xml version="1.0" encoding="iso-8859-1" standalone="yes"?>
<Export_XML_AlmaPro>
  <Dossiers>
    <Nom>DUPONT</Nom>
    <Prenom>Marie</Prenom>
    <Sexe>F</Sexe>
    <Adresse>22 rue des Lilas</Adresse>
    <AdresseL2></AdresseL2>
    <Ville>Paris</Ville>
    <CodePostal>75001</CodePostal>
    <Tel_Domicile>0123456789</Tel_Domicile>
    <Tel_Pro></Tel_Pro>
    <Tel_Port></Tel_Port>
    <FAX></FAX>
    <EMail>patient@example.com</EMail>
    <Nom_Jeune_Fille></Nom_Jeune_Fille>
    <Date_de_naissance>12/03/1965</Date_de_naissance>
    <Date_Contrat_Signe></Date_Contrat_Signe>
    <ContratSigne>FALSE</ContratSigne>
    <Decede>FALSE</Decede>
    <Profession_du_patient>Infirmière</Profession_du_patient>
    <Num_Secu>2650312075123</Num_Secu>
    <Remarque>Imported from StudioVision 12/06/2026</Remarque>
  </Dossiers>

  <Dossier_Consultations>
    <Date_Contact>15/03/2024</Date_Contact>
    <Titre_Contact>Consultation ophtalmologique</Titre_Contact>
    <Texte_Contact>Motif : renouvellement lunettes
AV sc OD=10/10  OG=10/10
Refraction : OD +0.50 (-0.25 x 90)  OG +0.75 plano</Texte_Contact>
  </Dossier_Consultations>

  <ATCD>
    <Texte_ATCD>Myopie légère depuis l'enfance</Texte_ATCD>
  </ATCD>

  <Allergie>
    <Texte_Allergie>Pénicilline</Texte_Allergie>
  </Allergie>

  <DocExt>
    <Nom_DocExt>OCT - macula_OD_20240315</Nom_DocExt>
    <NomFichier_DocExt>macula_OD_20240315.tif</NomFichier_DocExt>
    <Emis_DocExt>FALSE</Emis_DocExt>
    <Date_DocExt>15/03/2024</Date_DocExt>
  </DocExt>
</Export_XML_AlmaPro>
```

---

## 7. WHAT WORKED — SUCCESSFUL PIPELINE

The following complete pipeline was tested successfully on patient MEGRET Leo:

```
[StudioVision open on client PC]
         ↓ COM (win32com)
get_active_patient() → {code: "5182", nom: "MEGRET", prenom: "Leo"}
         ↓ ODBC (pyodbc → M:\fichier\PUBLIC.MDB)
get_patient_by_code("5182") → full patient dict
get_patient_consultations("5182") → 72 consultations with refraction/keratometry
get_patient_folder_path("5182") → M:\PHOTOS\05.000\5182megr.leo\
list_patient_files(folder) → 42 files (jpg, tif, pdf)
         ↓ PyMySQL (pymysql → 192.168.0.179:3306)
AlmaWriter.connect()
AlmaWriter.insert_patient() →
    INSERT INTO adm_patient (NUMEROPATIENT=19, NOM="MEGRET", PRENOM="Leo",
                              DATENAISS="1900-01-01", SEXE="M", NUMASSURERO=0)
    INSERT INTO adm_administratif (NumeroPatient=19, ...)
    INSERT INTO consu_contact × 72
         ↓ SMB (os.makedirs + shutil.copy2)
\\192.168.0.179\almapro\imageexterne\REP_19\ ← 42 documents copied
```

**Output:**
```
[3/4] Insertion dans MySQL AlmaPro (192.168.0.179)...
  - Connect- - MySQL 192.168.0.179:3306/almapro
  - Numéro de dossier : #19
  - adm_patient : MEGRET Leo (#19)
  - adm_administratif : OK
  - consu_contact : 72 consultation(s)
[4/4] Copie des documents...
  42 document(s) -> \\192.168.0.179\almapro\imageexterne\REP_19
IMPORT TERMINE - MEGRET Leo — Dossier : #19
```

---

## 8. WHAT FAILED — ERRORS AND LESSONS

### ERROR 1: mysql-connector-python 9.x incompatible with MySQL 5.0
**Symptom:** Python process crashes silently at `mysql.connector.connect()`
**Root cause:** mysql-connector-python 9.x uses `caching_sha2_password` protocol
introduced in MySQL 8.0. MySQL 5.0 uses `mysql_native_password`. The C extension
crashes without error message.

**Wrong:**
```python
import mysql.connector
conn = mysql.connector.connect(host="192.168.0.179", user="root", ...)
# CRASH — no error message, process just dies
```

**Correct:**
```python
import pymysql  # pure Python, compatible with all MySQL versions
conn = pymysql.connect(host="192.168.0.179", user="root", charset="latin1", ...)
```

**Lesson:** Always use PyMySQL for MySQL 5.0. Never use mysql-connector-python.

---

### ERROR 2: mysqld-nt.exe vs mysqld.exe
**Symptom:** `Start-Process mysqld.exe` fails — file not found
**Root cause:** MySQL 5.0 on Windows NT uses `mysqld-nt.exe` (NT variant).
The file `mysqld.exe` does NOT EXIST.

**Wrong:**
```powershell
Start-Process "C:\Program Files\MySQL\MySQL Server 5.0\bin\mysqld.exe"
# Error: file not found
```

**Correct:**
```powershell
Start-Process "C:\Program Files\MySQL\MySQL Server 5.0\bin\mysqld-nt.exe" --skip-grant-tables
```

---

### ERROR 3: DATENAISS NOT NULL without default
**Symptom:** `(1364, "Field 'DATENAISS' doesn't have a default value")`
**Root cause:** The `DATENAISS` column in `adm_patient` is NOT NULL with no DEFAULT value.
Many StudioVision patients have no birth date (field left empty).

**Wrong:**
```python
patient_row = {
    "NOM": "MEGRET",
    "PRENOM": "Leo",
    # DATENAISS missing → MySQL error
}
```

**Correct:**
```python
def fmt_date(value) -> str:
    """Returns '1900-01-01' if date is unknown/empty."""
    if not value:
        return "1900-01-01"  # placeholder for unknown DDN
    # ... normal date parsing ...

patient_row = {
    "NOM": "MEGRET",
    "PRENOM": "Leo",
    "DATENAISS": fmt_date(patient.get("ddn", "")),  # always set
}
```

---

### ERROR 4: --skip-grant-tables + --skip-networking = client can't connect
**Symptom:** `ERROR 2003 (HY000): Can't connect to MySQL server on 'localhost' (10061)`
**Root cause:** `--skip-networking` disables ALL TCP connections including localhost.
`mysql.exe` on Windows uses TCP by default, not Unix sockets (Windows has named pipes
but they're not the default).

**Wrong:**
```powershell
# Start mysqld-nt with both flags
mysqld-nt.exe --skip-grant-tables --skip-networking
# Then try to connect via mysql.exe → ERROR 2003
mysql.exe -u root -e "UPDATE mysql.user SET Password=''..."
```

**Correct:**
```powershell
# Use --skip-grant-tables WITHOUT --skip-networking
# Window 1:
mysqld-nt.exe --skip-grant-tables
# Window 2 (immediately):
mysql.exe -u root -e "UPDATE mysql.user SET Password='' WHERE User='root'; FLUSH PRIVILEGES;"
```

---

### ERROR 5: Resetting MySQL password without knowing AlmaPro's password first
**Symptom:** AlmaPro can't open: `Access denied for user 'root'@'localhost' (using password: YES)`
**Root cause:** AlmaPro stores its MySQL password in `config.dat` (WinDev encrypted binary).
When we reset the MySQL password to empty, AlmaPro kept sending its stored password,
which no longer matches.

**What we did wrong:**
```
1. Reset MySQL root password to empty ← did this without knowing what AlmaPro uses
2. AlmaPro sends password: YES (from config.dat)
3. MySQL expects empty password
4. AlmaPro cannot connect → unusable
```

**What should have been done:**
```
1. FIRST: find AlmaPro's MySQL password from config.dat or mysql.user hash
   (BEFORE touching anything)
2. THEN: use that password in Python scripts
3. NEVER reset MySQL password unless you can reconfigure AlmaPro too
```

**How to get the password SAFELY (for next attempt after reinstall):**
```powershell
# Immediately after AlmaPro installation, before anything else:
& "C:\Program Files\MySQL\MySQL Server 5.0\bin\mysql.exe" `
    -u root -p `
    -e "SELECT Host, User, Password FROM mysql.user WHERE User='root';"
# This shows the password hash. Note it.

# Also check if AlmaPro grants access from client:
# AlmaPro → Tools → Advanced Tools → "Droits d'accès MySQL"
# This runs: GRANT ALL ON almapro.* TO 'root'@'%' IDENTIFIED BY 'password'
# The actual SQL executed can be seen in mysql.log if logging is enabled
```

---

### ERROR 6: config.dat is WinDev encrypted binary
**Symptom:** Can't read MySQL password from config.dat
**Root cause:** AlmaPro is built with WinDev (PCSoft framework). WinDev stores
configuration in `.dat` files using a proprietary encryption scheme.

```
config.dat (53,248 bytes) → WinDev proprietary encrypted format
                           → CANNOT be read without WinDev decryption key
                           → mysql.user.MYD would have had the hash, but
                              our reset already overwrote it
```

**Lesson:** The MySQL password set by AlmaPro installer is ONLY accessible in:
1. `mysql.user` table (Password column — hash format)
2. AlmaPro's `config.dat` (encrypted — inaccessible)
3. Memory of running AlmaPro process

**Only safe approach:** capture it from `mysql.user` immediately after install.

---

### ERROR 7: XML import creates duplicates
**Symptom:** Importing XML for a patient who already exists creates a second record
**Root cause:** AlmaPro's XML import module ALWAYS creates a new patient record.
It has no duplicate detection.

**Wrong workflow:**
```
1. Import patient via XML → creates dossier #13
2. Import same patient again → creates dossier #17 (DUPLICATE)
```

**Correct approach:**
- Use direct MySQL INSERT with duplicate check (our v2 approach)
- Check before import: `SELECT COUNT(*) FROM adm_patient WHERE NOM=? AND PRENOM=? AND DATENAISS=?`

---

## 9. CURRENT STATE AND WHAT TO DO NEXT

### Current state (as of 29/05/2026)
```
✅ StudioVision reading works perfectly (72 consultations, 42 docs for test patient)
✅ MySQL direct INSERT works (tested — dossier #19 created successfully)
✅ Document copy to \\server\almapro\imageexterne\REP_19\ works
❌ AlmaPro is broken — MySQL password mismatch after our reset
❌ Needs full reinstall of AlmaPro + MySQL on VM
```

### Step 1: Reinstall AlmaPro on VM
1. Uninstall AlmaPro: `C:\almapro\uninstall.exe`
2. Delete residual folders: `C:\almapro\`, `C:\Program Files\MySQL\`
3. Reinstall AlmaPro from original installer
4. **IMMEDIATELY after install**, capture MySQL credentials:

```powershell
# Run this BEFORE anything else after install
$mysql = "C:\Program Files\MySQL\MySQL Server 5.0\bin\mysql.exe"

# Get the password hash AlmaPro set
& $mysql -u root -p -e "SELECT Host, User, Password FROM mysql.user;"

# Also try connecting without password first (maybe AlmaPro installs with empty pwd)
& $mysql -u root -e "SELECT 'connected' AS test;"
```

5. If AlmaPro installs MySQL with a password, try:
```powershell
# Enable query log to capture AlmaPro's connection
Add-Content "C:\Program Files\MySQL\MySQL Server 5.0\my.ini" "`nlog=C:\mysql_queries.log"
net stop MySQL; net start MySQL
# Open AlmaPro (it will connect successfully at this point)
# Read the log:
Get-Content "C:\mysql_queries.log" | Select-String "Connect|Password|root"
```

### Step 2: Update Python scripts with new password
```python
# In scripts/sv2alma_v2/alma_mysql_writer.py
MYSQL_PASSWORD = "DISCOVERED_PASSWORD"  # replace this
```

### Step 3: Open firewall port 3306
```powershell
New-NetFirewallRule -DisplayName "MySQL AlmaPro 3306" `
    -Direction Inbound -Protocol TCP -LocalPort 3306 -Action Allow
```

### Step 4: Test and run
```powershell
# On client PC
cd C:\sv2alma_v2
python diag_mysql.py --try-common
python studiovision_vers_almapro_v2.py --auto
```

---

## 10. FILES STRUCTURE

```
projet_final/
├── README-4-tech.md              ← French technical README (for IT team)
├── README-4-IA.md                ← This file (for next AI assistant)
│
├── scripts/
│   ├── xml_import/
│   │   └── generer_xml_import_almapro.py   ← XML from ibdata1 (standalone tool)
│   │
│   ├── sv2alma_v1/               ← V1: XML generation pipeline
│   │   ├── sv_reader.py          ← StudioVision reader (MDB + COM)
│   │   ├── alma_xml.py           ← AlmaPro XML generator
│   │   └── studiovision_vers_almapro.py    ← Main script v1
│   │
│   ├── sv2alma_v2/               ← V2: Direct MySQL pipeline ✅ WORKING
│   │   ├── sv_reader.py          ← StudioVision reader (same as v1)
│   │   ├── alma_xml.py           ← AlmaPro XML (fallback)
│   │   ├── alma_mysql_writer.py  ← Direct MySQL INSERT ← MAIN NEW PIECE
│   │   ├── studiovision_vers_almapro_v2.py ← Main script v2
│   │   ├── diag_mysql.py         ← MySQL connection diagnostic
│   │   └── setup_et_diagnostic.py ← Full environment check
│   │
│   ├── exe_gui/                  ← GUI executable for doctors
│   │   ├── sv_export.py          ← Tkinter GUI (single button interface)
│   │   └── sv_export.spec        ← PyInstaller spec (builds .exe)
│   │
│   └── mysql_tools/
│       └── reset_mysql_v3.ps1    ← MySQL password reset (PowerShell admin)
│
└── icons/
    ├── export_sv_alma.ico         ← Composite icon: SV left / AlmaPro right
    └── studiovision_extension.ico ← SV icon with blue badge (extension marker)
```

---

## 11. DEPENDENCIES AND INSTALLATION

```powershell
# Python (Windows, 64-bit, 3.11+)
pip install pyodbc pywin32 pymysql pyinstaller Pillow

# Microsoft Access Database Engine 64-bit (for PUBLIC.MDB access)
# Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920
# IMPORTANT: Must match Python bitness (64-bit Python → 64-bit ACE)
```

```
MySQL driver comparison:
┌─────────────────────────────┬─────────────────────────────┐
│ mysql-connector-python 9.x  │ PyMySQL 1.x+                │
├─────────────────────────────┼─────────────────────────────┤
│ ❌ Incompatible MySQL 5.0   │ ✅ Compatible MySQL 5.0+    │
│ ❌ Crashes silently         │ ✅ Proper error messages     │
│ C extension (fast)          │ Pure Python (slower, safer) │
└─────────────────────────────┴─────────────────────────────┘
USE PYMYSQL.
```

---

## 12. RECOMMENDED ARCHITECTURE FOR NEXT SESSION

```
GOAL: StudioVision (patient open) → AlmaPro MySQL → DMP

PREREQUISITES (from next AI session):
1. AlmaPro freshly reinstalled on VM
2. MySQL password captured and stored in alma_mysql_writer.py
3. Port 3306 open on VM firewall
4. diag_mysql.py confirms connection

IMPLEMENTATION ORDER:
1. Test diag_mysql.py → confirm MySQL accessible
2. Run studiovision_vers_almapro_v2.py --auto with test patient
3. Open AlmaPro → verify patient dossier appears correctly
4. Test DMP upload from AlmaPro (requires CPS card on secretary PC)

FUTURE IMPROVEMENTS (not yet implemented):
- GUI executable: sv_export.py → compile with PyInstaller → desktop shortcut
- Batch import: loop over all StudioVision patients
- Duplicate detection: check AlmaPro before insert
- Log file: trace all imports for RGPD compliance
- Auto-start: run on StudioVision patient change event (COM event listener)
```

---

## 13. QUICK REFERENCE — COMMANDS

```powershell
# ── CLIENT PC ────────────────────────────────────────────────────────

# Test everything
python setup_et_diagnostic.py

# Test MySQL only
python diag_mysql.py
python diag_mysql.py --try-common          # try common passwords
python diag_mysql.py --password "MYPASSWD" # test specific password

# Import current patient (StudioVision must be open on patient)
python studiovision_vers_almapro_v2.py --auto

# Import by name
python studiovision_vers_almapro_v2.py --nom DUPONT --prenom Marie

# Import by StudioVision code
python studiovision_vers_almapro_v2.py --code 5182

# List last 20 patients in AlmaPro
python studiovision_vers_almapro_v2.py --list

# Show AlmaPro table schema
python studiovision_vers_almapro_v2.py --schema adm_patient

# Build GUI executable
pyinstaller sv_export.spec
# → dist\ExportStudioVision_AlmaPro.exe

# ── SERVER VM (PowerShell admin) ──────────────────────────────────────

# Open MySQL port
New-NetFirewallRule -DisplayName "MySQL" -Direction Inbound -Protocol TCP -LocalPort 3306 -Action Allow

# Check MySQL users and passwords
& "C:\Program Files\MySQL\MySQL Server 5.0\bin\mysql.exe" -u root -e "SELECT Host,User,Password FROM mysql.user;"

# Grant access from client
& "C:\Program Files\MySQL\MySQL Server 5.0\bin\mysql.exe" -u root -e "GRANT ALL ON almapro.* TO 'root'@'%' IDENTIFIED BY 'PASSWORD'; FLUSH PRIVILEGES;"

# Reset MySQL password (if locked out — SEE ERROR 4 before using this)
.\reset_mysql_v3.ps1
```

---

*Generated 29/05/2026 — Project: StudioVision→AlmaPro→DMP — Cabinet ophtalmologie*
