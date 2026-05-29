# Projet StudioVision → AlmaPro → DMP
## Rapport technique — Cabinet d'ophtalmologie

---

## Contexte

Migration d'un cabinet médical d'ophtalmologie du logiciel **StudioVision** vers **AlmaPro**,
avec comme objectif final l'envoi des dossiers patients sur le **DMP** (Dossier Médical Partagé).

**Infrastructure réseau :**

| Machine | Rôle | IP |
|---------|------|----|
| VM Windows Server | Serveur AlmaPro + MySQL 5.0 | 192.168.0.179 |
| Poste libre | Client AlmaPro + StudioVision | Box-contacto |
| Poste secrétariat | Client AlmaPro + lecteur CPS | — |

---

## Architecture des données découverte

### StudioVision — PUBLIC.MDB
Base Access (ODBC), chemin réseau : `M:\fichier\PUBLIC.MDB`

| Table | Contenu |
|-------|---------|
| `Patients` | Identité complète (NOM, Prénom, Date de naissance, NIR, adresse...) |
| `Consultation` | 1 ligne par consultation avec DOMINANTE, REFRACTION, AVsc, TOD/TOG, Ordonnance |
| `tREFRACTION` | Mesures structurées (Sph, Cyl, Axe, Add OD/OG par type) liées par NumConsult |
| `tKERATO` | Kératométrie (K1, K2, Km, Axe OD/OG) liées par NumConsult |
| `Documents` | Chemin des photos/OCT/PDF via champ `Photo externe` |

**Dossier photos :** `M:\PHOTOS\<groupe>\<code_patient><nom>.<prenom>\`

**Lecture de la fiche ouverte :** via COM Windows — `win32com.client.GetActiveObject("Access.Application")`

### AlmaPro — MySQL 5.0
Daemon : `mysqld-nt.exe` (pas `mysqld.exe`)
Chemin : `C:\Program Files\MySQL\MySQL Server 5.0\bin\`
Base : `almapro` — charset `latin1`

**Tables principales :**

| Table | Colonnes clés |
|-------|--------------|
| `adm_patient` | NUMEROPATIENT (PK), NOM, PRENOM, DATENAISS (NOT NULL), SEXE, NUMASSURERO |
| `adm_administratif` | NumeroPatient, ADRESSEL1, VILLE, CODEPOSTAL, Telephone, Email, CodeMedecin |
| `consu_contact` | NUMEROPATIENT, NumeroMedecin, InitialesMedecin, NOM, Rtf_Cat1, Date |

**⚠ Pièges critiques :**
- `DATENAISS` est NOT NULL sans valeur par défaut → insérer `1900-01-01` si DDN inconnue
- Les consultations sont en **format RTF** dans le champ `Rtf_Cat1`
- Le daemon s'appelle `mysqld-nt.exe`, pas `mysqld.exe`

### Format XML AlmaPro (import "autre logiciel")
Encodage : `iso-8859-1` obligatoire
Dates : `jj/mm/aaaa` obligatoire
Spec : `C:\almapro\Importer des données d'un autre logiciel...pdf`

```xml
<?xml version="1.0" encoding="iso-8859-1" standalone="yes"?>
<Export_XML_AlmaPro>
  <Dossiers>
    <Nom>NOM</Nom><Prenom>PRENOM</Prenom><Sexe>M</Sexe>
    <Date_de_naissance>jj/mm/aaaa</Date_de_naissance>
    ...
  </Dossiers>
  <Dossier_Consultations>
    <Date_Contact>jj/mm/aaaa</Date_Contact>
    <Titre_Contact>...</Titre_Contact>
    <Texte_Contact>...</Texte_Contact>
  </Dossier_Consultations>
  <DocExt>
    <Nom_DocExt>...</Nom_DocExt>
    <NomFichier_DocExt>fichier.jpg</NomFichier_DocExt>
    <Emis_DocExt>FALSE</Emis_DocExt>
    <Date_DocExt>jj/mm/aaaa</Date_DocExt>
  </DocExt>
</Export_XML_AlmaPro>
```

---

## Scripts produits

### 1. `scripts/xml_import/generer_xml_import_almapro.py`
Génère un XML AlmaPro depuis ibdata1 (lecture binaire InnoDB).
Utilisé pour tester le format XML avant d'avoir accès à MySQL.

### 2. `scripts/sv2alma_v1/` — Export XML (sans MySQL)
Pipeline complet StudioVision → XML → dépôt sur partage réseau.
**Import manuel ensuite requis dans AlmaPro** (5 étapes interface graphique).

| Fichier | Rôle |
|---------|------|
| `sv_reader.py` | Lecture PUBLIC.MDB + COM |
| `alma_xml.py` | Génération XML format AlmaPro |
| `studiovision_vers_almapro.py` | Orchestration + CLI |

### 3. `scripts/sv2alma_v2/` — Import MySQL direct ✅ FONCTIONNEL
Pipeline complet StudioVision → MySQL AlmaPro **sans aucune étape manuelle**.

```bash
python studiovision_vers_almapro_v2.py --auto
```

**Résultat testé :** MEGRET Leo — 72 consultations + 42 documents importés.

| Fichier | Rôle |
|---------|------|
| `sv_reader.py` | Lecture PUBLIC.MDB + COM |
| `alma_xml.py` | Génération XML (fallback) |
| `alma_mysql_writer.py` | Écriture directe MySQL |
| `studiovision_vers_almapro_v2.py` | Orchestration + CLI |
| `diag_mysql.py` | Diagnostic connexion MySQL |
| `setup_et_diagnostic.py` | Vérification complète environnement |

### 4. `scripts/exe_gui/sv_export.py` — Interface graphique
Exécutable Windows avec interface simple pour le médecin.
Détecte le patient ouvert dans StudioVision, exporte en un clic.

**Compilation :**
```powershell
pip install pyinstaller pyodbc pywin32
pyinstaller sv_export.spec
# → dist\ExportStudioVision_AlmaPro.exe
```

### 5. `scripts/mysql_tools/reset_mysql_v3.ps1`
Reset du mot de passe root MySQL sur la VM (PowerShell admin).

---

## État actuel et prochaines étapes

### Ce qui fonctionne ✅
- Lecture automatique de la fiche patient ouverte dans StudioVision
- Extraction complète : 72 consultations, réfractions, kératométries, 42 documents
- Écriture directe dans MySQL AlmaPro (testé avec succès)
- Copie des documents vers le partage réseau AlmaPro

### Problème actuel ⚠
**AlmaPro ne démarre plus** suite au reset du mot de passe MySQL.
**Solution :** réinstaller AlmaPro + MySQL sur la VM (données perdues volontairement).

### Après réinstallation — actions immédiates

1. **Récupérer le nouveau mot de passe MySQL** (avant toute manipulation) :
   ```powershell
   # Sur la VM, lire le log MySQL dès l'installation
   Get-Content "C:\Program Files\MySQL\MySQL Server 5.0\data\*.err" | Select-String "password"
   ```
   Ou depuis AlmaPro ouvert : Outils → Outils avancés → Droits d'accès MySQL → noter le résultat du SELECT mysql.user

2. **Mettre à jour `alma_mysql_writer.py`** :
   ```python
   MYSQL_PASSWORD = "MOT_DE_PASSE_TROUVE"
   ```

3. **Ouvrir le port 3306** sur la VM (si refermé après réinstallation) :
   ```powershell
   New-NetFirewallRule -DisplayName "MySQL 3306" -Direction Inbound -Protocol TCP -LocalPort 3306 -Action Allow
   ```

4. **Tester :**
   ```powershell
   python diag_mysql.py --try-common
   ```

---

## Prérequis techniques

### Poste client (StudioVision)
```powershell
pip install pyodbc pywin32 pymysql
# + Microsoft Access Database Engine 64 bits
# https://www.microsoft.com/en-us/download/details.aspx?id=54920
```

### Chemins configurés
```python
DEFAULT_MDB    = r"M:\fichier\PUBLIC.MDB"
DEFAULT_PHOTOS = r"M:\PHOTOS"
MYSQL_HOST     = "192.168.0.179"
MYSQL_PASSWORD = ""   # à mettre à jour après réinstallation
```

---

## Médecin référent
MEGRET Olivier — RPPS : 7817136230
