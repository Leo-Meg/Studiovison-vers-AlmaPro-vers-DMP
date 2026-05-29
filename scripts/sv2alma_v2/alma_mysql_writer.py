"""
alma_mysql_writer.py
Ecrit directement dans MySQL AlmaPro via PyMySQL (compatible MySQL 5.0).

Schema confirme depuis almapro.db :
  adm_patient      : NUMEROPATIENT, NOM, PRENOM, DATENAISS, SEXE, NUMASSURERO
  adm_administratif: NumeroPatient, ADRESSEL1, ADRESSEL2, VILLE, CODEPOSTAL,
                     Telephone, TelephonePortable, Email, CodeMedecin
  consu_contact    : NumeroContact(auto), NUMEROPATIENT, NumeroMedecin,
                     InitialesMedecin, NOM, Rtf_Cat1, Efface, Date
"""

from __future__ import annotations
import re
from datetime import datetime, date
from typing import Any

try:
    import pymysql
    MYSQL_OK = True
except ImportError:
    MYSQL_OK = False

MYSQL_HOST     = "192.168.0.179"
MYSQL_PORT     = 3306
MYSQL_USER     = "root"
MYSQL_PASSWORD = ""
MYSQL_DATABASE = "almapro"
MYSQL_CHARSET  = "latin1"

DEFAULT_RPPS            = "7817136230"
DEFAULT_MEDECIN_NOM     = "MEGRET"
DEFAULT_MEDECIN_PRENOM  = "Olivier"
DEFAULT_MEDECIN_INITIALES = "OM"
DEFAULT_MEDECIN_NUM     = 1    # numero interne AlmaPro du medecin


class AlmaWriter:

    def __init__(self, host=MYSQL_HOST, port=MYSQL_PORT,
                 user=MYSQL_USER, password=MYSQL_PASSWORD,
                 database=MYSQL_DATABASE):
        if not MYSQL_OK:
            raise RuntimeError("PyMySQL non installe.\n  pip install pymysql")
        self._cfg = dict(host=host, port=port, user=user,
                         password=password, database=database)
        self._conn = None
        self._schemas: dict[str, list[str]] = {}

    # -- Connexion --------------------------------------------------------------

    def connect(self) -> None:
        cfg = self._cfg
        try:
            self._conn = pymysql.connect(
                host=cfg["host"], port=cfg["port"],
                user=cfg["user"], password=cfg["password"],
                database=cfg["database"],
                charset=MYSQL_CHARSET,
                connect_timeout=10,
                autocommit=False,
            )
            print(f"  - Connect- - MySQL {cfg['host']}:{cfg['port']}/{cfg['database']}")
            self._load_schemas()
        except pymysql.err.OperationalError as e:
            code, msg = e.args
            if code == 2003:
                raise RuntimeError(
                    f"MySQL inaccessible ({cfg['host']}:{cfg['port']}).\n"
                    "  Verifiez que la VM est demarree et le port 3306 ouvert."
                ) from e
            elif code == 1045:
                raise RuntimeError(
                    f"Acces refuse MySQL.\n"
                    "  Lancez reset_mysql_v3.ps1 sur la VM (PowerShell admin)."
                ) from e
            raise

    def disconnect(self) -> None:
        if self._conn:
            try: self._conn.close()
            except Exception: pass

    def _load_schemas(self) -> None:
        tables = [
            "adm_patient", "adm_administratif", "consu_contact",
            "adm_acces_et_les_modifications_indentite",
        ]
        with self._conn.cursor() as cur:
            for t in tables:
                try:
                    cur.execute(f"DESCRIBE `{t}`")
                    self._schemas[t] = [r[0] for r in cur.fetchall()]
                except Exception:
                    self._schemas[t] = []

    # -- Utilitaires ------------------------------------------------------------

    def _fmt_date(self, value) -> str:
        """Convertit n'importe quelle date en YYYY-MM-DD. '1900-01-01' si vide."""
        if value is None:
            return "1900-01-01"
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d")
        if isinstance(value, date):
            return value.strftime("%Y-%m-%d")
        s = str(value).strip()
        if not s:
            return "1900-01-01"
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return "1900-01-01"

    def _next_num(self) -> int:
        """Prochain NUMEROPATIENT libre dans adm_patient."""
        with self._conn.cursor() as cur:
            cur.execute("SELECT MAX(NUMEROPATIENT) FROM `adm_patient`")
            row = cur.fetchone()
            return (row[0] or 0) + 1

    def _insert(self, table: str, data: dict[str, Any]) -> int:
        """Insere uniquement les colonnes connues. Retourne lastrowid."""
        schema   = self._schemas.get(table, [])
        filtered = {k: v for k, v in data.items() if k in schema}
        if not filtered:
            return 0
        cols   = list(filtered.keys())
        vals   = list(filtered.values())
        ph     = ", ".join(["%s"] * len(cols))
        cl     = ", ".join(f"`{c}`" for c in cols)
        with self._conn.cursor() as cur:
            cur.execute(f"INSERT INTO `{table}` ({cl}) VALUES ({ph})", vals)
        self._conn.commit()
        return self._conn.insert_id()

    # -- Verification doublon ---------------------------------------------------

    def _patient_exists(self, patient: dict) -> bool:
        nom    = patient.get("nom","").upper().strip()
        prenom = patient.get("prenom","").strip()
        ddn    = self._fmt_date(patient.get("ddn",""))
        if ddn == "1900-01-01": ddn = None
        if not nom: return False
        try:
            with self._conn.cursor() as cur:
                if ddn:
                    cur.execute(
                        "SELECT COUNT(*) FROM `adm_patient` "
                        "WHERE `NOM`=%s AND `PRENOM`=%s AND `DATENAISS`=%s",
                        (nom, prenom, ddn)
                    )
                else:
                    cur.execute(
                        "SELECT COUNT(*) FROM `adm_patient` "
                        "WHERE `NOM`=%s AND `PRENOM`=%s",
                        (nom, prenom)
                    )
                return cur.fetchone()[0] > 0
        except Exception:
            return False

    # -- Insertion patient complet ----------------------------------------------

    def insert_patient(self, patient: dict,
                       consultations: list[dict] | None = None,
                       atcd: list[str] | None = None,
                       allergies: list[str] | None = None) -> int:

        if not self._conn:
            raise RuntimeError("Non connecte. Appelez connect() d'abord.")
        if self._patient_exists(patient):
            nom    = patient.get("nom","").upper()
            prenom = patient.get("prenom","")
            ddn    = patient.get("ddn","")
            raise ValueError(
                f"Patient {nom} {prenom} (DDN {ddn}) existe deja dans AlmaPro.\n"
                f"Utilisez --force pour forcer (cree un doublon)."
            )

        num = self._next_num()
        print(f"  - Num-ro de dossier : #{num}")

        sexe_raw = str(patient.get("sexe","M")).upper().strip()
        sexe = "F" if sexe_raw in ("F","FEMME","FEMALE","2") else "M"
        ddn_mysql = self._fmt_date(patient.get("ddn",""))
        today     = datetime.now().strftime("%Y-%m-%d")

        # -- adm_patient --------------------------------------------------------
        self._insert("adm_patient", {
            "NUMEROPATIENT": num,
            "NOM":           patient.get("nom","").upper()[:50],
            "PRENOM":        patient.get("prenom","")[:50],
            "DATENAISS":     ddn_mysql,
            "SEXE":          sexe,
            "NUMASSURERO":   0,
        })
        print(f"  - adm_patient : {patient.get('nom','').upper()} {patient.get('prenom','')} (#{num})")

        # -- adm_administratif --------------------------------------------------
        self._insert("adm_administratif", {
            "NumeroPatient":      num,
            "ADRESSEL1":          patient.get("adresse","")[:255],
            "ADRESSEL2":          "",
            "VILLE":              patient.get("ville","")[:35],
            "CODEPOSTAL":         patient.get("cp","")[:12],
            "Telephone":          patient.get("tel","")[:15],
            "TelephonePortable":  patient.get("tel","")[:15],
            "TelephonePro":       "",
            "FAX":                "",
            "Email":              patient.get("email","")[:50],
            "CodeMedecin":        DEFAULT_MEDECIN_NUM,
            "Commentaire":        (
                f"Import StudioVision {datetime.now().strftime('%d/%m/%Y')}. "
                + (f"ATCD: {'; '.join(atcd[:3])}." if atcd else "")
                + (f" Allergies: {', '.join(allergies[:3])}." if allergies else "")
                + (f" Code SV: {patient.get('code','')}.")
            )[:255],
        })
        print(f"  - adm_administratif : OK")

        # -- consu_contact (consultations) --------------------------------------
        consults = consultations or []
        if not consults:
            consults = [{
                "date":  datetime.now().strftime("%d/%m/%Y"),
                "titre": "Import StudioVision",
                "texte": (
                    f"Dossier importe depuis StudioVision.\n"
                    f"Patient : {patient.get('nom','').upper()} {patient.get('prenom','')}\n"
                    f"DDN     : {patient.get('ddn','')}\n"
                    f"Code SV : {patient.get('code','')}\n"
                    f"Import  : {datetime.now().strftime('%d/%m/%Y a %H:%M')}"
                ),
            }]

        for i, c in enumerate(consults, 1):
            self._insert_consultation(num, c, i)
        print(f"  - consu_contact : {len(consults)} consultation(s)")

        # -- Journal d'audit ----------------------------------------------------
        try:
            self._insert_journal(num, patient)
        except Exception:
            pass

        return num

    def _insert_consultation(self, num_patient: int,
                              c: dict, ordre: int = 1) -> None:
        date_mysql = self._fmt_date(c.get("date",""))
        texte = c.get("texte","")
        if not texte.startswith("{\\rtf"):
            texte = self._to_rtf(texte)
        titre = c.get("titre","Consultation")[:50]

        self._insert("consu_contact", {
            "NUMEROPATIENT":    num_patient,
            "NumeroMedecin":    DEFAULT_MEDECIN_NUM,
            "InitialesMedecin": DEFAULT_MEDECIN_INITIALES,
            "NOM":              titre,
            "Rtf_Cat1":         texte,
            "Efface":           0,
            "Date":             date_mysql,
        })

    def _insert_journal(self, num_patient: int, patient: dict) -> None:
        now_str = datetime.now().strftime("%d/%m/%Y a %H:%M:%S")
        message = (
            f"Le {now_str}, {DEFAULT_MEDECIN_PRENOM} {DEFAULT_MEDECIN_NOM} "
            f"a cree le dossier de "
            f"{patient.get('nom','').upper()} {patient.get('prenom','')} "
            f"(import StudioVision automatique)."
        )
        self._insert("adm_acces_et_les_modifications_indentite", {
            "NUM_DOSSIER":  num_patient,
            "NUMDOSSIER":   num_patient,
            "DATE_ACCES":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "TEXTE_ACCES":  message,
            "TYPE_ACCES":   1,
        })

    # -- Utilitaires publics ----------------------------------------------------

    def list_patients(self, limit: int = 20) -> list[dict]:
        cols = self._schemas.get("adm_patient", [])
        sel  = [f"`{c}`" for c in
                ["NUMEROPATIENT","NOM","PRENOM","DATENAISS"]
                if c in cols]
        if not sel: return []
        with self._conn.cursor() as cur:
            cur.execute(f"SELECT {', '.join(sel)} FROM `adm_patient` "
                        f"ORDER BY NUMEROPATIENT DESC LIMIT {limit}")
            rows = cur.fetchall()
        names = [c.strip('`') for c in sel]
        return [dict(zip(names, r)) for r in rows]

    def get_schema(self, table: str) -> list[str]:
        return self._schemas.get(table, [])

    @staticmethod
    def _to_rtf(text: str) -> str:
        if not text:
            return r"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}{\colortbl;}\f0\fs20 }"
        escaped = text.replace("\\","\\\\").replace("{","\\{").replace("}","\\}")
        parts = []
        for ch in escaped:
            code = ord(ch)
            if code > 127:
                parts.append(f"\\'{code:02x}")
            elif ch == "\n":
                parts.append("\\par\r\n")
            elif ch == "\r":
                pass
            else:
                parts.append(ch)
        return (
            r"{\rtf1\ansi\ansicpg1252\deff0"
            r"{\fonttbl{\f0\fswiss\fcharset0 Arial;}}"
            r"{\colortbl;}\f0\fs20 "
            + "".join(parts) + "}"
        )
