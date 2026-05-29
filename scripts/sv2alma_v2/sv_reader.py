"""
sv_reader.py — Lecture de PUBLIC.MDB (StudioVision)
═══════════════════════════════════════════════════════
Deux sources de données :

  1. COM (win32com) — lit les champs de la FICHE OUVERTE dans StudioVision
     → get_active_patient()  — même méthode que Studiovision-Autosync

  2. ODBC (pyodbc)  — lit PUBLIC.MDB directement
     → get_patient_by_code()
     → get_patient_consultations()
     → get_patient_documents()
     → get_patient_folder_path()

Tables PUBLIC.MDB confirmées par OphthaView + Studiovision-Autosync :
  Patients    : NOM, Prénom, Code patient, Date de naissance,
                Antécédants, Allergies, Traitements, Diagnostic OPH,
                Important, Téléphone, …
  Consultation: Code patient, N° consultation, Date, REFRACTION,
                Ordonnance, AutresPrescriptions, DOMINANTE, …
  Documents   : code patient, Date, DESCRIPTIONS, TEXTE,
                Photo externe, TypeVW, NumDocExterne
"""

from __future__ import annotations
import os
import re
from datetime import datetime, date
from pathlib import Path

try:
    import pyodbc
    PYODBC_OK = True
except ImportError:
    PYODBC_OK = False

try:
    import win32com.client
    WIN32_OK = True
except ImportError:
    WIN32_OK = False


# ── Champs de la fiche ouverte dans StudioVision (noms de contrôles Access) ─
_COM_FIELD_CODE   = "Code patient"
_COM_FIELD_NOM    = "NOM"
_COM_FIELD_PRENOM = "Prénom"


# ── Extensions de fichiers considérés comme documents patients ──────────────
DOC_EXTENSIONS = {
    ".jpg", ".jpeg", ".jfif", ".png", ".bmp",
    ".tif", ".tiff", ".dcm", ".pdf", ".rtf",
    ".doc", ".docx", ".odt",
}

TYPE_BY_EXT = {
    ".tif": "OCT", ".tiff": "OCT",
    ".dcm": "DICOM",
    ".pdf": "Document", ".rtf": "Document",
    ".doc": "Document", ".docx": "Document", ".odt": "Document",
}


# ════════════════════════════════════════════════════════════════
# CONNEXION ODBC
# ════════════════════════════════════════════════════════════════

def _connect(mdb_path: str | Path) -> "pyodbc.Connection":
    if not PYODBC_OK:
        raise RuntimeError(
            "pyodbc non installé.\n"
            "  pip install pyodbc\n"
            "  + Microsoft Access Database Engine :\n"
            "    https://www.microsoft.com/en-us/download/details.aspx?id=54920"
        )
    for drv in (
        "Microsoft Access Driver (*.mdb, *.accdb)",
        "Microsoft Access Driver (*.mdb)",
    ):
        try:
            return pyodbc.connect(
                f"DRIVER={{{drv}}};DBQ={mdb_path};",
                autocommit=True,
            )
        except Exception:
            continue
    raise RuntimeError(
        f"Impossible d'ouvrir {mdb_path}.\n"
        "  Installez le pilote Access 64 bits :\n"
        "  https://www.microsoft.com/en-us/download/details.aspx?id=54920"
    )


def _rows(conn: "pyodbc.Connection", sql: str, params=()) -> list[dict]:
    cur = conn.cursor()
    cur.execute(sql, params)
    cols = [c[0] for c in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


def _get(row: dict, *keys: str, default: str = "") -> str:
    """Retourne la première valeur non-vide parmi les clés candidates."""
    for k in keys:
        v = row.get(k)
        if v is not None:
            s = str(v).strip()
            if s and s.lower() not in ("none", "nan", "null"):
                return s
    return default


def _fmt_date(value) -> str:
    """Convertit une valeur date quelconque en jj/mm/aaaa."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    s = str(value).strip()[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return ""


# ════════════════════════════════════════════════════════════════
# SOURCE 1 — FICHE OUVERTE VIA COM (Access ouvert)
# ════════════════════════════════════════════════════════════════

def get_active_patient() -> dict | None:
    """
    Lit les champs Code patient / NOM / Prénom du formulaire
    actuellement ouvert dans StudioVision (Access).
    Retourne None si StudioVision n'est pas ouvert ou si aucun
    patient n'est affiché.
    Méthode identique à Studiovision-Autosync v4.
    """
    if not WIN32_OK:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        targets = {_COM_FIELD_CODE, _COM_FIELD_NOM, _COM_FIELD_PRENOM}
        data: dict = {}

        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            try:
                name = str(ctrl.Name)
                if name in targets:
                    data[name] = ctrl.Value
            except Exception:
                pass

        if not targets.issubset(data.keys()):
            return None

        return {
            "code":   str(data[_COM_FIELD_CODE]).strip(),
            "nom":    str(data[_COM_FIELD_NOM]).strip().upper(),
            "prenom": str(data[_COM_FIELD_PRENOM]).strip(),
        }
    except Exception:
        return None


# ════════════════════════════════════════════════════════════════
# SOURCE 2 — LECTURE DIRECTE DE PUBLIC.MDB
# ════════════════════════════════════════════════════════════════

def get_patient_by_code(mdb_path: str, code: str) -> dict | None:
    """
    Charge la fiche complète d'un patient depuis la table Patients
    de PUBLIC.MDB, par son Code patient.
    """
    conn = _connect(mdb_path)
    try:
        rows = _rows(
            conn,
            "SELECT * FROM Patients WHERE [Code patient] = ?",
            (int(code),),
        )
    finally:
        conn.close()

    if not rows:
        return None

    row = rows[0]
    return {
        "code":        _get(row, "Code patient"),
        "nom":         _get(row, "NOM").upper(),
        "prenom":      _get(row, "Prénom", "PRENOM", "Prenom"),
        "ddn":         _fmt_date(row.get("Date de naissance")
                                  or row.get("DateNaissance")
                                  or row.get("DDN")),
        "sexe":        _get(row, "Sexe", "SEXE", "Genre"),
        "num_secu":    _get(row, "NumSecu", "NoSecu", "Secu", "NIR"),
        "adresse":     _get(row, "Adresse", "ADRESSE", "Adresse1"),
        "ville":       _get(row, "Ville", "VILLE"),
        "cp":          _get(row, "CodePostal", "CP", "code_postal"),
        "tel":         _get(row, "Téléphone", "Telephone", "Tel", "Mobile"),
        "email":       _get(row, "Email", "EMAIL", "Mail"),
        "profession":  _get(row, "Profession", "PROFESSION"),
        "antecedents": _get(row, "Antécédants", "Antecedants",
                              "Antécédents", "ANTECEDANTS"),
        "allergies":   _get(row, "Allergies"),
        "traitements": _get(row, "Traitements"),
        "diagnostic":  _get(row, "Diagnostic OPH", "Diagnostic  OPH",
                              "DiagnosticOPH", "Diagnostic Oph", "Diagnostic"),
        "important":   _get(row, "Important", "IMPORTANT", "Notes importantes"),
        "_raw":        row,
    }


def get_patient_consultations(mdb_path: str, code: str) -> list[dict]:
    """
    Retourne toutes les consultations du patient (table Consultation),
    triées par date croissante.
    """
    conn = _connect(mdb_path)
    try:
        rows = _rows(
            conn,
            "SELECT * FROM Consultation WHERE [Code patient] = ? "
            "ORDER BY [Date]",
            (int(code),),
        )
    finally:
        conn.close()

    results = []
    for row in rows:
        date_str = _fmt_date(row.get("Date"))
        if not date_str:
            continue

        # Construire le texte de la consultation à partir des champs cliniques
        parts: list[str] = []

        dominante = _get(row, "DOMINANTE")
        if dominante:
            parts.append(f"Motif : {dominante}")

        refraction = _get(row, "REFRACTION")
        if refraction:
            parts.append(f"Réfraction :\n{refraction}")

        ordonnance = _get(row, "Ordonnance")
        autres_presc = _get(row, "AutresPrescriptions")
        if ordonnance or autres_presc:
            presc = "\n".join(filter(None, [ordonnance, autres_presc]))
            parts.append(f"Prescriptions :\n{presc}")

        prochain_rdv = _get(row, "ProchainRDV")
        if prochain_rdv:
            parts.append(f"Prochain RDV : {prochain_rdv}")

        # Champs supplémentaires ophtalmologiques
        for label, keys in (
            ("AV OD sc", ("AVscOD", "AV sc OD", "AVSCOD")),
            ("AV OG sc", ("AVscOG", "AV sc OG", "AVSCOG")),
            ("TOD",      ("TOD",)),
            ("TOG",      ("TOG",)),
        ):
            val = _get(row, *keys)
            if val:
                parts.append(f"{label} : {val}")

        texte = "\n\n".join(parts) if parts else "(consultation sans notes)"

        results.append({
            "date":  date_str,
            "titre": dominante[:60] if dominante else "Consultation",
            "texte": texte,
        })

    return results


def get_patient_documents(mdb_path: str, code: str) -> list[dict]:
    """
    Retourne tous les Documents du patient qui ont un champ
    [Photo externe] non-null (fichiers image / PDF associés).
    """
    conn = _connect(mdb_path)
    try:
        rows = _rows(
            conn,
            "SELECT * FROM Documents "
            "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL "
            "ORDER BY [Date]",
            (int(code),),
        )
    finally:
        conn.close()

    results = []
    for row in rows:
        photo = _get(row, "Photo externe")
        if not photo:
            continue
        results.append({
            "photo_externe": photo,        # chemin relatif \groupe\dossier\fichier
            "date":          _fmt_date(row.get("Date")),
            "description":   _get(row, "DESCRIPTIONS"),
            "texte":         _get(row, "TEXTE"),
        })
    return results


def get_patient_folder_path(mdb_path: str, code: str,
                             photos_root: str) -> Path | None:
    """
    Résolution du dossier patient sur le disque.
    Méthode identique à Studiovision-Autosync :
      SELECT TOP 1 [Photo externe] FROM Documents WHERE [code patient] = ?
      → \<groupe>\<dossier_patient>\fichier.ext
      → photos_root / groupe / dossier_patient
    """
    conn = _connect(mdb_path)
    try:
        rows = _rows(
            conn,
            "SELECT TOP 1 [Photo externe] FROM Documents "
            "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL",
            (int(code),),
        )
    finally:
        conn.close()

    if not rows:
        return None

    photo_val = rows[0].get("Photo externe", "") or ""
    photo_val = str(photo_val).strip()
    if not photo_val:
        return None

    # Format : \groupe\dossier_patient\fichier.ext  (ou ///)
    parts = [p for p in re.split(r"[/\\]", photo_val) if p]
    if len(parts) < 2:
        return None

    folder = Path(photos_root) / parts[0] / parts[1]
    return folder if folder.is_dir() else None


def list_patient_files(folder: Path) -> list[dict]:
    """
    Liste tous les fichiers dans le dossier patient (récursif).
    Retourne seulement ceux dont l'extension est dans DOC_EXTENSIONS.
    """
    if not folder or not folder.is_dir():
        return []

    files = []
    for root, dirs, fnames in os.walk(folder):
        dirs.sort()
        for fname in sorted(fnames):
            ext = Path(fname).suffix.lower()
            if ext in DOC_EXTENSIONS:
                full = Path(root) / fname
                mtime = full.stat().st_mtime
                files.append({
                    "src":   full,
                    "name":  fname,
                    "date":  datetime.fromtimestamp(mtime).strftime("%d/%m/%Y"),
                    "type":  TYPE_BY_EXT.get(ext, "Image"),
                })
    return files
