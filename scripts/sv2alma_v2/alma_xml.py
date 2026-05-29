"""
alma_xml.py — Générateur XML format AlmaPro v3.2+
═══════════════════════════════════════════════════
Prend un dict patient + listes de consultations + documents
et produit le XML attendu par AlmaPro pour l'import "autre logiciel".

Spécification : "Importer des données d'un autre logiciel à partir
d'un fichier XML compatible AlmaPro" (v3.2+, 2017)

Encodage  : iso-8859-1 (obligatoire)
Dates     : jj/mm/aaaa (obligatoire)
Ordre des balises : strict (voir generate_xml)
"""

from __future__ import annotations
import os
from datetime import datetime
from xml.dom import minidom
import xml.etree.ElementTree as ET


def _sub(parent: ET.Element, tag: str, text: str = "") -> ET.Element:
    el = ET.SubElement(parent, tag)
    el.text = text or ""
    return el


def _normalise_sexe(val: str) -> str:
    v = val.strip().upper()
    if v in ("F", "FEMME", "FEMALE", "2", "MME"):
        return "F"
    return "M"


def generate_xml(patient: dict,
                 consultations: list[dict],
                 documents: list[dict],
                 ordonnances: list[dict] | None = None,
                 atcd: list[str] | None = None,
                 allergies: list[str] | None = None) -> str:
    """
    Génère le XML AlmaPro complet.

    Args:
        patient       : dict avec clés nom, prenom, ddn, sexe, adresse…
        consultations : list[{date, titre, texte}]
        documents     : list[{name, date, type}] — fichiers dans docext/
        ordonnances   : list[{date, texte, id}]  — optionnel
        atcd          : list[str]                — antécédents texte
        allergies     : list[str]                — allergies texte

    Retourne la chaîne XML encodée iso-8859-1.
    """
    root = ET.Element("Export_XML_AlmaPro")

    # ── Données administratives (ordre strict AlmaPro) ────────────────────
    d = ET.SubElement(root, "Dossiers")
    _sub(d, "Nom",                    patient.get("nom", "").upper())
    _sub(d, "Prenom",                 patient.get("prenom", ""))
    _sub(d, "Sexe",                   _normalise_sexe(patient.get("sexe", "M")))
    _sub(d, "Adresse",                patient.get("adresse", ""))
    _sub(d, "AdresseL2",              "")
    _sub(d, "Ville",                  patient.get("ville", ""))
    _sub(d, "CodePostal",             patient.get("cp", ""))
    _sub(d, "Tel_Domicile",           patient.get("tel", ""))
    _sub(d, "Tel_Pro",                "")
    _sub(d, "Tel_Port",               "")
    _sub(d, "FAX",                    "")
    _sub(d, "EMail",                  patient.get("email", ""))
    _sub(d, "Nom_Jeune_Fille",        "")
    _sub(d, "Date_de_naissance",      patient.get("ddn", ""))
    _sub(d, "Date_Contrat_Signe",     "")
    _sub(d, "ContratSigne",           "FALSE")
    _sub(d, "Decede",                 "FALSE")
    _sub(d, "Profession_du_patient",  patient.get("profession", ""))
    _sub(d, "Num_Secu",               patient.get("num_secu", ""))

    # Remarque = champs cliniques riches de StudioVision
    remarque_parts = [
        f"Import StudioVision — {datetime.now().strftime('%d/%m/%Y')}.",
        f"Code StudioVision : {patient.get('code', '')}",
    ]
    if patient.get("diagnostic"):
        remarque_parts.append(f"Diagnostic : {patient['diagnostic'][:200]}")
    if patient.get("important"):
        remarque_parts.append(f"Note importante : {patient['important'][:300]}")
    _sub(d, "Remarque", "\n".join(remarque_parts))

    # ── Consultations ─────────────────────────────────────────────────────
    if not consultations:
        consultations = [{
            "date":  datetime.now().strftime("%d/%m/%Y"),
            "titre": "Import StudioVision",
            "texte": (
                f"Dossier importé depuis StudioVision.\n"
                f"Patient : {patient.get('nom','').upper()} "
                f"{patient.get('prenom','')}\n"
                f"DDN : {patient.get('ddn','')}\n"
                f"Code StudioVision : {patient.get('code','')}"
            ),
        }]

    for c in consultations:
        bloc = ET.SubElement(root, "Dossier_Consultations")
        _sub(bloc, "Date_Contact",  c.get("date", datetime.now().strftime("%d/%m/%Y")))
        _sub(bloc, "Titre_Contact", c.get("titre", "Consultation")[:100])
        _sub(bloc, "Texte_Contact", c.get("texte", ""))

    # ── Ordonnances ───────────────────────────────────────────────────────
    for i, o in enumerate(ordonnances or [], start=1):
        bloc = ET.SubElement(root, "Ordonnance")
        _sub(bloc, "Date_Ordonnance",  o.get("date", "01/01/2000"))
        _sub(bloc, "Texte_Ordonnance", o.get("texte", ""))
        _sub(bloc, "ID_Ordonnance",    str(o.get("id", i)))

    # ── Antécédents ───────────────────────────────────────────────────────
    for a in (atcd or []):
        ET.SubElement(ET.SubElement(root, "ATCD"), "Texte_ATCD").text = a

    # ── Allergies ─────────────────────────────────────────────────────────
    for a in (allergies or []):
        ET.SubElement(ET.SubElement(root, "Allergie"), "Texte_Allergie").text = a

    # ── Documents externes (images, PDF…) ─────────────────────────────────
    for doc in documents:
        bloc = ET.SubElement(root, "DocExt")
        nom_affichage = f"{doc.get('type','Image')} — {os.path.splitext(doc['name'])[0]}"
        _sub(bloc, "Nom_DocExt",       nom_affichage[:100])
        _sub(bloc, "NomFichier_DocExt", doc["name"])
        _sub(bloc, "Emis_DocExt",      "FALSE")   # FALSE = RECU
        _sub(bloc, "Date_DocExt",      doc.get("date", datetime.now().strftime("%d/%m/%Y")))

    # ── Sérialisation iso-8859-1 ──────────────────────────────────────────
    raw   = ET.tostring(root, encoding="unicode")
    dom   = minidom.parseString(raw)
    lines = (
        dom.toprettyxml(indent="   ", encoding="iso-8859-1")
           .decode("iso-8859-1", errors="replace")
           .split("\n")
    )
    lines[0] = '<?xml version="1.0" encoding="iso-8859-1" standalone="yes"?>'
    return "\n".join(lines)
