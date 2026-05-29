# -*- coding: utf-8 -*-
"""
Export StudioVision -> AlmaPro
Cliquer deux fois pour lancer l'export du patient ouvert dans StudioVision.
"""

import sys, os, re, shutil, traceback
from datetime import datetime
from pathlib import Path
from xml.dom import minidom
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import scrolledtext, messagebox

DEFAULT_MDB    = r"M:\fichier\PUBLIC.MDB"
DEFAULT_PHOTOS = r"M:\PHOTOS"
DEFAULT_DEST   = r"\\DESKTOP-MQKMFAQ\Users\Public\Documents\import"

DOC_EXTENSIONS = {".jpg",".jpeg",".jfif",".png",".bmp",".tif",".tiff",".dcm",".pdf",".rtf"}

# --- MDB ---------------------------------------------------------------

def _connect(mdb):
    import pyodbc
    for drv in ("Microsoft Access Driver (*.mdb, *.accdb)",
                "Microsoft Access Driver (*.mdb)"):
        try:
            return pyodbc.connect(f"DRIVER={{{drv}}};DBQ={mdb};", autocommit=True)
        except Exception:
            continue
    raise RuntimeError("Impossible d'ouvrir PUBLIC.MDB.\nVerifiez que Microsoft Access Database Engine est installe.")

def _rows(conn, sql, params=()):
    cur = conn.cursor(); cur.execute(sql, params)
    cols = [c[0] for c in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]

def _g(row, *keys, default=""):
    for k in keys:
        v = row.get(k)
        if v is not None:
            s = str(v).strip()
            if s and s.lower() not in ("none","nan","null"):
                return s
    return default

def _fdate(v, fmt="%d/%m/%Y"):
    if v is None: return ""
    if isinstance(v, datetime): return v.strftime(fmt)
    try:
        from datetime import date as _d
        if isinstance(v, _d): return v.strftime(fmt)
    except Exception: pass
    s = str(v).strip()
    for f in ("%Y-%m-%d","%d/%m/%Y","%m/%d/%Y","%d-%m-%Y"):
        try: return datetime.strptime(s[:10], f).strftime(fmt)
        except ValueError: pass
    return ""

def _fc(cols, cands):
    cl = {c.lower():c for c in cols}
    for c in cands:
        if c.lower() in cl: return cl[c.lower()]
    return None

# --- COM (fiche ouverte dans StudioVision) -----------------------------------

def get_active_patient():
    try:
        import win32com.client
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None: return None
        data = {}
        for i in range(form.Controls.Count):
            try:
                ctrl = form.Controls(i)
                name = str(ctrl.Name)
                if name in ("Code patient","NOM","Pr\xe9nom"):
                    data[name] = ctrl.Value
            except Exception: pass
        if "Code patient" in data:
            return {"code": str(data.get("Code patient","")).strip(),
                    "nom":  str(data.get("NOM","")).strip().upper(),
                    "prenom": str(data.get("Pr\xe9nom","")).strip()}
    except Exception: pass
    return None

# --- EXTRACTION PATIENT ------------------------------------------------------

def get_patient(mdb, code):
    conn = _connect(mdb)
    try:
        rows = _rows(conn, "SELECT * FROM Patients WHERE [Code patient] = ?", (int(code),))
    finally: conn.close()
    if not rows: return None
    r = rows[0]
    return {
        "code":        _g(r, "Code patient"),
        "nom":         _g(r, "NOM").upper(),
        "prenom":      _g(r, "Pr\xe9nom", "Prenom", "PRENOM"),
        "ddn":         _fdate(r.get("Date de naissance") or r.get("DDN")),
        "sexe":        _g(r, "Sexe","SEXE"),
        "num_secu":    _g(r, "NumSecu","NIR","Secu"),
        "adresse":     _g(r, "Adresse","ADRESSE"),
        "ville":       _g(r, "Ville","VILLE"),
        "cp":          _g(r, "CodePostal","CP"),
        "tel":         _g(r, "T\xe9l\xe9phone","Telephone","Tel"),
        "email":       _g(r, "Email","EMAIL"),
        "antecedents": _g(r, "Ant\xe9c\xe9dants","Antecedants","Ant\xe9c\xe9dents"),
        "allergies":   _g(r, "Allergies"),
        "traitements": _g(r, "Traitements"),
        "diagnostic":  _g(r, "Diagnostic OPH","Diagnostic  OPH","Diagnostic"),
        "important":   _g(r, "Important","IMPORTANT"),
    }

def get_consultations(mdb, code):
    conn = _connect(mdb)
    try:
        consults_raw = _rows(conn,
            "SELECT * FROM Consultation WHERE [Code patient] = ? ORDER BY [Date]",
            (int(code),))
        refrac_raw, kerato_raw = [], []
        try: refrac_raw = _rows(conn, "SELECT * FROM tREFRACTION")
        except Exception: pass
        try: kerato_raw = _rows(conn, "SELECT * FROM tKERATO")
        except Exception: pass
    finally: conn.close()

    rfx = {}
    for r in refrac_raw:
        nc = str(r.get("NumConsult","")).strip()
        if nc: rfx.setdefault(nc,[]).append(r)

    krt = {}
    for r in kerato_raw:
        nc = str(r.get("NumConsult","")).strip()
        if nc: krt.setdefault(nc,[]).append(r)

    result = []
    for c in consults_raw:
        ds = _fdate(c.get("Date"))
        if not ds: continue
        nc = str(c.get("N\xb0 consultation","")).strip()
        parts = []

        dom = _g(c,"DOMINANTE")
        if dom: parts.append(f"Motif : {dom}")

        av_sc_od = _g(c,"AVscOD","AV sc OD","AVSCOD","AVD sc")
        av_sc_og = _g(c,"AVscOG","AV sc OG","AVSCOG","AVG sc")
        av_cc_od = _g(c,"AVccOD","AV cc OD","AVCCOD","AVD cc")
        av_cc_og = _g(c,"AVccOG","AV cc OG","AVCCOG","AVG cc")
        if any([av_sc_od,av_sc_og,av_cc_od,av_cc_og]):
            parts.append("Acuite visuelle :\n"
                f"  OD sc={av_sc_od or '-'}  cc={av_cc_od or '-'}\n"
                f"  OG sc={av_sc_og or '-'}  cc={av_cc_og or '-'}")

        tod = _g(c,"TOD"); tog = _g(c,"TOG")
        if tod or tog:
            parts.append(f"PIO : OD={tod or '-'}  OG={tog or '-'} mmHg")

        ordo = _g(c,"Ordonnance")
        if ordo: parts.append(f"Ordonnance :\n{ordo}")
        autres = _g(c,"AutresPrescriptions")
        if autres: parts.append(f"Autres prescriptions :\n{autres}")

        refrac_t = _g(c,"REFRACTION")
        if refrac_t: parts.append(f"Refraction :\n{refrac_t}")

        # tREFRACTION
        TYPE_LABELS = {
            "1":"Subjectif loin","2":"Subjectif pres","3":"Subjectif inter",
            "4":"Cycloplegie","5":"Autoref","6":"Autoref",
            "7":"Contactologie","8":"Prescrit OD","9":"Prescrit OG",
            "10":"Prescrit pres","11":"Prescrit inter",
            "12":"Lunettes portees","13":"Lentilles portees",
        }
        if nc in rfx:
            for ref in rfx[nc]:
                typ = _g(ref,"TypeRef")
                lbl = TYPE_LABELS.get(typ, f"Type {typ}")
                s_od  = _g(ref,"SphOD","Sph_OD","SphD","SPHOD")
                c_od  = _g(ref,"CylOD","Cyl_OD","CylD","CYLOD")
                a_od  = _g(ref,"AxeOD","Axe_OD","AxD","AXEOD")
                add_od= _g(ref,"AddOD","Add_OD","AddD","ADDOD")
                s_og  = _g(ref,"SphOG","Sph_OG","SphG","SPHOG")
                c_og  = _g(ref,"CylOG","Cyl_OG","CylG","CYLOG")
                a_og  = _g(ref,"AxeOG","Axe_OG","AxG","AXEOG")
                add_og= _g(ref,"AddOG","Add_OG","AddG","ADDOG")
                av_od = _g(ref,"AVDL","AVDP")
                av_og = _g(ref,"AVGL","AVGP")
                if any([s_od,c_od,s_og,c_og]):
                    ln = f"Refraction ({lbl}) :\n"
                    if s_od or c_od:
                        ln += f"  OD : {s_od or '0'}  ({c_od or '0'} x {a_od or '0'})"
                        if add_od: ln += f"  Add {add_od}"
                        if av_od:  ln += f"  AV={av_od}"
                        ln += "\n"
                    if s_og or c_og:
                        ln += f"  OG : {s_og or '0'}  ({c_og or '0'} x {a_og or '0'})"
                        if add_og: ln += f"  Add {add_og}"
                        if av_og:  ln += f"  AV={av_og}"
                        ln += "\n"
                    parts.append(ln.rstrip())

        # tKERATO
        if nc in krt:
            for ker in krt[nc]:
                k1_od = _g(ker,"K1OD","K1_OD","K1D","Kflat_OD")
                k2_od = _g(ker,"K2OD","K2_OD","K2D","Ksteep_OD")
                km_od = _g(ker,"KmOD","Km_OD","KmD")
                ax_od = _g(ker,"AxeOD","Axe_OD","AxD","AXE_OD")
                k1_og = _g(ker,"K1OG","K1_OG","K1G","Kflat_OG")
                k2_og = _g(ker,"K2OG","K2_OG","K2G","Ksteep_OG")
                km_og = _g(ker,"KmOG","Km_OG","KmG")
                ax_og = _g(ker,"AxeOG","Axe_OG","AxG","AXE_OG")
                if any([k1_od,k1_og,km_od,km_og]):
                    ln = "Keratometrie :\n"
                    if k1_od or k2_od or km_od:
                        ln += f"  OD : K1={k1_od or '-'}  K2={k2_od or '-'}  Km={km_od or '-'}  Axe={ax_od or '-'}\n"
                    if k1_og or k2_og or km_og:
                        ln += f"  OG : K1={k1_og or '-'}  K2={k2_og or '-'}  Km={km_og or '-'}  Axe={ax_og or '-'}\n"
                    parts.append(ln.rstrip())

        texte = "\n\n".join(parts) if parts else "(Consultation sans notes)"
        result.append({"date":ds, "titre":(dom[:60] if dom else "Consultation"), "texte":texte})

    return sorted(result, key=lambda x: x["date"])

def get_documents(mdb, code, photos_root):
    conn = _connect(mdb)
    try:
        rows = _rows(conn,
            "SELECT * FROM Documents WHERE [code patient] = ? AND [Photo externe] IS NOT NULL ORDER BY [Date]",
            (int(code),))
        folder = None
        for r in rows:
            photo = _g(r,"Photo externe")
            if not photo: continue
            parts = [p for p in re.split(r"[/\\]", photo) if p]
            if len(parts) >= 2:
                c = Path(photos_root) / parts[0] / parts[1]
                if c.is_dir(): folder = c; break
    finally: conn.close()
    if not folder: return []
    files = []
    for root, _, fnames in os.walk(folder):
        for fname in sorted(fnames):
            ext = Path(fname).suffix.lower()
            if ext in DOC_EXTENSIONS:
                full = Path(root) / fname
                files.append({"src":full,"name":fname,
                               "date":datetime.fromtimestamp(full.stat().st_mtime).strftime("%d/%m/%Y")})
    return files

# --- XML ALMAPRO -------------------------------------------------------------

def _sub(parent, tag, text=""):
    el = ET.SubElement(parent, tag); el.text = text or ""; return el

def generate_xml(patient, consults, docs):
    root = ET.Element("Export_XML_AlmaPro")
    sexe = "F" if str(patient.get("sexe","")).upper().strip() in ("F","FEMME","2") else "M"
    d = ET.SubElement(root,"Dossiers")
    _sub(d,"Nom",                  patient["nom"])
    _sub(d,"Prenom",               patient["prenom"])
    _sub(d,"Sexe",                 sexe)
    _sub(d,"Adresse",              patient.get("adresse",""))
    _sub(d,"AdresseL2",            "")
    _sub(d,"Ville",                patient.get("ville",""))
    _sub(d,"CodePostal",           patient.get("cp",""))
    _sub(d,"Tel_Domicile",         patient.get("tel",""))
    _sub(d,"Tel_Pro",              "")
    _sub(d,"Tel_Port",             "")
    _sub(d,"FAX",                  "")
    _sub(d,"EMail",                patient.get("email",""))
    _sub(d,"Nom_Jeune_Fille",      "")
    _sub(d,"Date_de_naissance",    patient.get("ddn",""))
    _sub(d,"Date_Contrat_Signe",   "")
    _sub(d,"ContratSigne",         "FALSE")
    _sub(d,"Decede",               "FALSE")
    _sub(d,"Profession_du_patient","")
    _sub(d,"Num_Secu",             patient.get("num_secu",""))
    rem = [f"Import StudioVision {datetime.now().strftime('%d/%m/%Y')}"]
    if patient.get("diagnostic"): rem.append(f"Diag: {patient['diagnostic'][:200]}")
    if patient.get("antecedents"): rem.append(f"ATCD: {patient['antecedents'][:300]}")
    if patient.get("allergies"):   rem.append(f"Allergies: {patient['allergies'][:150]}")
    if patient.get("important"):   rem.append(f"Note: {patient['important'][:300]}")
    _sub(d,"Remarque","\n".join(rem)[:500])

    if not consults:
        consults = [{"date":datetime.now().strftime("%d/%m/%Y"),
                     "titre":"Import StudioVision",
                     "texte":f"Patient : {patient['nom']} {patient['prenom']}\nDDN : {patient.get('ddn','')}\nCode SV : {patient.get('code','')}"}]
    for c in consults:
        bloc = ET.SubElement(root,"Dossier_Consultations")
        _sub(bloc,"Date_Contact",  c["date"])
        _sub(bloc,"Titre_Contact", c["titre"][:100])
        _sub(bloc,"Texte_Contact", c["texte"])

    if patient.get("antecedents"):
        for item in re.split(r"[;\n/]", patient["antecedents"]):
            item = item.strip(" -")
            if len(item)>2: ET.SubElement(ET.SubElement(root,"ATCD"),"Texte_ATCD").text=item[:200]
    if patient.get("allergies"):
        for item in re.split(r"[,;\n/]", patient["allergies"]):
            item = item.strip(" -")
            if len(item)>2: ET.SubElement(ET.SubElement(root,"Allergie"),"Texte_Allergie").text=item[:100]
    for doc in docs:
        ext = Path(doc["name"]).suffix.lower()
        typ = {".tif":"OCT",".tiff":"OCT",".dcm":"DICOM",".pdf":"PDF"}.get(ext,"Image")
        bloc = ET.SubElement(root,"DocExt")
        _sub(bloc,"Nom_DocExt",        f"{typ} - {Path(doc['name']).stem}"[:100])
        _sub(bloc,"NomFichier_DocExt", doc["name"])
        _sub(bloc,"Emis_DocExt",       "FALSE")
        _sub(bloc,"Date_DocExt",       doc["date"])

    raw   = ET.tostring(root, encoding="unicode")
    dom   = minidom.parseString(raw)
    lines = (dom.toprettyxml(indent="   ", encoding="iso-8859-1")
               .decode("iso-8859-1", errors="replace").split("\n"))
    lines[0] = '<?xml version="1.0" encoding="iso-8859-1" standalone="yes"?>'
    return "\n".join(lines)

def write_export(dest, xml, docs):
    os.makedirs(dest, exist_ok=True)
    docext = os.path.join(dest,"docext"); os.makedirs(docext, exist_ok=True)
    n = 1
    while os.path.exists(os.path.join(dest,f"{n}.xml")): n+=1
    xml_path = os.path.join(dest,f"{n}.xml")
    with open(xml_path,"w",encoding="iso-8859-1",errors="replace") as f: f.write(xml)
    copied = 0
    for doc in docs:
        dst = os.path.join(docext,doc["name"])
        if os.path.exists(dst):
            base,ext = os.path.splitext(doc["name"])
            dst = os.path.join(docext,f"{base}_{n}{ext}")
        try: shutil.copy2(str(doc["src"]),dst); copied+=1
        except Exception: pass
    return xml_path, n, copied

# --- GUI ----------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Export StudioVision  ->  AlmaPro")
        self.resizable(False,False)
        self.configure(bg="#f0f4f8")
        # Icone : chercher dans _MEIPASS (exe PyInstaller) ou dossier du script
        try:
            if getattr(sys, 'frozen', False):
                base = Path(sys._MEIPASS)
            else:
                base = Path(__file__).parent
            ico = base / "export_sv_alma.ico"
            if ico.exists():
                self.iconbitmap(str(ico))
        except Exception:
            pass
        self._build()
        self.after(400, self._detect)

    def _build(self):
        tk.Label(self, text="Export StudioVision  ->  AlmaPro",
                 font=("Segoe UI",14,"bold"), bg="#f0f4f8", fg="#1a365d"
                 ).pack(pady=(18,4), padx=20)
        tk.Label(self, text="Exporte le dossier du patient ouvert dans StudioVision",
                 font=("Segoe UI",9), bg="#f0f4f8", fg="#4a5568"
                 ).pack(pady=(0,10))

        frm = tk.LabelFrame(self, text="Patient detecte",
                            font=("Segoe UI",9), bg="#f0f4f8", fg="#2d3748",
                            padx=12, pady=8)
        frm.pack(fill="x", padx=20, pady=(0,8))
        self._lbl_nom = tk.Label(frm, text="(en attente...)",
                                 font=("Segoe UI",11,"bold"), bg="#f0f4f8",
                                 fg="#2b6cb0", anchor="w")
        self._lbl_nom.pack(fill="x")
        self._lbl_ddn = tk.Label(frm, text="",
                                 font=("Segoe UI",9), bg="#f0f4f8",
                                 fg="#4a5568", anchor="w")
        self._lbl_ddn.pack(fill="x")

        self._btn = tk.Button(self,
            text="     Exporter vers AlmaPro",
            font=("Segoe UI",12,"bold"), bg="#2b6cb0", fg="white",
            activebackground="#2c5282", activeforeground="white",
            relief="flat", cursor="hand2", padx=20, pady=10,
            command=self._export, state="disabled")
        self._btn.pack(pady=(4,8), padx=20, fill="x")

        self._log = scrolledtext.ScrolledText(self,
            height=12, width=65,
            font=("Consolas",8), bg="#1a202c", fg="#e2e8f0",
            state="disabled", relief="flat", bd=0)
        self._log.pack(padx=20, pady=(0,8))

        tk.Button(self, text="Actualiser la detection du patient",
                  font=("Segoe UI",8), bg="#e2e8f0", fg="#4a5568",
                  relief="flat", cursor="hand2", command=self._detect
                  ).pack(pady=(0,14))
        self._patient = None

    def _log_msg(self, msg):
        self._log.configure(state="normal")
        self._log.insert("end", msg+"\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _detect(self):
        self._log_msg("Recherche du patient ouvert dans StudioVision...")
        active = get_active_patient()
        if active:
            try:
                p = get_patient(DEFAULT_MDB, active["code"])
                if p:
                    self._patient = p
                    self._lbl_nom.config(text=f"{p['nom']}  {p['prenom']}", fg="#276749")
                    self._lbl_ddn.config(text=f"DDN : {p['ddn']}   Code SV : {p['code']}")
                    self._btn.config(state="normal")
                    self._log_msg(f"  Patient detecte : {p['nom']} {p['prenom']} (DDN {p['ddn']})")
                    return
            except Exception as e:
                self._log_msg(f"  Erreur : {e}")
        self._lbl_nom.config(text="Aucun patient detecte dans StudioVision", fg="#c53030")
        self._lbl_ddn.config(text="Ouvrez la fiche patient dans StudioVision puis cliquez Actualiser")
        self._btn.config(state="disabled")
        self._log_msg("  Aucun patient detecte. Ouvrez StudioVision et la fiche patient.")

    def _export(self):
        if not self._patient: return
        self._btn.config(state="disabled", text="Export en cours...")
        p = self._patient
        self._log_msg(f"\nExport : {p['nom']} {p['prenom']}")
        self.update()
        try:
            code = p["code"]
            self._log_msg("  Extraction consultations et mesures...")
            consults = get_consultations(DEFAULT_MDB, code)
            self._log_msg(f"  {len(consults)} consultation(s)")

            self._log_msg("  Recherche documents...")
            docs = get_documents(DEFAULT_MDB, code, DEFAULT_PHOTOS)
            self._log_msg(f"  {len(docs)} document(s)")

            self._log_msg("  Generation XML AlmaPro...")
            xml = generate_xml(p, consults, docs)

            self._log_msg(f"  Ecriture vers {DEFAULT_DEST}...")
            xml_path, num, n_docs = write_export(DEFAULT_DEST, xml, docs)
            self._log_msg(f"  Fichier : {xml_path}")
            self._log_msg(f"  {n_docs} document(s) copie(s)")

            self._log_msg(f"\n{'='*48}")
            self._log_msg(f"EXPORT REUSSI - Dossier #{num}")
            self._log_msg(f"{'='*48}")
            self._log_msg("Dans AlmaPro (poste serveur) :")
            self._log_msg("  Import -> Importer dossiers AlmaPro")
            self._log_msg(f"  Chemin : {DEFAULT_DEST}")
            self._log_msg(f"  Cocher {p['nom']} {p['prenom']}")
            self._log_msg("  -> Importer les dossiers coches")

            messagebox.showinfo("Export reussi",
                f"Dossier #{num} pret a importer dans AlmaPro.\n\n"
                f"Patient : {p['nom']} {p['prenom']}\n"
                f"Consultations : {len(consults)}\n"
                f"Documents : {n_docs}\n\n"
                f"Dans AlmaPro :\nImport -> Importer dossiers AlmaPro\n"
                f"Chemin : {DEFAULT_DEST}")
        except Exception as e:
            self._log_msg(f"\nERREUR : {e}")
            self._log_msg(traceback.format_exc())
            messagebox.showerror("Erreur export", str(e))
        finally:
            self._btn.config(state="normal", text="     Exporter vers AlmaPro")

if __name__ == "__main__":
    App().mainloop()
