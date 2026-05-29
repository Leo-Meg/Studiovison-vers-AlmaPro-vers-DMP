"""
analyse_ibdata1.py
==================
Outils d'analyse binaire du fichier ibdata1 de MySQL 5.0 (InnoDB).

Usage :
    python analyse_ibdata1.py --ibdata ibdata1 --action info
    python analyse_ibdata1.py --ibdata ibdata1 --action patients
    python analyse_ibdata1.py --ibdata ibdata1 --action find --nom TEST --ddn 23/11/1990
    python analyse_ibdata1.py --ibdata ibdata1 --action dates
    python analyse_ibdata1.py --ibdata ibdata1 --action strings
"""

import re, sys, argparse
from datetime import date
from pathlib import Path

INNODB_PAGE_SIZE = 16384
INNODB_DATE_XOR  = 0x800000
FIELD_NOM_WIDTH  = 50

def decode_innodb_date(b1, b2, b3):
    val   = ((b1 << 16) | (b2 << 8) | b3) ^ INNODB_DATE_XOR
    day   =  val & 0x1F
    month = (val >> 5) & 0x0F
    year  =  val >> 9
    if 1900 < year < 2100 and 1 <= month <= 12 and 1 <= day <= 31:
        try: date(year, month, day); return f"{day:02d}/{month:02d}/{year}"
        except ValueError: pass
    return None

def find_dates_in_chunk(chunk, max_dates=20):
    seen = []
    for i in range(len(chunk) - 2):
        d = decode_innodb_date(chunk[i], chunk[i+1], chunk[i+2])
        if d and d not in seen:
            seen.append(d)
            if len(seen) >= max_dates: break
    return seen

def load(path):
    size = Path(path).stat().st_size
    print(f"Chargement {path} ({size/1024/1024:.1f} Mo)...", end=" ", flush=True)
    with open(path, "rb") as f: data = f.read()
    print("OK")
    return data

def action_info(data):
    pages = len(data) // INNODB_PAGE_SIZE
    print(f"Taille : {len(data):,} octets ({len(data)/1024/1024:.1f} Mo)")
    print(f"Pages  : {pages} pages de 16 Ko")
    full   = data.decode("latin-1", errors="replace")
    tables = set(re.findall(r'almapro/([a-z_]{4,40})', full))
    print(f"Tables : {len(tables)} references")
    for t in sorted(tables)[:20]: print(f"  almapro/{t}")

def action_find(data, nom, ddn):
    nom_key = (nom.upper() + " " * FIELD_NOM_WIDTH)[:FIELD_NOM_WIDTH].encode("latin-1")
    print(f"Recherche '{nom.upper()}' DDN={ddn}...")
    found, start, attempts = [], 0, 0
    while True:
        pos = data.find(nom_key, start)
        if pos == -1: break
        attempts += 1
        chunk = data[pos:pos+200]
        dates = find_dates_in_chunk(chunk[:150])
        if ddn in dates:
            print(f"  TROUVE offset {pos:,}  dates={dates[:5]}")
            found.append(pos)
            for i in range(0, 100, 16):
                row = chunk[i:i+16]
                h = " ".join(f"{b:02x}" for b in row)
                a = "".join(chr(b) if 32 <= b < 127 else "." for b in row)
                print(f"    +{i:03d}: {h:<48}  |{a}|")
            print()
        start = pos + 1
    print(f"Occurrences nom : {attempts} | Correspondances nom+DDN : {len(found)}")

def action_patients(data):
    seen, results = set(), []
    for m in re.finditer(rb'([A-Z][A-Z \-]{2,48}\x20{1,49})', data):
        pos = m.start()
        nom = m.group(1).decode("latin-1").strip()
        if len(nom) < 2: continue
        chunk = data[pos:pos+200]
        dates = [d for d in find_dates_in_chunk(chunk[:150]) if "1900" not in d]
        if not dates: continue
        key = (nom, dates[0])
        if key in seen: continue
        seen.add(key)
        results.append({"nom": nom, "ddn": dates[0], "offset": pos})
    print(f"{len(results)} patients potentiels:\n")
    for p in results[:40]:
        print(f"  {p['nom']:<30}  DDN={p['ddn']}  (offset {p['offset']:,})")

def action_dates(data):
    from collections import Counter
    c = Counter()
    for i in range(len(data)-2):
        d = decode_innodb_date(data[i], data[i+1], data[i+2])
        if d: c[d] += 1
    print(f"{len(c)} dates distinctes:")
    for d, n in sorted(c.items())[:50]: print(f"  {d}  ({n}x)")

def main():
    parser = argparse.ArgumentParser(description="Analyse ibdata1 MySQL 5.0")
    parser.add_argument("--ibdata",  required=True)
    parser.add_argument("--action",  required=True,
                        choices=["info","patients","find","dates","strings"])
    parser.add_argument("--nom",     default="")
    parser.add_argument("--ddn",     default="")
    args = parser.parse_args()
    if not Path(args.ibdata).exists():
        print(f"Erreur : {args.ibdata} introuvable"); sys.exit(1)
    data = load(args.ibdata); print()
    if args.action == "info":      action_info(data)
    elif args.action == "patients": action_patients(data)
    elif args.action == "dates":    action_dates(data)
    elif args.action == "find":
        if not args.nom or not args.ddn:
            print("--action find requiert --nom et --ddn"); sys.exit(1)
        action_find(data, args.nom, args.ddn)

if __name__ == "__main__":
    main()
