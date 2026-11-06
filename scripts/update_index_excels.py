#!/usr/bin/env python3
"""Reconcile all final daily JSON reports with two separate XLSX workbooks."""
from __future__ import annotations

import json
import math
import re
import sys
import zipfile
import posixpath
from datetime import date, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
INDEX = Path("archive/شاخص کل تا 12 آبان 1405.xlsx")
TRADES = Path("archive/ارزش و حجم معاملات روزانه.xlsx")
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

def q(tag):
    return f"{{{MAIN}}}{tag}"

def gregorian_to_jalali(gy, gm, gd):
    mdays = [0,31,59,90,120,151,181,212,243,273,304,334]
    if gy > 1600:
        jy, gy = 979, gy - 1600
    else:
        jy, gy = 0, gy - 621
    gy2 = gy + 1 if gm > 2 else gy
    days = 365*gy + (gy2+3)//4 - (gy2+99)//100 + (gy2+399)//400 - 80 + gd + mdays[gm-1]
    jy += 33*(days//12053)
    days %= 12053
    jy += 4*(days//1461)
    days %= 1461
    if days > 365:
        jy += (days-1)//365
        days = (days-1)%365
    if days < 186:
        jm, jd = 1+days//31, 1+days%31
    else:
        jm, jd = 7+(days-186)//30, 1+(days-186)%30
    return jy, jm, jd

def canonical_date(value):
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 20000 <= float(value) <= 100000:
            g = date(1899, 12, 30) + timedelta(days=int(float(value)))
            y, m, d = gregorian_to_jalali(g.year, g.month, g.day)
            return f"{y:04d}-{m:02d}-{d:02d}"
        return None
    s = re.sub(r"\s+", "", str(value).translate(DIGITS))
    m = re.fullmatch(r"(\d{1,4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if not m:
        return None
    y, mo, d = map(int, m.groups())
    if y >= 1600:
        y, mo, d = gregorian_to_jalali(y, mo, d)
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return None
    return f"{y:04d}-{mo:02d}-{d:02d}"

def col_number(ref):
    n = 0
    for ch in re.sub(r"[^A-Z]", "", ref.upper()):
        n = n*26 + ord(ch)-64
    return n

def col_letters(n):
    out = ""
    while n:
        n, rem = divmod(n-1, 26)
        out = chr(65+rem) + out
    return out

CELLREF = re.compile(r"(\$?[A-Z]{1,3}\$?)(\d+)")

def shift_refs(text, insert_at):
    def repl(m):
        row = int(m.group(2))
        return m.group(1) + str(row+1 if row >= insert_at else row)
    return CELLREF.sub(repl, text)

def resolve_sheet(files, wanted_name=None):
    wb = ET.fromstring(files["xl/workbook.xml"])
    rels = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
    targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall(f"{{{PKGREL}}}Relationship")}
    sheets = wb.find(q("sheets"))
    chosen = None
    for sh in sheets.findall(q("sheet")):
        if wanted_name is None or sh.attrib.get("name") == wanted_name:
            chosen = sh
            break
    if chosen is None:
        raise RuntimeError(f"Worksheet not found: {wanted_name}")
    target = targets[chosen.attrib[f"{{{REL}}}id"]]
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))

class Book:
    def __init__(self, path, sheet_name=None):
        self.path = Path(path)
        with zipfile.ZipFile(self.path, "r") as z:
            bad = z.testzip()
            if bad:
                raise RuntimeError(f"Damaged XLSX part: {bad}")
            self.files = {n: z.read(n) for n in z.namelist()}
        self.shared = []
        if "xl/sharedStrings.xml" in self.files:
            ss = ET.fromstring(self.files["xl/sharedStrings.xml"])
            self.shared = ["".join(t.text or "" for t in si.iter(q("t"))) for si in ss.findall(q("si"))]
        self.sheet_path = resolve_sheet(self.files, sheet_name)
        self.root = ET.fromstring(self.files[self.sheet_path])

    def value(self, cell):
        if cell is None:
            return None
        typ = cell.attrib.get("t")
        if typ == "inlineStr":
            return "".join(t.text or "" for t in cell.iter(q("t")))
        v = cell.find(q("v"))
        if v is None or v.text is None:
            return None
        if typ == "s":
            try:
                return self.shared[int(v.text)]
            except (ValueError, IndexError):
                return None
        if typ in ("str", "e"):
            return v.text
        try:
            n = float(v.text)
            return int(n) if n.is_integer() else n
        except ValueError:
            return v.text

    def rows(self):
        data = self.root.find(q("sheetData"))
        if data is None:
            raise RuntimeError("Worksheet has no sheetData.")
        out = {}
        for row in data.findall(q("row")):
            rn = int(row.attrib["r"])
            out[rn] = {col_number(c.attrib["r"]): c for c in row.findall(q("c"))}
        return out

    def find_headers(self, need_index):
        date_hits, index_hits = [], []
        for rn, cells in self.rows().items():
            if rn > 40:
                continue
            for cn, cell in cells.items():
                val = self.value(cell)
                if not isinstance(val, str):
                    continue
                clean = re.sub(r"[\s\u200c]+", "", val.translate(DIGITS))
                if "تاریخ" in clean and "تاریخدرج" not in clean:
                    date_hits.append((rn, cn))
                if need_index and "آخرین" in clean and "واحد" in clean:
                    index_hits.append((rn, cn))
        if not date_hits:
            raise RuntimeError("Date header not found within first 40 rows.")
        if need_index and not index_hits:
            raise RuntimeError("Index header 'آخرین (واحد)' not found within first 40 rows.")
        dr, dc = date_hits[0]
        if need_index:
            ir, ic = index_hits[0]
            hr = max(dr, ir)
            if dc == ic:
                raise RuntimeError("Date and index columns resolved to the same column.")
            return hr, dc, ic
        return dr, dc, None

    def dates(self, col, first_row):
        dates = set()
        for rn, cells in self.rows().items():
            if rn >= first_row:
                d = canonical_date(self.value(cells.get(col)))
                if d:
                    dates.add(d)
        return dates

    def insert(self, header_row, col_values, values, fallback_styles=None):
        """Insert one complete data row directly under headers; col_values maps column->value."""
        at = header_row + 1
        row_map = self.rows()
        sample_row = None
        date_col = min(col_values)
        for rn in sorted(row_map):
            if rn >= at and row_map[rn].get(date_col) is not None:
                if canonical_date(self.value(row_map[rn][date_col])):
                    sample_row = rn
                    break
        sample_cells = row_map.get(sample_row, {}) if sample_row else {}
        styles = {}
        for col in col_values:
            cell = sample_cells.get(col)
            style = cell.attrib.get("s") if cell is not None else None
            if style is None and fallback_styles:
                style = fallback_styles.get(col)
            if style is not None:
                styles[col] = style
        data = self.root.find(q("sheetData"))
        sample_el = next((r for r in data.findall(q("row")) if sample_row and int(r.attrib["r"]) == sample_row), None)

        for row in data.findall(q("row")):
            old = int(row.attrib["r"])
            if old >= at:
                row.set("r", str(old+1))
                for cell in row.findall(q("c")):
                    cell.set("r", shift_refs(cell.attrib["r"], at))
                for child in row.iter(q("f")):
                    if child.text:
                        child.text = shift_refs(child.text, at)

        for el in self.root.iter():
            for attr in ("ref", "sqref"):
                ref = el.attrib.get(attr)
                if ref:
                    el.set(attr, shift_refs(ref, at))

        attrs = {"r": str(at)}
        if sample_el is not None:
            for key in ("ht", "customHeight", "s", "customFormat"):
                if key in sample_el.attrib:
                    attrs[key] = sample_el.attrib[key]
        new_row = ET.Element(q("row"), attrs)
        for col, value in sorted(col_values.items()):
            cell = ET.SubElement(new_row, q("c"), {"r": f"{col_letters(col)}{at}"})
            if col in styles:
                cell.set("s", styles[col])
            if isinstance(value, str):
                cell.set("t", "inlineStr")
                inline = ET.SubElement(cell, q("is"))
                ET.SubElement(inline, q("t")).text = value
            else:
                if isinstance(value, float) and not math.isfinite(value):
                    raise RuntimeError("Refusing non-finite numeric value.")
                ET.SubElement(cell, q("v")).text = str(value)

        rows = data.findall(q("row"))
        idx = next((i for i, row in enumerate(rows) if int(row.attrib["r"]) > at), len(rows))
        data.insert(idx, new_row)

    def save(self):
        ET.register_namespace("", MAIN)
        ET.register_namespace("r", REL)
        ET.register_namespace("pr", PKGREL)
        self.files[self.sheet_path] = ET.tostring(self.root, encoding="utf-8", xml_declaration=True)
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for name, body in self.files.items():
                z.writestr(name, body)
        with zipfile.ZipFile(tmp, "r") as z:
            bad = z.testzip()
            if bad:
                raise RuntimeError(f"Saved XLSX failed integrity test: {bad}")
        tmp.replace(self.path)

def jalali_text(date_string, sample):
    sep = "/" if sample and "/" in str(sample) else "-"
    y, m, d = date_string.split("-")
    out = sep.join((y, m, d))
    if sample and any(ch in str(sample) for ch in "۰۱۲۳۴۵۶۷۸۹"):
        out = out.translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))
    return out

def load_records():
    found = {}
    for path in sorted(Path("archive").glob("index-*.json")):
        obj = json.loads(path.read_text(encoding="utf-8"))
        d = canonical_date(obj.get("date"))
        if not d:
            raise RuntimeError(f"Invalid/missing date in {path}")
        if obj.get("state") != "بسته":
            print(f"SKIP {d}: market not closed")
            continue
        for key in ("index", "tval", "tvol"):
            v = obj.get(key)
            if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(float(v)):
                raise RuntimeError(f"{path}: invalid {key}={v!r}")
        if d in found:
            raise RuntimeError(f"Duplicate JSON archive date: {d}")
        found[d] = {"date": d, "index": float(obj["index"]), "tval": int(obj["tval"]), "tvol": int(obj["tvol"])}
    if not found:
        raise RuntimeError("No closed-market index JSON reports found.")
    return [found[d] for d in sorted(found)]

def update_index(records):
    if not INDEX.is_file():
        raise RuntimeError(f"Index workbook missing: {INDEX}")
    b = Book(INDEX)
    hr, dc, ic = b.find_headers(True)
    known = b.dates(dc, hr+1)
    added, skipped = [], []
    for rec in records:
        d = rec["date"]
        if d in known:
            skipped.append(d)
            continue
        existing_rows = b.rows()
        sample = next((b.value(cells.get(dc)) for rn, cells in sorted(existing_rows.items()) if rn > hr and canonical_date(b.value(cells.get(dc)))), None)
        b.insert(hr, {dc: jalali_text(d, sample), ic: rec["index"]}, {dc: jalali_text(d, sample), ic: rec["index"]})
        known.add(d)
        added.append(d)
    if added:
        b.save()
    v = Book(INDEX)
    h, dcol, icol = v.find_headers(True)
    after = v.dates(dcol, h+1)
    missing = set(r["date"] for r in records)-after
    if missing:
        raise RuntimeError(f"Index workbook missing archive dates after update: {sorted(missing)}")
    top = v.rows().get(h+1, {})
    top_date = canonical_date(v.value(top.get(dcol)))
    if top_date != max(after) or v.value(top.get(icol)) is None:
        raise RuntimeError(f"Index newest row not immediately below headers: row {h+1} date={top_date}, expected={max(after)}")
    return added, skipped, f"header row={h}; date col={col_letters(dcol)}; index col={col_letters(icol)}; top date={top_date}"

def update_trades(records):
    if not TRADES.is_file():
        raise RuntimeError(f"Separate trading workbook missing: {TRADES}")
    b = Book(TRADES, "ارزش و حجم معاملات")
    hr, dc, _ = b.find_headers(False)
    if hr != 1 or dc != 1:
        raise RuntimeError(f"Unexpected independent trading workbook layout: header row={hr}, date col={dc}")
    known = b.dates(1, 2)
    added, skipped = [], []
    for rec in records:
        d = rec["date"]
        if d in known:
            skipped.append(d)
            continue
        b.insert(1, {1: d, 2: rec["tval"], 3: rec["tvol"]}, {}, {1: "1", 2: "11", 3: "11"})
        known.add(d)
        added.append(d)
    if added:
        rows = b.root.find(q("sheetData")).findall(q("row"))
        maxrow = max(int(x.attrib["r"]) for x in rows)
        for af in b.root.findall(q("autoFilter")):
            af.set("ref", f"A1:C{maxrow}")
        dim = b.root.find(q("dimension"))
        if dim is not None:
            dim.set("ref", f"A1:C{maxrow}")
        b.save()
    v = Book(TRADES, "ارزش و حجم معاملات")
    h, dc, _ = v.find_headers(False)
    after = v.dates(dc, h+1)
    missing = set(r["date"] for r in records)-after
    if missing:
        raise RuntimeError(f"Trading workbook missing archive dates: {sorted(missing)}")
    top = v.rows().get(h+1, {})
    top_date = canonical_date(v.value(top.get(1)))
    if top_date != max(after) or v.value(top.get(2)) is None or v.value(top.get(3)) is None:
        raise RuntimeError(f"Trading workbook newest row incomplete or misplaced at row {h+1}.")
    return added, skipped, f"header row=1; columns A/B/C; top date={top_date}"

def main():
    records = load_records()
    print(f"Closed JSON reports found: {len(records)}")
    for r in records:
        print(f"{r['date']}: index={r['index']}, tval={r['tval']} rial, tvol={r['tvol']} shares")
    ia, ise, ii = update_index(records)
    ta, tse, ti = update_trades(records)
    print("\nINDEX WORKBOOK")
    print("Inserted:", ia or "none")
    print("Already present:", ise or "none")
    print(ii)
    print("\nSEPARATE TRADING WORKBOOK")
    print("Inserted:", ta or "none")
    print("Already present:", tse or "none")
    print(ti)
    print("\nCHECK: all archived JSON dates matched in both separate workbooks; duplicate insertions: none.")
    return 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        raise
