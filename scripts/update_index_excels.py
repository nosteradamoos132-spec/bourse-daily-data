#!/usr/bin/env python3
"""Reconcile all closed daily BRSAPI JSON reports into one five-column market workbook."""
from __future__ import annotations

import json
import math
import re
import sys
import zipfile
import posixpath
from datetime import date, timedelta
from pathlib import Path
from copy import deepcopy
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
MASTER = Path("archive/داده‌های روزانه بازار.xlsx")
BILLION = 1_000_000_000
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

def row_cells(row):
    return {col_number(cell.attrib["r"]): cell for cell in row.findall(q("c")) if cell.attrib.get("r")}

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
                if need_index and ("شاخصکل" in clean or ("آخرین" in clean and "واحد" in clean)):
                    index_hits.append((rn, cn))
        if not date_hits:
            raise RuntimeError("Date header not found within first 40 rows.")
        if need_index and not index_hits:
            raise RuntimeError("Index value header 'شاخص کل' (or 'آخرین (واحد)') not found within first 40 rows.")
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

    def apply_column_number_format(self, columns, first_data_row, format_code):
        """Set a display format for numeric cells without changing their stored values."""
        styles_path = "xl/styles.xml"
        if styles_path not in self.files:
            raise RuntimeError("Workbook has no xl/styles.xml file.")

        ET.register_namespace("", MAIN)
        styles = ET.fromstring(self.files[styles_path])
        num_fmts = styles.find(q("numFmts"))
        if num_fmts is None:
            num_fmts = ET.Element(q("numFmts"), {"count": "0"})
            styles.insert(0, num_fmts)

        num_fmt = next(
            (item for item in num_fmts.findall(q("numFmt"))
             if item.attrib.get("formatCode") == format_code),
            None,
        )
        if num_fmt is None:
            used_ids = {
                int(item.attrib["numFmtId"])
                for item in num_fmts.findall(q("numFmt"))
                if item.attrib.get("numFmtId", "").isdigit()
            }
            format_id = 164
            while format_id in used_ids:
                format_id += 1
            ET.SubElement(
                num_fmts,
                q("numFmt"),
                {"numFmtId": str(format_id), "formatCode": format_code},
            )
        else:
            format_id = int(num_fmt.attrib["numFmtId"])
        num_fmts.set("count", str(len(num_fmts.findall(q("numFmt")))))

        cell_xfs = styles.find(q("cellXfs"))
        if cell_xfs is None:
            raise RuntimeError("Workbook styles have no cellXfs collection.")
        xfs = cell_xfs.findall(q("xf"))
        if not xfs:
            raise RuntimeError("Workbook styles have no base cell format.")

        style_ids = {}
        sheet_data = self.root.find(q("sheetData"))
        for row in sheet_data.findall(q("row")):
            if int(row.attrib.get("r", "0")) < first_data_row:
                continue
            cells = row_cells(row)
            for column in columns:
                cell = cells.get(column)
                if cell is None or cell.find(q("v")) is None:
                    continue
                try:
                    base_style_id = int(cell.attrib.get("s", "0"))
                except ValueError as exc:
                    raise RuntimeError("Invalid cell style index.") from exc
                if not (0 <= base_style_id < len(xfs)):
                    raise RuntimeError("Cell style index is outside the workbook style table.")

                key = (base_style_id, format_id)
                if key not in style_ids:
                    base_xf = xfs[base_style_id]
                    formatted_xf = deepcopy(base_xf)
                    formatted_xf.set("numFmtId", str(format_id))
                    formatted_xf.set("applyNumberFormat", "1")

                    existing_id = None
                    candidate_children = [
                        ET.tostring(child, encoding="unicode")
                        for child in list(formatted_xf)
                    ]
                    for index, candidate in enumerate(cell_xfs.findall(q("xf"))):
                        candidate_children_existing = [
                            ET.tostring(child, encoding="unicode")
                            for child in list(candidate)
                        ]
                        if (
                            dict(candidate.attrib) == dict(formatted_xf.attrib)
                            and candidate_children_existing == candidate_children
                        ):
                            existing_id = index
                            break

                    if existing_id is None:
                        existing_id = len(cell_xfs.findall(q("xf")))
                        cell_xfs.append(formatted_xf)
                        cell_xfs.set("count", str(len(cell_xfs.findall(q("xf")))))
                    style_ids[key] = existing_id

                cell.set("s", str(style_ids[key]))

        self.files[styles_path] = ET.tostring(
            styles, encoding="utf-8", xml_declaration=True
        )

    def number_format(self, cell):
        if cell is None:
            return None
        styles_path = "xl/styles.xml"
        styles = ET.fromstring(self.files[styles_path])
        cell_xfs = styles.find(q("cellXfs"))
        num_fmts = styles.find(q("numFmts"))
        if cell_xfs is None or num_fmts is None:
            return None
        try:
            style_id = int(cell.attrib.get("s", "0"))
            format_id = int(cell_xfs.findall(q("xf"))[style_id].attrib.get("numFmtId", "0"))
        except (ValueError, IndexError):
            return None
        for item in num_fmts.findall(q("numFmt")):
            if int(item.attrib.get("numFmtId", "-1")) == format_id:
                return item.attrib.get("formatCode")
        return None

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
        problems = []
        for key in ("index", "index_equalWeight", "tval", "tvol"):
            v = obj.get(key)
            if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(float(v)):
                problems.append(f"invalid {key}={v!r}")

        report_time = obj.get("time")
        if not isinstance(report_time, str) or not re.fullmatch(r"\d{2}:\d{2}:\d{2}", report_time) or report_time < "17:00:00":
            problems.append(f"report time is too early or invalid: {report_time!r}")

        for key in ("mv", "tno", "tval", "tvol"):
            v = obj.get(key)
            if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(float(v)) or float(v) <= 0:
                problems.append(f"incomplete/non-positive {key}={v!r}")

        if problems:
            print(f"SKIP INVALID REPORT {d} ({path.name}): {'; '.join(problems)}")
            continue

        if d in found:
            raise RuntimeError(f"Duplicate valid JSON archive date: {d}")
        found[d] = {"date": d, "index": float(obj["index"]), "index_equalWeight": float(obj["index_equalWeight"]), "tval": int(obj["tval"]), "tvol": int(obj["tvol"])}
    if not found:
        raise RuntimeError("No closed-market index JSON reports found.")
    return [found[d] for d in sorted(found)]

def set_text_cell(cell, value):
    for child in list(cell):
        if child.tag in (q("v"), q("is")):
            cell.remove(child)
    cell.set("t", "inlineStr")
    inline = ET.SubElement(cell, q("is"))
    ET.SubElement(inline, q("t")).text = value


def set_numeric_cell(cell, value):
    for child in list(cell):
        if child.tag in (q("v"), q("is")):
            cell.remove(child)
    cell.attrib.pop("t", None)
    if isinstance(value, float) and not math.isfinite(value):
        raise RuntimeError("Refusing to write a non-finite numeric value.")
    ET.SubElement(cell, q("v")).text = str(value)


def update_master(records):
    valid_dates = {record["date"] for record in records}
    if not MASTER.is_file():
        raise RuntimeError(f"Unified workbook missing: {MASTER}")

    book = Book(MASTER, "داده‌های روزانه بازار")
    rows = book.rows()
    header_row = 1
    header = rows.get(header_row, {})
    expected_headers = {
        1: "تاریخ (شمسی)",
        2: "شاخص کل",
        3: "شاخص هم‌وزن",
    }
    for column, expected in expected_headers.items():
        actual = book.value(header.get(column))
        if actual != expected:
            raise RuntimeError(
                f"Unexpected workbook header at column {column}: {actual!r}; expected {expected!r}"
            )

    for column, label in ((4, "ارزش معاملات (B)"), (5, "حجم معاملات (B)")):
        cell = header.get(column)
        if cell is None:
            raise RuntimeError(f"Missing header cell in column {column}")
        set_text_cell(cell, label)

    row_by_date = {}
    for row_number, cells in rows.items():
        if row_number <= header_row:
            continue
        date_key = canonical_date(book.value(cells.get(1)))
        if not date_key:
            if any(book.value(cell) not in (None, "") for cell in cells.values()):
                raise RuntimeError(f"Data row {row_number} has no valid date.")
            continue
        if date_key in row_by_date:
            raise RuntimeError(f"Duplicate date in unified workbook: {date_key}")
        row_by_date[date_key] = row_number

    inserted = []
    refreshed = []
    for record in records:
        date_key = record["date"]
        row_number = row_by_date.get(date_key)
        if row_number is None:
            book.insert(
                header_row,
                {
                    1: date_key,
                    2: record["index"],
                    3: record["index_equalWeight"],
                    4: record["tval"] / BILLION,
                    5: record["tvol"] / BILLION,
                },
                {},
            )
            inserted.append(date_key)
            row_by_date[date_key] = header_row + 1
            rows = book.rows()
            row_by_date = {}
            for current_row, cells in rows.items():
                if current_row > header_row:
                    current_date = canonical_date(book.value(cells.get(1)))
                    if current_date:
                        row_by_date[current_date] = current_row
        else:
            cells = rows.get(row_number, {})
            for column, value in (
                (2, record["index"]),
                (3, record["index_equalWeight"]),
            ):
                cell = cells.get(column)
                if cell is None:
                    cell = ET.SubElement(
                        next(row for row in book.root.find(q("sheetData")).findall(q("row"))
                             if int(row.attrib["r"]) == row_number),
                        q("c"),
                        {"r": f"{col_letters(column)}{row_number}"},
                    )
                set_numeric_cell(cell, value)

            for column, key in ((4, "tval"), (5, "tvol")):
                cell = cells.get(column)
                if cell is None:
                    cell = ET.SubElement(
                        next(row for row in book.root.find(q("sheetData")).findall(q("row"))
                             if int(row.attrib["r"]) == row_number),
                        q("c"),
                        {"r": f"{col_letters(column)}{row_number}"},
                    )
                set_numeric_cell(cell, record[key] / BILLION)
            refreshed.append(date_key)

    sheet_data = book.root.find(q("sheetData"))
    row_nodes = list(sheet_data.findall(q("row")))
    header_node = next(
        (row for row in row_nodes if int(row.attrib.get("r", "0")) == header_row),
        None,
    )
    data_rows = []
    for row in row_nodes:
        row_number = int(row.attrib.get("r", "0"))
        if row_number <= header_row:
            continue
        cells = row_cells(row)
        date_key = canonical_date(book.value(cells.get(1)))
        if date_key in valid_dates:
            data_rows.append((date_key, row))
        elif date_key:
            # Keep manually curated historical rows that do not have daily JSON files.
            # The JSON archive remains authoritative for dates it contains.
            data_rows.append((date_key, row))
            print(f"PRESERVE HISTORICAL WORKBOOK ROW: {date_key}")
    data_rows.sort(key=lambda item: item[0], reverse=True)

    for child in list(sheet_data):
        sheet_data.remove(child)
    sheet_data.append(header_node)
    for row_number, (_, row) in enumerate(data_rows, start=header_row + 1):
        row.set("r", str(row_number))
        for cell in row.findall(q("c")):
            column = col_number(cell.attrib["r"])
            cell.set("r", f"{col_letters(column)}{row_number}")
        sheet_data.append(row)

    last_row = header_row + len(data_rows)
    dimension = book.root.find(q("dimension"))
    if dimension is not None:
        dimension.set("ref", f"A1:E{last_row}")
    auto_filter = book.root.find(q("autoFilter"))
    if auto_filter is not None:
        auto_filter.set("ref", f"A1:E{last_row}")

    # Display value with thousand separators and volume with three decimals.
    # These are display formats only; stored numeric values keep full precision.
    book.apply_column_number_format((4,), header_row + 1, "#,##0.000")
    book.apply_column_number_format((5,), header_row + 1, "0.000")

    book.save()

    check = Book(MASTER, "داده‌های روزانه بازار")
    check_rows = check.rows()
    verified = {}
    for row_number, cells in check_rows.items():
        if row_number <= header_row:
            continue
        date_key = canonical_date(check.value(cells.get(1)))
        if date_key:
            verified[date_key] = cells

    missing = sorted(set(record["date"] for record in records) - set(verified))
    if missing:
        raise RuntimeError(f"JSON dates missing from unified workbook: {missing}")

    for record in records:
        date_key = record["date"]
        cells = verified[date_key]
        if float(check.value(cells.get(2))) != record["index"]:
            raise RuntimeError(f"{date_key}: شاخص کل mismatch")
        if float(check.value(cells.get(3))) != record["index_equalWeight"]:
            raise RuntimeError(f"{date_key}: شاخص هم‌وزن mismatch")
        expected_value = record["tval"] / BILLION
        expected_volume = record["tvol"] / BILLION
        actual_value = check.value(cells.get(4))
        actual_volume = check.value(cells.get(5))
        if not math.isclose(float(actual_value), expected_value, rel_tol=1e-12):
            raise RuntimeError(f"{date_key}: value B conversion mismatch")
        if not math.isclose(float(actual_volume), expected_volume, rel_tol=1e-12):
            raise RuntimeError(f"{date_key}: volume B conversion mismatch")
        if check.number_format(cells.get(4)) != "#,##0.000":
            raise RuntimeError(f"{date_key}: value display format is not #,##0.000")
        if check.number_format(cells.get(5)) != "0.000":
            raise RuntimeError(f"{date_key}: volume display format is not 0.000")

    newest_row = check_rows.get(header_row + 1, {})
    newest_date = canonical_date(check.value(newest_row.get(1)))
    if newest_date != max(verified):
        raise RuntimeError(
            f"Newest date is not directly below headers: {newest_date!r}; expected {max(verified)!r}"
        )

    print(f"Unified workbook: {MASTER}")
    print(f"JSON dates synchronized: {len(records)}")
    print(f"New rows inserted: {inserted or 'none'}")
    print(f"Existing rows refreshed with JSON indices: {refreshed or 'none'}")
    print(f"Newest date: {newest_date}")
    print("Value display format: #,##0.000; volume display format: 0.000; stored numeric precision preserved.")
    print("CHECK: all dates matched; both index columns match their daily JSON; value and volume remain populated and both display formats are verified.")


def main():
    records = load_records()
    print(f"Closed JSON reports found: {len(records)}")
    for record in records:
        print(
            f"{record['date']}: index={record['index']}, "
            f"equalWeight={record['index_equalWeight']}, "
            f"tval={record['tval'] / BILLION:.3f} B, "
            f"tvol={record['tvol'] / BILLION:.3f} B"
        )
    update_master(records)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        raise
