#!/usr/bin/env python3
"""Single scheduled BRSAPI-to-Excel engine for the Iranian daily market archive."""
from __future__ import annotations

import json
import csv
import os
import subprocess
import time
from datetime import datetime, timedelta
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo
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

def report_fingerprint(report):
    """Stable identity for one final daily report, ignoring JSON key order."""
    fields = (
        "date",
        "state",
        "time",
        "index",
        "index_change",
        "index_equalWeight",
        "index_equalWeight_change",
        "mv",
        "tno",
        "tval",
        "tvol",
    )
    return tuple((field, report.get(field)) for field in fields)

def load_records():
    found = {}
    fingerprints = {}
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

        fingerprint = report_values_fingerprint(obj)
        if d in found:
            if fingerprints[d] == fingerprint:
                print(
                    f"IGNORE DUPLICATE ARCHIVE COPY: {path.name} contains "
                    f"the same verified report for {d}."
                )
                continue
            raise RuntimeError(
                f"CONFLICTING REPORTS FOR SAME DATE {d}: archive {path.name} "
                "has different market values. Excel update is blocked for review."
            )
        fingerprints[d] = fingerprint
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



# ---------------------------------------------------------------------------
# BRSAPI request scheduler, durable state, and verified workbook completion
# ---------------------------------------------------------------------------

TEHRAN = ZoneInfo("Asia/Tehran")
API_URL = "https://api.brsapi.ir/Tsetmc/Index.php"
AUTO_STATE_PATH = Path("state/daily-engine-state.json")
MANUAL_STATE_PATH = Path("state/daily-engine-manual-state.json")
STATE_PATH = AUTO_STATE_PATH
MANUAL_MODE = False
MANUAL_TRIGGER_PATH = Path("state/manual-resume-trigger.json")
ARCHIVE_DIR = Path("archive")
WORKBOOK_PATH = Path("archive/داده‌های روزانه بازار.xlsx")
MAX_DAILY_REQUESTS = 85
MANUAL_MAX_DAILY_REQUESTS = 15
MAX_ATTEMPTS_PER_STAGE = 3
MIN_REQUEST_INTERVAL_SECONDS = 30 * 60

STAGE_NAMES = [
    "18:00",
    "21:00",
    "23:30",
    "02:00",
    "04:00",
    "06:00",
]

def configure_execution_scope(manual_resume):
    """Keep manual recovery state isolated from the automatic daily engine."""
    global STATE_PATH, MANUAL_MODE, MAX_DAILY_REQUESTS
    MANUAL_MODE = bool(manual_resume)
    if MANUAL_MODE:
        STATE_PATH = MANUAL_STATE_PATH
        MAX_DAILY_REQUESTS = MANUAL_MAX_DAILY_REQUESTS
        print(
            f"MANUAL ISOLATED STATE: {STATE_PATH}; "
            f"manual API cap={MAX_DAILY_REQUESTS}/Tehran day."
        )
    else:
        STATE_PATH = AUTO_STATE_PATH
        MAX_DAILY_REQUESTS = 85

def tehran_now():
    return datetime.now(TEHRAN)

def jalali_date_for(dt):
    y, m, d = gregorian_to_jalali(dt.year, dt.month, dt.day)
    return f"{y:04d}-{m:02d}-{d:02d}"

def parse_dt(value):
    if not value:
        return None
    result = datetime.fromisoformat(value)
    if result.tzinfo is None:
        result = result.replace(tzinfo=TEHRAN)
    return result.astimezone(TEHRAN)

def load_state():
    if not STATE_PATH.is_file():
        return {
            "schema_version": 1,
            "active_cycle": None,
            "last_completed_at": None,
            "last_cycle": None,
            "daily_usage": {},
        }

    state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    if not isinstance(state, dict) or state.get("schema_version") != 1:
        raise RuntimeError("Engine state file is invalid or has an unsupported schema.")
    if not isinstance(state.get("daily_usage", {}), dict):
        raise RuntimeError("Engine daily request counter is invalid.")
    if state.get("active_cycle") is not None and not isinstance(state["active_cycle"], dict):
        raise RuntimeError("Engine active-cycle state is invalid.")
    state.setdefault("active_cycle", None)
    state.setdefault("last_completed_at", None)
    state.setdefault("last_cycle", None)
    state.setdefault("daily_usage", {})
    return state

def save_state(state):
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    temporary = STATE_PATH.with_suffix(".json.tmp")
    temporary.write_text(
        json.dumps(state, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(STATE_PATH)

def git(*args, check=True):
    return subprocess.run(["git", *args], check=check, text=True)

def commit_and_push(paths, message):
    git("add", "--", *[str(path) for path in paths])
    staged = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        check=False,
    )
    if staged.returncode == 0:
        print("Git: no changes to commit.")
        return

    git("config", "user.name", "github-actions[bot]")
    git("config", "user.email", "41898282+github-actions[bot]@users.noreply.github.com")
    git("commit", "-m", message)

    # The workflow checks out the main branch explicitly. Rebase only if another
    # legitimate repository commit arrived while this job was running.
    try:
        git("pull", "--rebase", "origin", "main")
    except subprocess.CalledProcessError:
        subprocess.run(["git", "rebase", "--abort"], check=False)
        raise
    git("push", "origin", "main")
    git("fetch", "origin", "main")
    git("reset", "--hard", "origin/main")

def persist_state(state, message, extra_paths=None):
    save_state(state)
    paths = [STATE_PATH]
    if extra_paths:
        paths.extend(extra_paths)
    commit_and_push(paths, message)

def valid_report(data, expected_date):
    if not isinstance(data, dict):
        return False, "response is not a JSON object"

    if data.get("date") != expected_date:
        return False, f"report date {data.get('date')!r} != expected {expected_date}"

    if data.get("state") != "بسته":
        return False, f"market state is not closed: {data.get('state')!r}"

    report_time = data.get("time")
    if (
        not isinstance(report_time, str)
        or not re.fullmatch(r"\d{2}:\d{2}:\d{2}", report_time)
        or report_time < "17:00:00"
    ):
        return False, f"invalid or premature report time: {report_time!r}"

    required_numeric = (
        "index",
        "index_change",
        "index_equalWeight",
        "index_equalWeight_change",
        "mv",
        "tno",
        "tval",
        "tvol",
    )
    for field in required_numeric:
        value = data.get(field)
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(float(value))
        ):
            return False, f"missing or invalid numeric field: {field}"

    for field in ("mv", "tno", "tval", "tvol"):
        if float(data[field]) <= 0:
            return False, f"non-positive required market field: {field}"

    return True, "valid"

def read_archived_report(expected_date):
    path = ARCHIVE_DIR / f"index-{expected_date}.json"
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        ok, reason = valid_report(data, expected_date)
        if ok:
            return data
        print(f"Existing archive is not usable: {reason}")
    except Exception as exc:
        print(f"Existing archive cannot be read: {type(exc).__name__}: {exc}")
    return None

def report_values_fingerprint(report):
    """Market-value identity independent of the reported date and timestamp."""
    fields = (
        "index",
        "index_change",
        "index_equalWeight",
        "index_equalWeight_change",
        "mv",
        "tno",
        "tval",
        "tvol",
    )
    values = []
    for field in fields:
        value = report.get(field)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return None
        if not math.isfinite(float(value)):
            return None
        values.append((field, float(value)))
    return tuple(values)

def find_duplicate_archived_report(report):
    """A report is duplicate only when its own market date is already archived."""
    if not isinstance(report, dict):
        return None

    response_date = canonical_date(report.get("date"))
    if not response_date:
        return None

    same_date_path = ARCHIVE_DIR / f"index-{response_date}.json"
    if not same_date_path.is_file():
        return None

    try:
        archived = json.loads(same_date_path.read_text(encoding="utf-8-sig"))
        archived_ok, _ = valid_report(archived, response_date)
        if archived_ok:
            return (
                f"report date {response_date} already exists in "
                f"{same_date_path.as_posix()}"
            )
    except Exception:
        return None

    return None

def is_start_day(now):
    # Python weekday(): Monday=0 ... Sunday=6.
    # Iran trading days: Saturday through Wednesday.
    if now.weekday() not in {5, 6, 0, 1, 2}:
        print("START SKIPPED: Thursday or Friday in Iran.")
        return False

    jalali_date = jalali_date_for(now)
    jalali_year = int(jalali_date.split("-")[0])
    calendar_path = Path(f"calendar/holidays-{jalali_year}.json")
    if not calendar_path.is_file():
        raise RuntimeError(f"Official Iranian holiday calendar is missing: {calendar_path}")

    calendar_data = json.loads(calendar_path.read_text(encoding="utf-8"))
    if calendar_data.get("year") != jalali_year:
        raise RuntimeError(f"Holiday calendar year mismatch: {calendar_path}")

    holidays = set(calendar_data.get("holidays", []))
    if jalali_date in holidays:
        print(f"START SKIPPED: {jalali_date} is an official Iranian holiday.")
        return False

    return True

def stage_for_now(now, cycle_start_date):
    """Return stage index for the active cycle, independent of today's weekday."""
    day_offset = (now.date() - cycle_start_date).days
    minute_of_day = now.hour * 60 + now.minute

    if day_offset == 0:
        if 18 * 60 <= minute_of_day < 21 * 60:
            return 0
        if 21 * 60 <= minute_of_day < 23 * 60 + 30:
            return 1
        if minute_of_day >= 23 * 60 + 30:
            return 2
        return None

    if day_offset == 1:
        if minute_of_day < 2 * 60:
            return 2
        if 2 * 60 <= minute_of_day < 4 * 60:
            return 3
        if 4 * 60 <= minute_of_day < 6 * 60:
            return 4
        # Three final attempts: 06:00, 06:30, and 07:00.
        # A short execution-delay grace is allowed, but no API call after 07:15.
        if 6 * 60 <= minute_of_day <= 7 * 60 + 15:
            return 5

    return None

def cycle_start_date(cycle):
    started = parse_dt(cycle.get("started_at"))
    if started is None:
        raise RuntimeError("Active cycle has no valid start timestamp.")
    return started.date()

def manual_resume_allowed(now, cycle):
    """Manual resume is limited to the active previous-day cycle, before 07:30."""
    if not cycle or cycle.get("status") not in {"fetching", "pending_excel"}:
        return False
    start_date = cycle_start_date(cycle)
    deadline = datetime.combine(
        start_date + timedelta(days=1),
        datetime.min.time(),
        tzinfo=TEHRAN,
    ) + timedelta(hours=7, minutes=30)
    return now.date() == start_date + timedelta(days=1) and now < deadline

def read_manual_trigger():
    if not MANUAL_TRIGGER_PATH.is_file():
        return None
    try:
        payload = json.loads(MANUAL_TRIGGER_PATH.read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(
            f"Manual resume trigger is not valid JSON: {type(exc).__name__}: {exc}"
        ) from exc
    if (
        not isinstance(payload, dict)
        or payload.get("action") != "resume_active_cycle"
        or not isinstance(payload.get("request_id"), str)
        or not payload["request_id"].strip()
    ):
        raise RuntimeError("Manual resume trigger has an invalid structure.")
    return payload

def close_failed_cycle(state, now, reason):
    cycle = state.get("active_cycle")
    if not cycle:
        return

    cycle["status"] = "failed"
    cycle["closed_at"] = now.isoformat()
    cycle["failure_reason"] = reason
    state["last_cycle"] = {
        "market_date": cycle.get("market_date"),
        "status": "failed",
        "started_at": cycle.get("started_at"),
        "closed_at": cycle.get("closed_at"),
        "stage_attempts": cycle.get("stage_attempts", [0] * 6),
        "failure_reason": reason,
    }
    state["active_cycle"] = None
    persist_state(state, f"Close failed BRSAPI cycle {cycle.get('market_date')}")

def close_duplicate_cycle(state, now, reason):
    """End this trading-day cycle immediately without storing a new report."""
    cycle = state.get("active_cycle")
    if not cycle:
        return

    cycle["status"] = "duplicate"
    cycle["closed_at"] = now.isoformat()
    cycle["duplicate_reason"] = reason
    state["last_cycle"] = {
        "market_date": cycle.get("market_date"),
        "status": "duplicate",
        "started_at": cycle.get("started_at"),
        "closed_at": cycle.get("closed_at"),
        "stage_attempts": cycle.get("stage_attempts", [0] * 6),
        "duplicate_reason": reason,
        "new_report_registered": False,
    }
    state["active_cycle"] = None
    persist_state(
        state,
        f"Close duplicate BRSAPI cycle {cycle.get('market_date')}",
    )
    print(
        f"DUPLICATE REPORT: cycle {cycle.get('market_date')} stopped immediately. "
        "No new JSON or Excel row was registered; no further API attempts will occur."
    )

def record_failure(state, cycle, now, result, detail, http_status=None):
    cycle["last_attempt"] = {
        "at": now.isoformat(),
        "result": result,
        "detail": str(detail)[:500],
        "http_status": http_status,
    }
    persist_state(
        state,
        f"Record BRSAPI attempt {cycle.get('market_date')} ({result})",
    )

def request_brsapi(expected_date):
    api_key = os.environ.get("BRSAPI_KEY", "").strip()
    if not api_key:
        raise RuntimeError("Missing GitHub Actions secret BRSAPI_KEY.")

    query = urlencode({"key": api_key, "type": "1"})
    request = Request(
        f"{API_URL}?{query}",
        headers={
            "Accept": "application/json",
            "User-Agent": "bourse-daily-data-daily-engine/2.0",
        },
    )

    try:
        with urlopen(request, timeout=45) as response:
            http_status = response.status
            body = response.read()

        if http_status != 200:
            return None, None, http_status, f"unexpected HTTP status {http_status}"

        try:
            data = json.loads(body.decode("utf-8-sig"))
        except Exception as exc:
            return None, None, http_status, f"invalid JSON: {type(exc).__name__}: {exc}"

        duplicate_reason = find_duplicate_archived_report(data)
        if duplicate_reason:
            return None, None, http_status, f"DUPLICATE: {duplicate_reason}"

        ok, reason = valid_report(data, expected_date)
        if not ok:
            return None, None, http_status, reason

        return data, body, http_status, "valid final report"

    except HTTPError as exc:
        body = exc.read()
        detail = f"HTTP {exc.code}"
        try:
            error_data = json.loads(body.decode("utf-8-sig"))
            detail = (
                f"HTTP {exc.code}; "
                f"{error_data.get('status', '')}; "
                f"{error_data.get('message_error', '')}; "
                f"usage_today={error_data.get('account', {}).get('usage_today')}"
            )
        except Exception:
            pass
        return None, None, exc.code, detail

    except (URLError, TimeoutError, OSError) as exc:
        return None, None, None, f"{type(exc).__name__}: {exc}"

def verify_report_in_workbook(report):
    if not WORKBOOK_PATH.is_file():
        raise RuntimeError(f"Unified workbook is missing: {WORKBOOK_PATH}")

    book = Book(WORKBOOK_PATH, "داده‌های روزانه بازار")
    rows = book.rows()
    target_date = report["date"]
    matches = []

    for row_number, cells in rows.items():
        if row_number <= 1:
            continue
        date_value = canonical_date(book.value(cells.get(1)))
        if date_value == target_date:
            matches.append((row_number, cells))

    if len(matches) != 1:
        raise RuntimeError(
            f"Workbook must contain exactly one row for {target_date}; found {len(matches)}"
        )

    row_number, cells = matches[0]
    if row_number != 2:
        raise RuntimeError(
            f"Latest report {target_date} is not in row 2; it is in row {row_number}"
        )

    expected = {
        2: float(report["index"]),
        3: float(report["index_equalWeight"]),
        4: float(report["tval"]) / BILLION,
        5: float(report["tvol"]) / BILLION,
    }

    for column, expected_value in expected.items():
        actual_value = book.value(cells.get(column))
        if actual_value is None or not math.isclose(
            float(actual_value),
            expected_value,
            rel_tol=1e-12,
            abs_tol=1e-9,
        ):
            raise RuntimeError(
                f"Workbook mismatch for {target_date}, column {column}: "
                f"{actual_value!r} != {expected_value!r}"
            )

    if book.number_format(cells.get(4)) != "#,##0.000":
        raise RuntimeError("Workbook trading-value number format is incorrect.")
    if book.number_format(cells.get(5)) != "0.000":
        raise RuntimeError("Workbook trading-volume number format is incorrect.")

    print(
        f"WORKBOOK VERIFIED: {target_date} is in row 2; "
        "index, equal-weight index, value, volume, and display formats match JSON."
    )

def merge_manual_history():
    """Insert verified manual history only for dates not already in the workbook."""
    source = Path("archive/manual-history.csv")
    if not source.is_file():
        print("Manual history CSV not present; no manual rows to merge.")
        return []

    records = {}
    with source.open("r", encoding="utf-8-sig", newline="") as handle:
        for item in csv.DictReader(handle):
            date_key = canonical_date(item.get("date"))
            if not date_key:
                raise RuntimeError(f"Invalid manual-history date: {item.get('date')!r}")
            if date_key in records:
                raise RuntimeError(f"Duplicate manual-history date: {date_key}")

            numeric = {}
            for field in ("index", "equal_weight", "value_billion_rial"):
                try:
                    number = float(item.get(field))
                except (TypeError, ValueError):
                    raise RuntimeError(f"Invalid manual-history field {field} for {date_key}")
                if not math.isfinite(number) or number <= 0:
                    raise RuntimeError(f"Non-positive/invalid manual-history field {field} for {date_key}")
                numeric[field] = number

            volume_text = (item.get("volume_billion_shares") or "").strip()
            volume = None
            if volume_text:
                try:
                    volume = float(volume_text)
                except ValueError:
                    raise RuntimeError(f"Invalid manual volume for {date_key}")
                if not math.isfinite(volume) or volume <= 0:
                    raise RuntimeError(f"Non-positive/invalid manual volume for {date_key}")

            records[date_key] = {
                "index": numeric["index"],
                "equal_weight": numeric["equal_weight"],
                "value": numeric["value_billion_rial"],
                "volume": volume,
            }

    if not records:
        return []

    book = Book(WORKBOOK_PATH, "داده‌های روزانه بازار")
    rows = book.rows()
    header = rows.get(1, {})
    for column, label in {
        1: "تاریخ (شمسی)",
        2: "شاخص کل",
        3: "شاخص هم‌وزن",
    }.items():
        if book.value(header.get(column)) != label:
            raise RuntimeError(f"Unexpected workbook header in column {column}")

    existing = {}
    for row_number, cells in rows.items():
        if row_number <= 1:
            continue
        date_key = canonical_date(book.value(cells.get(1)))
        if date_key:
            if date_key in existing:
                raise RuntimeError(f"Duplicate workbook date: {date_key}")
            existing[date_key] = row_number

    inserted = []
    for date_key in sorted(records):
        if date_key in existing:
            print(f"KEEP EXISTING WORKBOOK ROW: {date_key}")
            continue

        item = records[date_key]
        values = {
            1: date_key,
            2: item["index"],
            3: item["equal_weight"],
            4: item["value"],
        }
        if item["volume"] is not None:
            values[5] = item["volume"]
        book.insert(1, values, {})
        inserted.append(date_key)

    if not inserted:
        print(f"Manual history checked: {len(records)} dates already represented.")
        return []

    sheet_data = book.root.find(q("sheetData"))
    row_nodes = list(sheet_data.findall(q("row")))
    header_node = next(
        (row for row in row_nodes if int(row.attrib.get("r", "0")) == 1),
        None,
    )
    if header_node is None:
        raise RuntimeError("Workbook header row 1 is missing.")

    data_rows = []
    for row in row_nodes:
        row_number = int(row.attrib.get("r", "0"))
        if row_number <= 1:
            continue
        cells = row_cells(row)
        date_key = canonical_date(book.value(cells.get(1)))
        if not date_key:
            if any(book.value(cell) not in (None, "") for cell in cells.values()):
                raise RuntimeError(f"Workbook row {row_number} has values but no valid date.")
            continue
        data_rows.append((date_key, row))

    data_rows.sort(key=lambda pair: pair[0], reverse=True)
    for child in list(sheet_data):
        sheet_data.remove(child)
    sheet_data.append(header_node)

    for row_number, (_, row) in enumerate(data_rows, start=2):
        row.set("r", str(row_number))
        for cell in row.findall(q("c")):
            column = col_number(cell.attrib["r"])
            cell.set("r", f"{col_letters(column)}{row_number}")
        sheet_data.append(row)

    last_row = 1 + len(data_rows)
    dimension = book.root.find(q("dimension"))
    if dimension is not None:
        dimension.set("ref", f"A1:E{last_row}")
    auto_filter = book.root.find(q("autoFilter"))
    if auto_filter is not None:
        auto_filter.set("ref", f"A1:E{last_row}")

    book.apply_column_number_format((4,), 2, "#,##0.000")
    book.apply_column_number_format((5,), 2, "0.000")
    book.save()

    check = Book(WORKBOOK_PATH, "داده‌های روزانه بازار")
    checked_rows = check.rows()
    verified = {}
    for row_number, cells in checked_rows.items():
        if row_number > 1:
            date_key = canonical_date(check.value(cells.get(1)))
            if date_key:
                verified[date_key] = cells

    for date_key in inserted:
        cells = verified.get(date_key)
        if cells is None:
            raise RuntimeError(f"Manual history date missing after save: {date_key}")
        item = records[date_key]
        expected_values = (
            item["index"],
            item["equal_weight"],
            item["value"],
        )
        actual_values = (
            check.value(cells.get(2)),
            check.value(cells.get(3)),
            check.value(cells.get(4)),
        )
        for actual, expected in zip(actual_values, expected_values):
            if actual is None or not math.isclose(
                float(actual), expected, rel_tol=1e-12, abs_tol=1e-9
            ):
                raise RuntimeError(f"Manual-history verification failed for {date_key}")
        if item["volume"] is not None:
            actual_volume = check.value(cells.get(5))
            if actual_volume is None or not math.isclose(
                float(actual_volume), item["volume"], rel_tol=1e-12, abs_tol=1e-9
            ):
                raise RuntimeError(f"Manual-history volume verification failed for {date_key}")

    print(f"Manual history rows inserted and verified: {len(inserted)}")
    return inserted

def process_pending_excel(state):
    cycle = state.get("active_cycle")
    if not cycle or cycle.get("status") != "pending_excel":
        return False

    market_date = cycle.get("market_date")
    report = read_archived_report(market_date)
    if report is None:
        message = (
            f"Cycle {market_date} is pending Excel, but its archived JSON "
            "is missing or no longer valid. BRSAPI will not be called again."
        )
        cycle["last_excel_error"] = message
        cycle["last_excel_attempt_at"] = tehran_now().isoformat()
        persist_state(state, f"Record pending Excel error {market_date}")
        print(message)
        return False

    try:
        print(f"EXCEL PHASE ONLY: using archive/index-{market_date}.json; no API request.")
        records = load_records()
        update_master(records)
        merge_manual_history()
        verify_report_in_workbook(report)

        cycle["last_excel_verified_at"] = tehran_now().isoformat()
        cycle["last_excel_error"] = None
        save_state(state)
        commit_and_push(
            [WORKBOOK_PATH, STATE_PATH],
            f"Update and verify daily market workbook {market_date}",
        )

        # Verify the workbook again after the GitHub push, from the checked-out
        # main branch version, before declaring the full cycle complete.
        verify_report_in_workbook(report)

        completed_at = tehran_now()
        cycle["status"] = "completed"
        cycle["completed_at"] = completed_at.isoformat()
        state["last_completed_at"] = completed_at.isoformat()
        state["last_cycle"] = {
            "market_date": market_date,
            "status": "completed",
            "started_at": cycle.get("started_at"),
            "completed_at": completed_at.isoformat(),
            "stage_attempts": cycle.get("stage_attempts", [0] * 6),
            "report_path": f"archive/index-{market_date}.json",
            "workbook_path": str(WORKBOOK_PATH),
        }
        state["active_cycle"] = None
        persist_state(state, f"Mark BRSAPI and Excel cycle complete {market_date}")
        print(
            f"FULL SUCCESS: JSON {market_date} archived; workbook updated, "
            "pushed, and verified. No more BRSAPI requests for this cycle."
        )
        return True

    except Exception as exc:
        cycle["last_excel_attempt_at"] = tehran_now().isoformat()
        cycle["last_excel_error"] = f"{type(exc).__name__}: {exc}"
        persist_state(state, f"Record Excel-only retry {market_date}")
        print(
            f"EXCEL UPDATE/VERIFY FAILED: {type(exc).__name__}: {exc}. "
            "The next scheduled run will retry Excel only; BRSAPI is locked."
        )
        return False

def attempt_api_once(state, now, stage_index):
    cycle = state["active_cycle"]
    market_date = cycle["market_date"]

    stage_attempts = cycle.setdefault("stage_attempts", [0] * 6)
    if len(stage_attempts) != 6:
        raise RuntimeError("Active cycle stage counter is invalid.")

    if stage_attempts[stage_index] >= MAX_ATTEMPTS_PER_STAGE:
        print(f"STAGE {STAGE_NAMES[stage_index]}: three requests already used.")
        return

    last_request = parse_dt(cycle.get("last_request_at"))
    if last_request is not None:
        elapsed = (now - last_request).total_seconds()
        if elapsed < MIN_REQUEST_INTERVAL_SECONDS:
            print(
                f"RETRY DEFERRED: only {elapsed:.0f} seconds since the previous "
                "request; minimum interval is 1800 seconds."
            )
            return

    today_key = now.date().isoformat()
    usage = state.setdefault("daily_usage", {})
    used_today = int(usage.get(today_key, 0))
    if used_today >= MAX_DAILY_REQUESTS:
        print(
            f"DAILY LIMIT REACHED: {used_today}/{MAX_DAILY_REQUESTS} "
            f"requests on Tehran date {today_key}; no request sent."
        )
        return False

    if not os.environ.get("BRSAPI_KEY", "").strip():
        raise RuntimeError("Missing GitHub Actions secret BRSAPI_KEY.")

    # Persist the quota debit and attempt slot BEFORE the network call. If the
    # runner is cancelled or crashes, a later run cannot repeat this slot.
    usage[today_key] = used_today + 1
    stage_attempts[stage_index] += 1
    cycle["last_request_at"] = now.isoformat()
    attempt_number = stage_attempts[stage_index]
    attempt_id = f"{market_date}-stage-{stage_index + 1}-attempt-{attempt_number}"
    cycle["last_attempt"] = {
        "id": attempt_id,
        "at": now.isoformat(),
        "stage": STAGE_NAMES[stage_index],
        "stage_attempt": attempt_number,
        "result": "request_started",
        "http_status": None,
    }
    persist_state(
        state,
        f"Reserve BRSAPI request {attempt_id}",
    )

    print(
        f"BRSAPI REQUEST {attempt_id}; "
        f"Tehran day usage {usage[today_key]}/{MAX_DAILY_REQUESTS}."
    )

    data, body, http_status, detail = request_brsapi(market_date)
    if data is None and detail.startswith("DUPLICATE:"):
        close_duplicate_cycle(
            state,
            tehran_now(),
            detail.removeprefix("DUPLICATE:").strip(),
        )
        return False

    if data is None:
        record_failure(
            state,
            cycle,
            tehran_now(),
            "failed_or_invalid",
            detail,
            http_status,
        )
        print(
            f"BRSAPI attempt failed: HTTP={http_status}; {detail}. "
            "The next scheduled attempt follows the 30-minute rule."
        )
        return False

    archive_path = ARCHIVE_DIR / f"index-{market_date}.json"

    # Recheck the canonical archive after the network response. A verified
    # same-date report is authoritative; never overwrite or append it twice.
    existing_report = read_archived_report(market_date)
    if existing_report is not None:
        identical = report_values_fingerprint(existing_report) == report_values_fingerprint(data)
        cycle["status"] = "pending_excel"
        cycle["report_path"] = str(archive_path)
        cycle["report_validated_at"] = tehran_now().isoformat()
        cycle["last_attempt"] = {
            "id": attempt_id,
            "at": now.isoformat(),
            "stage": STAGE_NAMES[stage_index],
            "stage_attempt": attempt_number,
            "result": (
                "duplicate_report_reused"
                if identical
                else "same_date_archive_kept_api_payload_not_overwritten"
            ),
            "http_status": http_status,
        }
        save_state(state)
        commit_and_push(
            [STATE_PATH],
            f"Reuse existing BRSAPI archive {market_date}",
        )
        if identical:
            print(
                f"DUPLICATE REPORT: {market_date} already exists with identical "
                "content. Existing JSON is reused; no second archive row is created."
            )
        else:
            print(
                f"SAME-DATE CONFLICT: an existing verified JSON for {market_date} "
                "is retained as authoritative; the new API payload is not written."
            )
        return process_pending_excel(state)

    archive_path.parent.mkdir(parents=True, exist_ok=True)
    archive_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    cycle["status"] = "pending_excel"
    cycle["report_path"] = str(archive_path)
    cycle["report_validated_at"] = tehran_now().isoformat()
    cycle["last_attempt"] = {
        "id": attempt_id,
        "at": now.isoformat(),
        "stage": STAGE_NAMES[stage_index],
        "stage_attempt": attempt_number,
        "result": "valid_json_archived",
        "http_status": http_status,
    }
    save_state(state)
    commit_and_push(
        [archive_path, STATE_PATH],
        f"Archive complete BRSAPI report {market_date}",
    )
    print(f"VALID JSON SAVED: {archive_path}")
    return process_pending_excel(state)

def start_new_cycle(state, now):
    # A missed 18:00 scheduled event may be recovered at any later scheduled
    # tick that evening. The six-stage timetable remains anchored to 18:00.
    if now.hour < 18:
        return False

    market_date = jalali_date_for(now)
    last_cycle = state.get("last_cycle") or {}
    if last_cycle.get("market_date") == market_date:
        print(
            f"CYCLE ALREADY CLOSED FOR {market_date}: "
            f"status={last_cycle.get('status')!r}; no second cycle will start today."
        )
        return False

    if not is_start_day(now):
        return False

    scheduled_start = now.replace(hour=18, minute=0, second=0, microsecond=0)
    cycle = {
        "market_date": market_date,
        "started_at": scheduled_start.isoformat(),
        "created_at": now.isoformat(),
        "status": "fetching",
        "stage_attempts": [0, 0, 0, 0, 0, 0],
        "last_request_at": None,
        "last_attempt": None,
        "report_path": None,
    }
    state["active_cycle"] = cycle
    existing = read_archived_report(market_date)
    if existing is not None:
        cycle["status"] = "pending_excel"
        cycle["report_path"] = str(ARCHIVE_DIR / f"index-{market_date}.json")
        print(
            f"VALID EXISTING JSON FOUND for {market_date}; "
            "skip BRSAPI and repair/verify Excel only."
        )
    persist_state(state, f"Start daily BRSAPI cycle {market_date}")

    return True


def recover_automatic_previous_day_cycle(state, now):
    """Recover a missed automatic cycle during its remaining overnight window.

    Returns the created cycle, None when no eligible recovery applies, or False
    when the recovery deadline has passed and a trading day was missed.
    """
    if now.hour >= 18:
        return None

    previous_day = now.date() - timedelta(days=1)
    scheduled_start = datetime.combine(
        previous_day,
        datetime.min.time(),
        tzinfo=TEHRAN,
    ).replace(hour=18)
    market_date = jalali_date_for(scheduled_start)

    if not is_start_day(scheduled_start):
        print(f"AUTOMATIC RECOVERY: {market_date} was not an eligible trading day.")
        return None

    last_cycle = state.get("last_cycle") or {}
    if last_cycle.get("market_date") == market_date:
        print(
            f"AUTOMATIC RECOVERY: {market_date} was already closed with "
            f"status {last_cycle.get('status')!r}."
        )
        return False if last_cycle.get("status") == "failed" else None

    deadline = datetime.combine(
        now.date(),
        datetime.min.time(),
        tzinfo=TEHRAN,
    ) + timedelta(hours=7, minutes=30)

    if now >= deadline:
        closed_at = now.isoformat()
        state["last_cycle"] = {
            "market_date": market_date,
            "status": "failed",
            "started_at": scheduled_start.isoformat(),
            "closed_at": closed_at,
            "stage_attempts": [0, 0, 0, 0, 0, 0],
            "failure_reason": (
                "The automatic cycle was not started and its 07:30 Tehran "
                "recovery deadline has passed."
            ),
        }
        persist_state(
            state,
            f"Record missed automatic BRSAPI cycle {market_date}",
        )
        print(
            f"AUTOMATIC CYCLE MISSED: {market_date}; recovery deadline "
            "07:30 Tehran has passed. No API request was sent."
        )
        return False

    cycle = {
        "market_date": market_date,
        "started_at": scheduled_start.isoformat(),
        "created_at": now.isoformat(),
        "status": "fetching",
        "stage_attempts": [0, 0, 0, 0, 0, 0],
        "last_request_at": None,
        "last_attempt": None,
        "report_path": None,
    }
    existing = read_archived_report(market_date)
    if existing is not None:
        cycle["status"] = "pending_excel"
        cycle["report_path"] = str(ARCHIVE_DIR / f"index-{market_date}.json")
        print(
            f"VALID EXISTING JSON FOUND for {market_date}; "
            "automatic recovery will repair/verify Excel only."
        )

    state["active_cycle"] = cycle
    persist_state(state, f"Automatically recover missed BRSAPI cycle {market_date}")
    print(
        f"AUTOMATIC PREVIOUS-DAY CYCLE RECOVERED: {market_date}; "
        "continuing with the original overnight retry timetable."
    )
    return cycle


def start_manual_previous_day_cycle(state, now):
    """Recover yesterday's missed 18:00 cycle in the isolated manual state file."""
    deadline = datetime.combine(
        now.date(),
        datetime.min.time(),
        tzinfo=TEHRAN,
    ) + timedelta(hours=7, minutes=30)
    if now < datetime.combine(now.date(), datetime.min.time(), tzinfo=TEHRAN) or now >= deadline:
        print("MANUAL RECOVERY BLOCKED: this command is valid only from 00:00 through 07:29:59 Tehran.")
        return None

    previous_day = now.date() - timedelta(days=1)
    scheduled_start = datetime.combine(
        previous_day,
        datetime.min.time(),
        tzinfo=TEHRAN,
    ).replace(hour=18)
    market_date = jalali_date_for(scheduled_start)

    last_cycle = state.get("last_cycle") or {}
    if last_cycle.get("market_date") == market_date:
        print(
            f"MANUAL RECOVERY BLOCKED: a cycle for {market_date} is already closed "
            f"with status {last_cycle.get('status')!r}."
        )
        return None

    if not is_start_day(scheduled_start):
        print(f"MANUAL RECOVERY BLOCKED: {market_date} was not an eligible trading day.")
        return None

    cycle = {
        "market_date": market_date,
        "started_at": scheduled_start.isoformat(),
        "created_at": now.isoformat(),
        "status": "fetching",
        "stage_attempts": [0, 0, 0, 0, 0, 0],
        "last_request_at": None,
        "last_attempt": None,
        "report_path": None,
    }
    existing = read_archived_report(market_date)
    if existing is not None:
        cycle["status"] = "pending_excel"
        cycle["report_path"] = str(ARCHIVE_DIR / f"index-{market_date}.json")
        print(
            f"VALID EXISTING JSON FOUND for {market_date}; "
            "manual recovery will repair/verify Excel only."
        )

    state["active_cycle"] = cycle
    persist_state(state, f"Manually recover previous-day BRSAPI cycle {market_date}")
    print(
        f"MANUAL PREVIOUS-DAY CYCLE CREATED: {market_date}; "
        "manual state is isolated from the automatic cycle; "
        "manual quota cap and 30-minute controls remain active."
    )
    return cycle

def main():
    now = tehran_now()
    print(
        f"Tehran time: {now.isoformat()}; "
        f"Jalali date: {jalali_date_for(now)}; weekday={now.strftime('%A')}"
    )

    event_name = os.environ.get("GITHUB_EVENT_NAME", "")
    dispatch_resume = (
        event_name == "workflow_dispatch"
        and os.environ.get("MANUAL_RESUME_ACTIVE_CYCLE", "").lower() == "true"
    )
    trigger = read_manual_trigger() if event_name == "push" else None
    push_resume = trigger is not None
    manual_resume = dispatch_resume or push_resume
    configure_execution_scope(manual_resume)

    if event_name != "schedule" and not manual_resume:
        print("DIAGNOSTIC ONLY: no explicit manual-resume command was supplied.")
        state = load_state()
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 0

    state = load_state()

    if manual_resume:
        cycle = state.get("active_cycle")
        trigger_id = trigger.get("request_id") if trigger else None
        already_consumed = (
            trigger_id is not None
            and state.get("last_manual_trigger_id") == trigger_id
        )

        if push_resume:
            # Consume the one-shot command durably before any network request.
            state["last_manual_trigger_id"] = trigger_id
            MANUAL_TRIGGER_PATH.unlink(missing_ok=True)
            persist_state(
                state,
                f"Consume manual resume command {trigger_id}",
                extra_paths=[MANUAL_TRIGGER_PATH],
            )

        if already_consumed:
            print(f"MANUAL RESUME IGNORED: command {trigger_id} was already consumed.")
            return 0
        if not cycle:
            cycle = start_manual_previous_day_cycle(state, now)
        if not cycle:
            print("MANUAL RESUME BLOCKED: no eligible previous-day cycle could be recovered.")
            return 1
        if not manual_resume_allowed(now, cycle):
            print("MANUAL RESUME BLOCKED: only an active previous-day cycle may be resumed before 07:30 Tehran.")
            return 1
        print(
            f"MANUAL RESUME ENABLED: continuing only isolated manual cycle "
            f"{cycle.get('market_date')}; automatic state and schedule are untouched."
        )
    usage = state.setdefault("daily_usage", {})
    cutoff_date = (now.date() - timedelta(days=10)).isoformat()
    state["daily_usage"] = {
        day: count for day, count in usage.items() if day >= cutoff_date
    }

    cycle = state.get("active_cycle")

    # If a complete JSON has already been archived, lock out BRSAPI and finish
    # the workbook phase, even if an earlier run stopped between these steps.
    if cycle and cycle.get("status") == "fetching":
        archived = read_archived_report(cycle.get("market_date"))
        if archived is not None:
            cycle["status"] = "pending_excel"
            cycle["report_path"] = str(
                ARCHIVE_DIR / f"index-{cycle['market_date']}.json"
            )
            persist_state(
                state,
                f"Recover archived JSON and continue Excel {cycle['market_date']}",
            )

    cycle = state.get("active_cycle")
    if cycle and cycle.get("status") == "pending_excel":
        return 0 if process_pending_excel(state) else 1

    cycle = state.get("active_cycle")
    if cycle:
        start_date = cycle_start_date(cycle)
        deadline = datetime.combine(
            start_date + timedelta(days=1),
            datetime.min.time(),
            tzinfo=TEHRAN,
        ) + timedelta(hours=7, minutes=30)

        if now >= deadline:
            close_failed_cycle(
                state,
                now,
                "No complete JSON and verified workbook by 07:30 Tehran.",
            )
            return 1

    if cycle is None:
        if manual_resume:
            print("MANUAL RESUME BLOCKED: active cycle is already closed; no new cycle will be started manually.")
            return 1
        if now.hour < 18:
            recovered = recover_automatic_previous_day_cycle(state, now)
            if recovered is False:
                return 1
            if recovered is None:
                return 0
            cycle = state.get("active_cycle")
        else:
            if not start_new_cycle(state, now):
                return 0
            cycle = state.get("active_cycle")

        if cycle and cycle.get("status") == "pending_excel":
            return 0 if process_pending_excel(state) else 1

    cycle = state.get("active_cycle")
    if not cycle or cycle.get("status") != "fetching":
        return 0

    stage_index = stage_for_now(now, cycle_start_date(cycle))
    if stage_index is None:
        print("No API request is due at this Tehran time.")
        return 0

    result = attempt_api_once(state, now, stage_index)
    return 1 if result is False else 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"FATAL: {type(exc).__name__}: {exc}", file=sys.stderr)
        raise
