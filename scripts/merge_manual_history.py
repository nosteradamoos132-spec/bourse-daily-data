#!/usr/bin/env python3
"""Merge manually verified historical rows without overwriting existing dates."""
from __future__ import annotations

import csv
import math
import sys
from pathlib import Path

from update_index_excels import Book, MASTER, canonical_date, q, col_number, col_letters

SOURCE = Path("archive/manual-history.csv")
SHEET = "داده‌های روزانه بازار"
HEADER_ROW = 1

def load_manual_rows():
    records = {}
    with SOURCE.open("r", encoding="utf-8-sig", newline="") as handle:
        for item in csv.DictReader(handle):
            d = canonical_date(item.get("date"))
            if not d:
                raise RuntimeError(f"Invalid date in {SOURCE}: {item.get('date')!r}")
            if d in records:
                raise RuntimeError(f"Duplicate date in manual history: {d}")
            values = {}
            for field in ("index", "equal_weight", "value_billion_rial"):
                try:
                    number = float(item[field])
                except (TypeError, ValueError):
                    raise RuntimeError(f"Invalid {field} for {d}")
                if not math.isfinite(number):
                    raise RuntimeError(f"Non-finite {field} for {d}")
                values[field] = number
            volume_text = (item.get("volume_billion_shares") or "").strip()
            volume = None
            if volume_text:
                volume = float(volume_text)
                if not math.isfinite(volume):
                    raise RuntimeError(f"Non-finite volume for {d}")
            records[d] = {
                "index": values["index"],
                "equal_weight": values["equal_weight"],
                "value": values["value_billion_rial"],
                "volume": volume,
            }
    return records

def main():
    if not MASTER.is_file():
        raise RuntimeError(f"Unified workbook missing: {MASTER}")
    records = load_manual_rows()
    book = Book(MASTER, SHEET)
    rows = book.rows()
    header = rows.get(HEADER_ROW, {})
    expected = {
        1: "تاریخ (شمسی)",
        2: "شاخص کل",
        3: "شاخص هم‌وزن",
    }
    for column, label in expected.items():
        if book.value(header.get(column)) != label:
            raise RuntimeError(f"Unexpected workbook header in column {column}")
    existing = {}
    for row_number, cells in rows.items():
        if row_number <= HEADER_ROW:
            continue
        d = canonical_date(book.value(cells.get(1)))
        if d:
            if d in existing:
                raise RuntimeError(f"Duplicate workbook date: {d}")
            existing[d] = row_number

    inserted = []
    for d in sorted(records):
        if d in existing:
            print(f"KEEP EXISTING ROW UNCHANGED: {d}")
            continue
        item = records[d]
        values = {
            1: d,
            2: item["index"],
            3: item["equal_weight"],
            4: item["value"],
        }
        if item["volume"] is not None:
            values[5] = item["volume"]
        book.insert(HEADER_ROW, values, {})
        inserted.append(d)

    if inserted:
        sheet_data = book.root.find(q("sheetData"))
        nodes = list(sheet_data.findall(q("row")))
        header_node = next(
            (row for row in nodes if int(row.attrib.get("r", "0")) == HEADER_ROW),
            None,
        )
        data_rows = []
        for row in nodes:
            rn = int(row.attrib.get("r", "0"))
            if rn <= HEADER_ROW:
                continue
            cells = {col_number(cell.attrib["r"]): cell for cell in row.findall(q("c"))}
            d = canonical_date(book.value(cells.get(1)))
            if not d:
                if any(book.value(cell) not in (None, "") for cell in cells.values()):
                    raise RuntimeError(f"Data row {rn} has no valid date")
                continue
            data_rows.append((d, row))
        data_rows.sort(key=lambda pair: pair[0], reverse=True)
        for child in list(sheet_data):
            sheet_data.remove(child)
        sheet_data.append(header_node)
        for rn, (_, row) in enumerate(data_rows, start=HEADER_ROW + 1):
            row.set("r", str(rn))
            for cell in row.findall(q("c")):
                column = col_number(cell.attrib["r"])
                cell.set("r", f"{col_letters(column)}{rn}")
            sheet_data.append(row)
        last_row = HEADER_ROW + len(data_rows)
        dimension = book.root.find(q("dimension"))
        if dimension is not None:
            dimension.set("ref", f"A1:E{last_row}")
        auto_filter = book.root.find(q("autoFilter"))
        if auto_filter is not None:
            auto_filter.set("ref", f"A1:E{last_row}")
        book.save()

    check = Book(MASTER, SHEET)
    check_rows = check.rows()
    verified = {}
    for rn, cells in check_rows.items():
        if rn <= HEADER_ROW:
            continue
        d = canonical_date(check.value(cells.get(1)))
        if d:
            verified[d] = cells
    missing = sorted(set(records) - set(verified))
    if missing:
        raise RuntimeError(f"Manual dates missing after merge: {missing}")
    for d in inserted:
        cells = verified[d]
        item = records[d]
        actual = [check.value(cells.get(c)) for c in (2, 3, 4)]
        expected_values = [item["index"], item["equal_weight"], item["value"]]
        for got, want in zip(actual, expected_values):
            if not math.isclose(float(got), want, rel_tol=1e-12, abs_tol=1e-9):
                raise RuntimeError(f"Value verification failed for {d}")
        if item["volume"] is not None:
            got = check.value(cells.get(5))
            if not math.isclose(float(got), item["volume"], rel_tol=1e-12, abs_tol=1e-9):
                raise RuntimeError(f"Volume verification failed for {d}")
    print(f"Manual historical rows supplied: {len(records)}")
    print(f"New rows inserted: {len(inserted)}")
    print(f"Existing dates left untouched: {len(records) - len(inserted)}")
    print(f"Verified manual dates: {len(records)}")
    print("Missing volume values remain blank; workbook rows sorted newest first.")

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        raise
