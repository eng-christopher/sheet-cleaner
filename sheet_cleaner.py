"""
Sheet Cleaner - turns a messy CSV/Excel export into a clean, formatted Excel report.

Problem it solves:
    Exports from shops, CRMs and old systems come out messy: dates in several
    formats, prices written as text ("$1,200.50"), names in random case, extra
    spaces, empty rows and repeated orders. Someone usually fixes this by hand
    every week.

What it does:
    1. Reads a CSV or XLSX file.
    2. Trims spaces, fixes text case, removes empty rows.
    3. Converts dates (many formats) to real dates and prices to real numbers.
    4. Removes exact duplicate orders.
    5. Writes a formatted Excel file with three sheets:
         - Clean Data  (ready to filter, sum and chart)
         - Summary     (totals and what was fixed)
         - Issues      (rows that could not be fixed automatically, with the reason)

Usage:
    python sheet_cleaner.py sample_data/messy_orders.csv --out output/clean_orders.xlsx
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$")

DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%b %d %Y", "%d %b %Y", "%Y/%m/%d"]

# Expected columns -> how to clean each one.
COLUMN_TYPES = {
    "order_id": "id",
    "order_date": "date",
    "customer": "name",
    "email": "email",
    "product": "text",
    "quantity": "int",
    "unit_price": "money",
    "status": "status",
}

STATUS_MAP = {
    "paid": "Paid", "complete": "Paid", "completed": "Paid",
    "pending": "Pending", "awaiting payment": "Pending",
    "refunded": "Refunded", "refund": "Refunded",
    "cancelled": "Cancelled", "canceled": "Cancelled",
}


class Cleaner:
    def __init__(self) -> None:
        self.fixes: dict[str, int] = {}
        self.issues: list[tuple[int, str]] = []

    def fix(self, name: str) -> None:
        self.fixes[name] = self.fixes.get(name, 0) + 1

    # ----- value cleaners ------------------------------------------------- #
    def clean_date(self, value: str, row: int):
        value = value.strip()
        if not value:
            self.issues.append((row, "missing date"))
            return None
        for fmt in DATE_FORMATS:
            try:
                parsed = datetime.strptime(value, fmt).date()
                if fmt != "%Y-%m-%d":
                    self.fix("date converted to standard format")
                return parsed
            except ValueError:
                continue
        self.issues.append((row, f"unrecognized date '{value}'"))
        return None

    def clean_money(self, value: str, row: int):
        raw = value.strip()
        text = re.sub(r"[^\d.,\-]", "", raw)
        if not text:
            self.issues.append((row, "missing price"))
            return None
        if "," in text and "." in text:
            text = text.replace(",", "")          # 1,200.50
        elif "," in text and len(text.split(",")[-1]) == 2:
            text = text.replace(",", ".")         # 12,50 (European)
        else:
            text = text.replace(",", "")
        try:
            number = round(float(text), 2)
        except ValueError:
            self.issues.append((row, f"invalid price '{raw}'"))
            return None
        if raw != text:
            self.fix("price converted from text to number")
        return number

    def clean_int(self, value: str, row: int):
        try:
            return int(float(value.strip()))
        except ValueError:
            self.issues.append((row, f"invalid quantity '{value}'"))
            return None

    def clean_name(self, value: str) -> str:
        cleaned = re.sub(r"\s+", " ", value).strip()
        if cleaned.isupper() or cleaned.islower():
            self.fix("name case fixed")
            cleaned = cleaned.title()
        return cleaned

    def clean_status(self, value: str) -> str:
        key = value.strip().lower()
        result = STATUS_MAP.get(key, value.strip().title())
        if result != value:
            self.fix("status standardized")
        return result

    def clean_cell(self, kind: str, value: str, row: int):
        original = value
        value = "" if value is None else str(value)
        if value != value.strip() or "  " in value:
            self.fix("extra spaces removed")
        if kind == "date":
            return self.clean_date(value, row)
        if kind == "money":
            return self.clean_money(value, row)
        if kind == "int":
            return self.clean_int(value, row)
        if kind == "name":
            return self.clean_name(value)
        if kind == "email":
            cleaned = value.strip().lower()
            if cleaned and not EMAIL_RE.match(cleaned):
                self.issues.append((row, f"invalid email '{cleaned}'"))
            return cleaned
        if kind == "status":
            return self.clean_status(value)
        return re.sub(r"\s+", " ", value).strip() if original is not None else ""


# --------------------------------------------------------------------------- #
def read_rows(path: Path) -> list[dict[str, str]]:
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        ws = load_workbook(path, read_only=True, data_only=True).active
        rows = list(ws.iter_rows(values_only=True))
        header = [str(h).strip().lower().replace(" ", "_") for h in rows[0]]
        return [{h: ("" if v is None else str(v)) for h, v in zip(header, r)} for r in rows[1:]]
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        reader.fieldnames = [h.strip().lower().replace(" ", "_") for h in reader.fieldnames or []]
        return list(reader)


def clean(rows: list[dict[str, str]]) -> tuple[list[dict], Cleaner, dict[str, int]]:
    cleaner = Cleaner()
    clean_rows: list[dict] = []
    seen: set[tuple] = set()
    stats = {"received": len(rows), "empty_removed": 0, "duplicates_removed": 0}

    for index, row in enumerate(rows, start=2):  # row 1 is the header
        if not any((v or "").strip() for v in row.values()):
            stats["empty_removed"] += 1
            continue
        issues_before = len(cleaner.issues)
        cleaned = {col: cleaner.clean_cell(kind, row.get(col, ""), index) for col, kind in COLUMN_TYPES.items()}
        if len(cleaner.issues) > issues_before:
            continue  # row has a problem we can't fix safely -> goes to Issues sheet
        if cleaned["quantity"] is not None and cleaned["unit_price"] is not None:
            cleaned["total"] = round(cleaned["quantity"] * cleaned["unit_price"], 2)
        key = tuple(cleaned.values())
        if key in seen:
            stats["duplicates_removed"] += 1
            continue
        seen.add(key)
        clean_rows.append(cleaned)

    clean_rows.sort(key=lambda r: (r["order_date"] or date.min, r["order_id"]))
    return clean_rows, cleaner, stats


# --------------------------------------------------------------------------- #
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_header(ws, widths: list[int]) -> None:
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"


def write_excel(clean_rows, cleaner: Cleaner, stats, out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    ws = wb.active
    ws.title = "Clean Data"
    headers = ["Order ID", "Order Date", "Customer", "Email", "Product", "Quantity", "Unit Price", "Status", "Total"]
    ws.append(headers)
    for r in clean_rows:
        ws.append([r["order_id"], r["order_date"], r["customer"], r["email"], r["product"],
                   r["quantity"], r["unit_price"], r["status"], r.get("total")])
    for row in ws.iter_rows(min_row=2):
        row[1].number_format = "yyyy-mm-dd"
        row[6].number_format = row[8].number_format = '"$"#,##0.00'
    style_header(ws, [11, 12, 20, 28, 22, 10, 12, 11, 12])
    ws.auto_filter.ref = ws.dimensions

    ws2 = wb.create_sheet("Summary")
    ws2.append(["Item", "Value"])
    paid_total = sum(r.get("total") or 0 for r in clean_rows if r["status"] == "Paid")
    summary = [
        ("Rows received", stats["received"]),
        ("Clean rows", len(clean_rows)),
        ("Empty rows removed", stats["empty_removed"]),
        ("Duplicate orders removed", stats["duplicates_removed"]),
        ("Rows sent to Issues", len({r for r, _ in cleaner.issues})),
        ("Revenue from paid orders", paid_total),
        ("", ""),
        ("Automatic fixes", "Count"),
        *sorted(cleaner.fixes.items()),
    ]
    for item in summary:
        ws2.append(list(item))
    ws2["B7"].number_format = '"$"#,##0.00'
    ws2["A9"].font = ws2["B9"].font = Font(bold=True)
    style_header(ws2, [36, 16])

    ws3 = wb.create_sheet("Issues")
    ws3.append(["Original Row", "Problem"])
    for row_number, problem in cleaner.issues:
        ws3.append([row_number, problem])
    style_header(ws3, [14, 45])

    wb.save(out)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Clean a messy CSV/XLSX export into a formatted Excel report.")
    parser.add_argument("input", type=Path)
    parser.add_argument("--out", type=Path, default=Path("output/clean_orders.xlsx"))
    args = parser.parse_args(argv)

    if not args.input.exists():
        print(f"File not found: {args.input}", file=sys.stderr)
        return 1
    rows = read_rows(args.input)
    clean_rows, cleaner, stats = clean(rows)
    write_excel(clean_rows, cleaner, stats, args.out)

    print(f"Rows received: {stats['received']}")
    print(f"Clean rows: {len(clean_rows)}")
    print(f"Empty rows removed: {stats['empty_removed']}")
    print(f"Duplicates removed: {stats['duplicates_removed']}")
    print(f"Rows with issues: {len({r for r, _ in cleaner.issues})}")
    print("Fixes:", ", ".join(f"{k} ({v})" for k, v in sorted(cleaner.fixes.items())))
    print(f"Saved: {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
