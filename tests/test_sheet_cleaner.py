import sys
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import sheet_cleaner as sc  # noqa: E402

SAMPLE = Path(__file__).resolve().parent.parent / "sample_data" / "messy_orders.csv"


def test_date_formats():
    c = sc.Cleaner()
    assert c.clean_date("10/02/2026", 2) == date(2026, 10, 2)
    assert c.clean_date("Oct 03 2026", 2) == date(2026, 10, 3)
    assert c.clean_date("32/13/2026", 2) is None
    assert c.issues


def test_money_formats():
    c = sc.Cleaner()
    assert c.clean_money("$1,200.50", 2) == 1200.50
    assert c.clean_money("19,90", 2) == 19.90
    assert c.clean_money("45", 2) == 45.0


def test_names_and_status():
    c = sc.Cleaner()
    assert c.clean_name("  MARIA   GARCIA ") == "Maria Garcia"
    assert c.clean_status("Completed") == "Paid"
    assert c.clean_status("canceled") == "Cancelled"


def test_sample_end_to_end(tmp_path):
    out = tmp_path / "clean.xlsx"
    assert sc.main([str(SAMPLE), "--out", str(out)]) == 0
    wb = load_workbook(out)
    assert wb.sheetnames == ["Clean Data", "Summary", "Issues"]
    data = list(wb["Clean Data"].iter_rows(min_row=2, values_only=True))
    assert len(data) == 7           # 13 rows - 2 empty - 1 duplicate - 3 with issues
    assert len(list(wb["Issues"].iter_rows(min_row=2))) == 3
    ids = [row[0] for row in data]
    assert len(ids) == len(set(ids))  # no duplicate orders
