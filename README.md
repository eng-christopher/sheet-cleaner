# Sheet Cleaner — from messy export to clean Excel report in seconds

Exports from online shops, CRMs and older systems are rarely ready to use: dates in five formats, prices stored as text, names in random case, extra spaces, empty rows and repeated orders. Fixing this by hand every week wastes hours and introduces mistakes.

This Python script does the cleanup automatically and produces a formatted Excel file.

## What it does
- **Dates**: recognizes 7 common formats (`10/02/2026`, `Oct 03 2026`, `2026/10/04`...) and converts them to real Excel dates
- **Prices**: turns text like `$1,200.50` or `19,90` into real numbers you can sum
- **Text**: trims extra spaces and fixes ALL CAPS / lowercase names
- **Status**: standardizes values (`completed` → `Paid`, `canceled` → `Cancelled`)
- **Removes** empty rows and duplicate orders
- **Calculates** the line total for each order
- **Never hides a problem**: rows that can't be fixed safely (impossible date, invalid email, quantity written as text) go to an *Issues* sheet with the reason

## Output: one Excel file, three sheets
| Sheet | Contents |
|---|---|
| Clean Data | Formatted table with filters, real dates and currency |
| Summary | Rows received, removed, fixed, and revenue from paid orders |
| Issues | Original row number and the exact problem |

## Result on the sample data
13 rows received → 7 clean orders, 2 empty rows and 1 duplicate removed, 3 rows flagged for review, 34 automatic fixes.

## How to run
```bash
pip install openpyxl
python sheet_cleaner.py sample_data/messy_orders.csv --out output/clean_orders.xlsx
```
Accepts `.csv` and `.xlsx`.

## Tests
```bash
python -m pytest tests
```
