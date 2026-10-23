# Excel / CSV Automation — Merge & Clean Reports

Merges a pile of messy spreadsheet exports (CSV + Excel, each with slightly
different headers) into **one clean, formatted Excel report** — the kind of
job every small business does by hand at the end of the month.

## What it does

1. **Reads** any mix of `.csv` / `.xlsx` / `.xls` files — pass files or a
   whole folder.
2. **Normalises headers** to `snake_case` and maps aliases, so `Order ID`,
   `order_id` and `Order Id` become one column, and `Qty` becomes `quantity`.
3. **Cleans the data**: trims stray whitespace, treats empty cells as
   missing, drops fully-empty rows and exact duplicate records (even when
   the duplicate comes from a different file), and converts numeric columns
   to real numbers.
4. **Writes a formatted Excel report** with two sheets:
   - `Clean Data` — all merged rows, tagged with their source file
   - `Summary` — per-product totals (orders, quantity, revenue) + TOTAL row

   Headers are styled, columns auto-sized, panes frozen, money columns get a
   `#,##0.00` number format.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
# Merge everything in the demo folder
python merge_reports.py sample_data

# Merge specific files into a specific report
python merge_reports.py jan.csv feb.xlsx mar.csv --output output/monthly_report.xlsx
```

Sample data is included so it runs immediately: three monthly sales exports
(January CSV, February CSV, March Excel) that deliberately contain the
classic problems — mismatched headers, one cross-file duplicate, stray
spaces, and a blank cell.

## Sample output

```
Files merged : sales_feb.csv, sales_jan.csv, sales_mar.xlsx
Rows in      : 12
Duplicates   : 1 removed
Rows out     : 11
Report saved : output/merged_report.xlsx
```

Summary sheet (excerpt):

| product             | orders | total_quantity | total_revenue |
|---------------------|-------:|---------------:|--------------:|
| Webcam HD           |      2 |              3 |        177.00 |
| Wireless Mouse      |      3 |              6 |        151.47 |
| Monitor Stand       |      2 |              3 |        135.00 |
| Mechanical Keyboard |      1 |              1 |         89.99 |
| USB-C Cable         |      2 |              7 |         69.93 |
| Laptop Sleeve       |      1 |              0 |          0.00 |
| **TOTAL**           | **11** |              |      **623.39** |

## Tests

```bash
python -m pytest tests/ -v
```

## Customising

- **Header aliases** live in `COLUMN_ALIASES` in `merge_reports.py`.
- **Numeric columns** live in `NUMERIC_COLUMNS`.
- The summary groups by `product` / `quantity` / `unit_price`; for non-sales
  data the Clean Data sheet is still produced and the summary is skipped.

## License

MIT — free to use as a starting point for your own reporting automation.
