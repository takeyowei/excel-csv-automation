"""Tests for merge_reports. Run with: python -m pytest tests/ -v"""

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from merge_reports import (  # noqa: E402
    build_summary,
    clean_data,
    load_file,
    merge_files,
    normalize_column,
)

PROJECT_DIR = Path(__file__).resolve().parent.parent
SAMPLE = PROJECT_DIR / "sample_data"


def test_normalize_column_variants():
    assert normalize_column("Order ID ") == "order_id"
    assert normalize_column("Unit-Price") == "unit_price"
    assert normalize_column("Qty") == "quantity"          # alias map
    assert normalize_column("order_date") == "order_date"


def test_load_file_normalises_headers():
    df = load_file(SAMPLE / "sales_mar.xlsx")
    assert list(df.columns) == ["order_id", "product", "quantity",
                                "unit_price", "order_date", "source_file"]


def test_merge_and_clean_sample_data(tmp_path):
    out = tmp_path / "report.xlsx"
    result = merge_files([str(SAMPLE)], out)
    stats = result["stats"]
    assert stats["rows_in"] == 12          # 5 (Jan) + 4 (Feb) + 3 (Mar)
    assert stats["duplicates_removed"] == 1  # order 1003 appears twice
    assert stats["rows_out"] == 11
    assert out.exists()

    clean = pd.read_excel(out, sheet_name="Clean Data")
    # stray spaces around " Wireless Mouse " were trimmed
    assert "Wireless Mouse" in set(clean["product"])
    assert not any(isinstance(p, str) and p != p.strip() for p in clean["product"])
    # the March row with a blank quantity survived, quantity is empty
    sleeve = clean[clean["product"] == "Laptop Sleeve"].iloc[0]
    assert pd.isna(sleeve["quantity"])


def test_summary_totals(tmp_path):
    out = tmp_path / "report.xlsx"
    merge_files([str(SAMPLE)], out)
    summary = pd.read_excel(out, sheet_name="Summary")
    total = summary[summary["product"] == "TOTAL"].iloc[0]
    assert total["orders"] == 11
    # Mouse 151.47 + Keyboard 89.99 + Cable 69.93 + Stand 135 + Webcam 177 + Sleeve 0
    assert abs(total["total_revenue"] - 623.39) < 0.01
    mouse = summary[summary["product"] == "Wireless Mouse"].iloc[0]
    assert mouse["total_quantity"] == 6


def test_clean_data_drops_blank_rows():
    raw = pd.DataFrame({"product": ["A", "", "A"], "quantity": ["1", "", "1"],
                        "source_file": ["x.csv"] * 3})
    clean, stats = clean_data(raw)
    assert stats["rows_in"] == 3
    assert stats["rows_out"] == 1  # blank row dropped, exact duplicate removed
