import sys
from datetime import date
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from clean import clean_category, clean_name, parse_amount, parse_date  # noqa: E402
from report import clean_frame  # noqa: E402


@pytest.mark.parametrize("raw, want", [
    ("2026-03-04", date(2026, 3, 4)),
    ("04.03.2026", date(2026, 3, 4)),
    ("04/03/2026", date(2026, 3, 4)),
    ("4 Mar 2026", date(2026, 3, 4)),
    ("Mar 04, 2026", date(2026, 3, 4)),
    ("46085", date(2026, 3, 4)),        # Excel serial
    ("31/02/2026", None),
    ("sometime in May", None),
    ("", None),
])
def test_parse_date(raw, want):
    assert parse_date(raw) == want


def test_us_dates():
    assert parse_date("04/03/2026", day_first=False) == date(2026, 4, 3)


@pytest.mark.parametrize("raw, want", [
    ("€1.234,50", 1234.5), ("1,234.50 EUR", 1234.5), ("EUR 99.90", 99.9), ("540,00", 540.0),
    ("$1,200", 1200.0), ("(45.00)", -45.0), ("n/a", None), ("", None), (12, 12.0),
])
def test_parse_amount(raw, want):
    assert parse_amount(raw) == want


def test_clean_name_and_category():
    assert clean_name("  acme   builders LTD ") == "Acme Builders Ltd"
    assert clean_name("SIA baltic build") == "SIA Baltic Build"
    assert clean_name("   ") is None
    assert clean_category(" hw ", {"hw": "Hardware"}) == "Hardware"
    assert clean_category("") == "Uncategorised"


def test_clean_frame_drops_and_dedupes():
    raw = pd.DataFrame({
        "Order No": ["A1", "A1", "A2", "A3", "A4"],
        "Order Date": ["2026-01-05", "2026-01-05", "31/02/2026", "2026-01-07", "2026-01-08"],
        "Client": ["x co", "x co", "y", "", "z"],
        "Product Category": ["hw", "hw", "sw", "Services", ""],
        "Total": ["10,00", "10,00", "5", "7", "n/a"],
    })
    df, issues = clean_frame(raw, day_first=True)
    assert list(df["order_id"]) == ["A1", "A3"]
    assert df.loc[df.order_id == "A3", "customer"].item() == "Unknown"
    assert sorted(issues["problem"]) == ["Duplicate order", "Missing amount", "Missing customer", "Unreadable date"]
