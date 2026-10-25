"""Messy CSV export → cleaned data + formatted Excel report.

Usage:
    python report.py sample_data/orders_messy.csv --out sample_output/sales_report.xlsx
    python report.py orders.csv --us-dates --currency USD

Expected columns (names are matched loosely, e.g. "Order Date" or "date"):
order id, date, customer, category, amount. Extra columns are kept.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from clean import clean_category, clean_name, parse_amount, parse_date

COLUMN_HINTS = {
    "order_id": ["order id", "order_id", "order no", "order number", "invoice", "id"],
    "date": ["order date", "date", "created", "invoice date"],
    "customer": ["customer", "client", "company", "buyer", "name"],
    "category": ["category", "product category", "type", "segment"],
    "amount": ["amount", "total", "revenue", "value", "price", "sum"],
}
CATEGORY_ALIASES = {"hw": "Hardware", "hardware": "Hardware", "sw": "Software", "software": "Software",
                    "services": "Services", "service": "Services", "consulting": "Services"}

HEADER_FILL = PatternFill("solid", start_color="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FONT = Font(bold=True)
THIN = Side(style="thin", color="C9D1DB")


def match_columns(cols: list[str]) -> dict[str, str]:
    norm = {c: re.sub(r"[^a-z0-9]+", " ", c.lower()).strip() for c in cols}
    found: dict[str, str] = {}
    for key, hints in COLUMN_HINTS.items():
        for hint in hints:
            hit = next((c for c, n in norm.items() if n == hint and c not in found.values()), None)
            if hit:
                found[key] = hit
                break
    missing = [k for k in ("date", "customer", "amount") if k not in found]
    if missing:
        raise SystemExit(f"Could not find columns for: {', '.join(missing)}. Columns in file: {cols}")
    return found


def clean_frame(raw: pd.DataFrame, day_first: bool) -> tuple[pd.DataFrame, pd.DataFrame]:
    cols = match_columns(list(raw.columns))
    issues = []
    rows = []
    for i, r in raw.iterrows():
        line = i + 2  # header is line 1
        d = parse_date(r[cols["date"]], day_first)
        amt = parse_amount(r[cols["amount"]])
        cust = clean_name(r[cols["customer"]])
        if d is None:
            issues.append({"line": line, "problem": "Unreadable date", "value": r[cols["date"]], "action": "Row dropped"})
            continue
        if amt is None:
            issues.append({"line": line, "problem": "Missing amount", "value": r[cols["amount"]], "action": "Row dropped"})
            continue
        if cust is None:
            issues.append({"line": line, "problem": "Missing customer", "value": "", "action": "Set to 'Unknown'"})
            cust = "Unknown"
        rows.append({
            "order_id": str(r[cols["order_id"]]).strip() if "order_id" in cols else str(line),
            "date": pd.Timestamp(d),
            "customer": cust,
            "category": clean_category(r[cols["category"]], CATEGORY_ALIASES) if "category" in cols else "Uncategorised",
            "amount": round(amt, 2),
        })
    df = pd.DataFrame(rows)
    before = len(df)
    dupes = df[df.duplicated(subset=["order_id", "date", "customer", "amount"], keep="first")]
    for _, r in dupes.iterrows():
        issues.append({"line": "", "problem": "Duplicate order", "value": r["order_id"], "action": "Duplicate removed"})
    df = df.drop_duplicates(subset=["order_id", "date", "customer", "amount"], keep="first").sort_values("date")
    assert len(df) == before - len(dupes)
    df["month"] = df["date"].dt.to_period("M").astype(str)
    return df.reset_index(drop=True), pd.DataFrame(issues, columns=["line", "problem", "value", "action"])


def build_report(df: pd.DataFrame, issues: pd.DataFrame, raw_rows: int, out: Path, currency: str) -> None:
    monthly = df.pivot_table(index="month", columns="category", values="amount", aggfunc="sum", fill_value=0)
    monthly["Total"] = monthly.sum(axis=1)
    monthly = monthly.reset_index().rename(columns={"month": "Month"})
    customers = (df.groupby("customer").agg(orders=("order_id", "count"), revenue=("amount", "sum"))
                 .sort_values("revenue", ascending=False).reset_index())
    customers["share"] = customers["revenue"] / customers["revenue"].sum()
    top = customers.head(10).rename(columns={"customer": "Customer", "orders": "Orders", "revenue": "Revenue", "share": "Share"})
    summary = pd.DataFrame([
        ("Revenue", df["amount"].sum()),
        ("Orders", len(df)),
        ("Average order", df["amount"].mean()),
        ("Customers", df["customer"].nunique()),
        ("Best month", monthly.loc[monthly["Total"].idxmax(), "Month"]),
        ("Top customer", customers.loc[customers["customer"] != "Unknown", "customer"].iloc[0]),
        ("Rows in file", raw_rows),
        ("Rows fixed or dropped", len(issues)),
    ], columns=["Metric", "Value"])
    clean_out = df.drop(columns="month").rename(columns=str.title).rename(columns={"Order_Id": "Order ID"})

    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        monthly.to_excel(xw, sheet_name="Monthly", index=False)
        top.to_excel(xw, sheet_name="Top customers", index=False)
        clean_out.to_excel(xw, sheet_name="Clean data", index=False)
        issues.to_excel(xw, sheet_name="Data issues", index=False)
    style_workbook(out, currency, n_months=len(monthly), n_cats=len(monthly.columns) - 2)


def style_workbook(path: Path, currency: str, n_months: int, n_cats: int) -> None:
    money = {"EUR": '#,##0.00 "€"', "USD": '"$"#,##0.00', "GBP": '"£"#,##0.00'}.get(currency, "#,##0.00")
    wb = load_workbook(path)
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.fill, cell.font = HEADER_FILL, HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = Border(bottom=THIN)
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(10, width + 2), 48)

    s = wb["Summary"]
    for r in range(2, s.max_row + 1):
        label = s.cell(r, 1).value
        if label in ("Revenue", "Average order"):
            s.cell(r, 2).number_format = money
        elif label in ("Orders", "Customers", "Rows in file", "Rows fixed or dropped"):
            s.cell(r, 2).number_format = "#,##0"
        s.cell(r, 2).alignment = Alignment(horizontal="right")
    s.auto_filter.ref = None

    m = wb["Monthly"]
    for row in m.iter_rows(min_row=2, min_col=2):
        for c in row:
            c.number_format = money
    total_row = m.max_row + 1
    m.cell(total_row, 1, "Total").font = TOTAL_FONT
    for col in range(2, m.max_column + 1):
        letter = get_column_letter(col)
        c = m.cell(total_row, col, f"=SUM({letter}2:{letter}{total_row - 1})")
        c.number_format, c.font = money, TOTAL_FONT
    total_col = get_column_letter(m.max_column)
    m.conditional_formatting.add(f"{total_col}2:{total_col}{total_row - 1}",
                                 DataBarRule(start_type="min", end_type="max", color="5B8DEF"))
    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "col", "stacked", 100
    chart.title, chart.y_axis.title = "Revenue by month", currency
    chart.add_data(Reference(m, min_col=2, max_col=1 + n_cats, min_row=1, max_row=1 + n_months), titles_from_data=True)
    chart.set_categories(Reference(m, min_col=1, min_row=2, max_row=1 + n_months))
    chart.height, chart.width = 9, 18
    m.add_chart(chart, f"A{total_row + 3}")

    t = wb["Top customers"]
    for r in range(2, t.max_row + 1):
        t.cell(r, 3).number_format = money
        t.cell(r, 4).number_format = "0.0%"
    t.conditional_formatting.add(f"C2:C{t.max_row}", DataBarRule(start_type="num", start_value=0, end_type="max", color="5B8DEF"))

    c = wb["Clean data"]
    for r in range(2, c.max_row + 1):
        c.cell(r, 2).number_format = "yyyy-mm-dd"
        c.cell(r, 5).number_format = money
    wb.save(path)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Clean a sales CSV and build a formatted Excel report.")
    ap.add_argument("csv", type=Path)
    ap.add_argument("--out", type=Path, default=Path("output/sales_report.xlsx"))
    ap.add_argument("--us-dates", action="store_true", help="read 03/04/2026 as March 4 (default: 3 April)")
    ap.add_argument("--currency", default="EUR", choices=["EUR", "USD", "GBP", "OTHER"])
    args = ap.parse_args(argv)

    raw = pd.read_csv(args.csv, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    df, issues = clean_frame(raw, day_first=not args.us_dates)
    if df.empty:
        print("No usable rows after cleaning. See the column names and date format.", file=sys.stderr)
        return 1
    build_report(df, issues, len(raw), args.out, args.currency)
    print(f"{len(raw)} rows in -> {len(df)} clean orders, {len(issues)} issues logged -> {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
