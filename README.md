# Excel Report Builder

Turn a messy CSV export (orders, invoices, sales) into a clean, formatted Excel report in one command, with a log of every row it fixed or dropped.

```
$ python report.py sample_data/orders_messy.csv --out sample_output/sales_report.xlsx
246 rows in -> 240 clean orders, 9 issues logged -> sample_output/sales_report.xlsx
```

## What's in the report

| Sheet | Contents |
|---|---|
| **Summary** | Revenue, orders, average order, customers, best month, top customer, rows fixed or dropped |
| **Monthly** | Revenue by month × category with live `SUM` totals, data bars and a stacked column chart |
| **Top customers** | Top 10 by revenue with order count and share of total |
| **Clean data** | Every usable row, typed (real dates and numbers), filterable |
| **Data issues** | Each problem with its line number, the raw value and what was done |

Headers are styled, panes frozen, columns sized, and money formatted in EUR, USD or GBP.

## The mess it handles

The demo file ([`sample_data/orders_messy.csv`](sample_data/orders_messy.csv)) mixes everything real exports throw at you:

- **Dates** as `2026-03-04`, `04.03.2026`, `04/03/2026`, `4 Mar 2026`, `Mar 04, 2026` and Excel serials (`46085`); impossible ones like `31/02/2026` are logged and dropped. `--us-dates` switches 03/04 to month-first.
- **Amounts** as `€1.234,50`, `1,234.50 EUR`, `EUR 99.90`, `540,00`, `(45.00)` for negatives, or `n/a`.
- **Names** with stray spaces and random case: `  acme   builders ltd` → `Acme Builders Ltd`; company forms like SIA, AS and UAB keep their capitals.
- **Categories** with typos and aliases: `hw`, `HARDWARE`, ` software `, `consulting` → Hardware / Software / Services.
- **Duplicate orders** removed; blank customers set to "Unknown" and logged.
- **Column names** matched loosely: "Order No", "Order Date", "Client", "Total" all work without renaming.

## Quick start

```bash
python -m venv .venv
.venv/Scripts/activate        # Windows; `source .venv/bin/activate` on macOS/Linux
pip install -r requirements.txt
python make_sample.py          # regenerate the messy demo file
python report.py sample_data/orders_messy.csv --out output/report.xlsx --currency EUR
pytest -q
```

## Adapting it

The cleaning rules live in `clean.py` as small, tested functions, so adding a new date format, currency or category alias is a one-line change. Typical extensions: several input files merged into one report, a scheduled run that emails the workbook, or a Google Sheets output instead of Excel.

## License

MIT
