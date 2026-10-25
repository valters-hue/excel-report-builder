"""Generate a deliberately messy orders export for the demo (fixed seed)."""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(7)
CUSTOMERS = ["SIA Baltic Build", "acme builders ltd", "  Nordic   Foods AS", "Riga Print House", "BMW Latvia",
             "green leaf cafe", "UAB Vilnius Tech", "Tallinn Logistics OÜ", "Daugava Dental"]
CATEGORIES = ["Hardware", "hw", "Software", " software ", "Services", "consulting", "", "HARDWARE"]


def fmt_date(d: date) -> str:
    return random.choice([
        d.isoformat(), d.strftime("%d.%m.%Y"), d.strftime("%d/%m/%Y"), d.strftime("%d %b %Y"),
        d.strftime("%b %d, %Y"), str((d - date(1899, 12, 30)).days),
    ])


def fmt_amount(x: float) -> str:
    return random.choice([
        f"{x:.2f}", f"€{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
        f"{x:,.2f} EUR", f"EUR {x:.2f}", f"{x:.2f}".replace(".", ","),
    ])


def main() -> None:
    out = Path("sample_data/orders_messy.csv")
    out.parent.mkdir(exist_ok=True)
    rows = []
    start = date(2026, 1, 3)
    for i in range(1, 241):
        d = start + timedelta(days=random.randint(0, 290))
        amount = round(random.lognormvariate(5.6, 0.7), 2)
        rows.append([f"INV-{1000 + i}", fmt_date(d), random.choice(CUSTOMERS), random.choice(CATEGORIES), fmt_amount(amount)])
    rows += [rows[5][:], rows[17][:], rows[90][:]]                      # duplicates
    for r in (rows[30], rows[60], rows[120]):                             # blank customers
        r[2] = ""
    rows += [["INV-1301", "31/02/2026", "Riga Print House", "Software", "120"],  # impossible date
             ["INV-1302", "2026-05-14", "Daugava Dental", "Services", "n/a"],    # missing amount
             ["INV-1303", "sometime in May", "BMW Latvia", "Hardware", "540,00"]]
    random.shuffle(rows)
    with out.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Order No", "Order Date", "Client", "Product Category", "Total"])
        w.writerows(rows)
    print(f"Wrote {len(rows)} rows to {out}")


if __name__ == "__main__":
    main()
