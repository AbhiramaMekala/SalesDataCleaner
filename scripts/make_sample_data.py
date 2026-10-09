"""Generate realistic, deliberately messy monthly sales CSVs for demo and testing.

Each file has the kinds of problems real exports have:
  * the same order exported twice (exact duplicates and "near" duplicates
    that differ only in spacing or capitalisation)
  * blank quantity, unit price or region cells
  * dates in several formats (2026-07-03, 07/04/2026, 5-Jul-2026, 2026/07/06)
  * stray spaces and inconsistent capitalisation in text columns
  * prices written with a "$" sign

Usage:  python scripts/make_sample_data.py            # writes to data/raw/
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

PRODUCTS = {
    # product: (category, typical unit price)
    "Wireless Mouse": ("Accessories", 24.99),
    "Mechanical Keyboard": ("Accessories", 89.00),
    "USB-C Hub": ("Accessories", 39.50),
    "27in Monitor": ("Displays", 279.00),
    "Laptop Stand": ("Accessories", 45.00),
    "Noise-Cancelling Headphones": ("Audio", 199.99),
    "Bluetooth Speaker": ("Audio", 59.95),
    "Webcam 1080p": ("Video", 69.00),
}
REGIONS = ["Northeast", "Southeast", "Midwest", "West"]
DATE_STYLES = ["%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y", "%Y/%m/%d"]
HEADER = ["order_id", "order_date", "region", "product", "category", "quantity", "unit_price"]


def messy_text(value: str, rng: random.Random) -> str:
    roll = rng.random()
    if roll < 0.08:
        return f"  {value} "
    if roll < 0.14:
        return value.upper()
    if roll < 0.20:
        return value.lower()
    return value


def make_month(year: int, month: int, n_orders: int, first_id: int,
               rng: random.Random) -> list[list[str]]:
    start = date(year, month, 1)
    days = ((start.replace(day=28) + timedelta(days=4)).replace(day=1) - start).days
    rows = []
    for i in range(n_orders):
        product = rng.choice(list(PRODUCTS))
        category, price = PRODUCTS[product]
        day = start + timedelta(days=rng.randrange(days))
        qty = rng.choices([1, 2, 3, 4, 5, 10], weights=[45, 25, 12, 8, 6, 4])[0]
        unit = round(price * rng.uniform(0.95, 1.05), 2)

        row = [
            f"ORD-{first_id + i:05d}",
            day.strftime(rng.choice(DATE_STYLES)),
            messy_text(rng.choice(REGIONS), rng),
            messy_text(product, rng),
            category,
            str(qty),
            f"${unit}" if rng.random() < 0.15 else str(unit),
        ]
        # Missing values (~3% each)
        if rng.random() < 0.03:
            row[5] = ""
        if rng.random() < 0.03:
            row[6] = ""
        if rng.random() < 0.03:
            row[2] = ""
        rows.append(row)

    # Duplicates: ~4% exact copies, ~2% near-copies (spacing/case only)
    for row in rng.sample(rows, k=max(1, n_orders // 25)):
        rows.append(list(row))
    for row in rng.sample(rows, k=max(1, n_orders // 50)):
        near = list(row)
        near[3] = f" {near[3].strip().lower()} "
        rows.append(near)
    rng.shuffle(rows)
    return rows


def main(out_dir: str = "data/raw", seed: int = 2026) -> None:
    rng = random.Random(seed)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    next_id = 10001
    for month, n in [(7, 420), (8, 465), (9, 510)]:
        rows = make_month(2026, month, n, next_id, rng)
        next_id += n
        path = out / f"sales_2026_{month:02d}.csv"
        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(HEADER)
            writer.writerows(rows)
        print(f"Wrote {len(rows):>4} rows -> {path}")


if __name__ == "__main__":
    main()
