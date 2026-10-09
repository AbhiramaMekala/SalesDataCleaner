"""Command-line entry point.

    python -m sales_cleaner                       # data/raw -> output/
    python -m sales_cleaner -i exports/ -o reports/ --pattern "sales_*.csv"
"""

import argparse
import sys
import time
from pathlib import Path

from .cleaner import clean_sales, load_monthly_files, save_clean_csv
from .report import build_report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="sales_cleaner",
        description="Clean monthly sales CSVs and build a formatted Excel report.")
    parser.add_argument("-i", "--input", default="data/raw", help="folder with monthly CSV files")
    parser.add_argument("-o", "--output", default="output", help="folder for the results")
    parser.add_argument("--pattern", default="*.csv", help="which files to read (glob)")
    args = parser.parse_args(argv)

    started = time.perf_counter()
    try:
        raw, files = load_monthly_files(args.input, args.pattern)
    except (FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    clean, log = clean_sales(raw)
    log.files_read = files

    out = Path(args.output)
    csv_path = save_clean_csv(clean, out / "clean_sales.csv")
    xlsx_path = build_report(clean, log, out / "sales_report.xlsx")
    elapsed = time.perf_counter() - started

    print(f"Read {log.rows_read:,} rows from {len(files)} file(s)")
    print(f"  removed {log.duplicates_removed:,} duplicates, "
          f"filled {log.values_filled:,} missing values, "
          f"fixed {log.dates_reformatted:,} dates, "
          f"dropped {log.rows_dropped_bad_date + log.rows_dropped_missing_key:,} unusable rows")
    print(f"Wrote {log.rows_written:,} clean rows -> {csv_path}")
    print(f"Report -> {xlsx_path}")
    print(f"Done in {elapsed:.2f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
