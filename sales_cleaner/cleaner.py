"""Load monthly sales CSVs and clean them into one tidy table.

Cleaning steps, in order:
  1. Normalise text  - trim spaces, collapse inner spaces, unify capitalisation
  2. Fix numbers     - strip "$" and "," from prices, convert to numbers
  3. Fix dates       - parse several known formats into real dates
  4. Remove duplicates - exact repeats, then repeated order ids
  5. Fill missing values - price/quantity from the same product's median,
                           region -> "Unknown", category from the product
  6. Add revenue and month columns

Every step records what it changed in a CleaningLog, which ends up in the
Excel report so the reader can see exactly what was done to the data.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

REQUIRED_COLUMNS = ["order_id", "order_date", "region", "product",
                    "category", "quantity", "unit_price"]
TEXT_COLUMNS = ["order_id", "region", "product", "category"]

# Order matters: a value is parsed with the first format that fits.
# Slash dates with the year last are read US-style (month/day/year).
DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y", "%Y/%m/%d", "%d %B %Y", "%b %d, %Y"]

OUTPUT_COLUMNS = ["order_id", "order_date", "month", "region", "product", "category",
                  "quantity", "unit_price", "revenue", "source_file"]


@dataclass
class CleaningLog:
    files_read: list[str] = field(default_factory=list)
    rows_read: int = 0
    text_values_tidied: int = 0
    dates_reformatted: int = 0
    rows_dropped_bad_date: int = 0
    rows_dropped_missing_key: int = 0
    exact_duplicates_removed: int = 0
    repeated_order_ids_removed: int = 0
    filled_unit_price: int = 0
    filled_quantity: int = 0
    filled_region: int = 0
    filled_category: int = 0
    rows_written: int = 0

    @property
    def duplicates_removed(self) -> int:
        return self.exact_duplicates_removed + self.repeated_order_ids_removed

    @property
    def values_filled(self) -> int:
        return (self.filled_unit_price + self.filled_quantity
                + self.filled_region + self.filled_category)

    def as_rows(self) -> list[tuple[str, int | str]]:
        """Human-readable (label, value) pairs for the report."""
        return [
            ("Files read", len(self.files_read)),
            ("Rows read", self.rows_read),
            ("Text values tidied (spaces / capitalisation)", self.text_values_tidied),
            ("Dates converted to YYYY-MM-DD", self.dates_reformatted),
            ("Exact duplicate rows removed", self.exact_duplicates_removed),
            ("Repeated order IDs removed", self.repeated_order_ids_removed),
            ("Missing unit prices filled (product median)", self.filled_unit_price),
            ("Missing quantities filled (product median)", self.filled_quantity),
            ("Missing regions filled ('Unknown')", self.filled_region),
            ("Missing categories filled (from product)", self.filled_category),
            ("Rows dropped: unreadable date", self.rows_dropped_bad_date),
            ("Rows dropped: missing order ID or product", self.rows_dropped_missing_key),
            ("Rows written", self.rows_written),
        ]


# ---------------------------------------------------------------- loading

def load_monthly_files(input_dir: str | Path, pattern: str = "*.csv") -> tuple[pd.DataFrame, list[str]]:
    """Read every CSV in a folder into one DataFrame, remembering each row's file."""
    paths = sorted(Path(input_dir).glob(pattern))
    if not paths:
        raise FileNotFoundError(f"No files matching {pattern!r} in {input_dir}")

    frames = []
    for path in paths:
        # Read everything as text first; types are fixed explicitly during cleaning.
        df = pd.read_csv(path, dtype=str, keep_default_na=False, na_values=[""])
        df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise ValueError(f"{path.name} is missing columns: {', '.join(missing)}")
        df = df[REQUIRED_COLUMNS].copy()
        df["source_file"] = path.name
        frames.append(df)
    return pd.concat(frames, ignore_index=True), [p.name for p in paths]


# ---------------------------------------------------------------- steps

def _tidy_text(df: pd.DataFrame, log: CleaningLog) -> pd.DataFrame:
    for col in TEXT_COLUMNS:
        original = df[col]
        tidy = original.str.strip().str.replace(r"\s+", " ", regex=True)
        tidy = tidy.where(tidy != "", None)

        if col != "order_id":
            # Use the most common spelling of each value as the canonical one,
            # so "usb-c hub" and "USB-C HUB" both become "USB-C Hub".
            key = tidy.str.casefold()
            canonical = (tidy.groupby(key).agg(lambda s: s.value_counts().idxmax()))
            tidy = key.map(canonical).where(tidy.notna(), None)
        else:
            tidy = tidy.str.upper()

        changed = (original.fillna("") != tidy.fillna("")) & original.notna()
        log.text_values_tidied += int(changed.sum())
        df[col] = tidy
    return df


def _to_number(series: pd.Series) -> pd.Series:
    cleaned = series.str.replace(r"[$,\s]", "", regex=True)
    return pd.to_numeric(cleaned, errors="coerce")


def parse_dates(series: pd.Series) -> pd.Series:
    """Parse mixed-format date strings; anything unreadable becomes NaT."""
    s = series.str.strip()
    result = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    for fmt in DATE_FORMATS:
        todo = result.isna() & s.notna()
        if not todo.any():
            break
        result.loc[todo] = pd.to_datetime(s[todo], format=fmt, errors="coerce")
    return result


def _fix_dates(df: pd.DataFrame, log: CleaningLog) -> pd.DataFrame:
    raw = df["order_date"]
    parsed = parse_dates(raw)
    already_iso = raw.str.strip().str.fullmatch(r"\d{4}-\d{2}-\d{2}").fillna(False)
    log.dates_reformatted = int((parsed.notna() & ~already_iso).sum())

    bad = parsed.isna()
    log.rows_dropped_bad_date = int(bad.sum())
    df = df.loc[~bad].copy()
    df["order_date"] = parsed[~bad].dt.normalize()
    return df


def _remove_duplicates(df: pd.DataFrame, log: CleaningLog) -> pd.DataFrame:
    # Compare on the cleaned business columns (not the file name), so the same
    # order exported in two files, or with different spacing, still counts.
    before = len(df)
    df = df.drop_duplicates(subset=REQUIRED_COLUMNS, keep="first")
    log.exact_duplicates_removed = before - len(df)

    before = len(df)
    df = df.drop_duplicates(subset=["order_id"], keep="first")
    log.repeated_order_ids_removed = before - len(df)
    return df


def _fill_missing(df: pd.DataFrame, log: CleaningLog) -> pd.DataFrame:
    # Price: median price of the same product; fall back to the overall median.
    missing = df["unit_price"].isna()
    log.filled_unit_price = int(missing.sum())
    product_price = df.groupby("product")["unit_price"].transform("median")
    df["unit_price"] = (df["unit_price"].fillna(product_price)
                        .fillna(df["unit_price"].median()).round(2))

    # Quantity: median quantity of the same product, rounded to a whole unit.
    missing = df["quantity"].isna() | (df["quantity"] <= 0)
    log.filled_quantity = int(missing.sum())
    df.loc[df["quantity"] <= 0, "quantity"] = pd.NA
    product_qty = df.groupby("product")["quantity"].transform("median")
    df["quantity"] = df["quantity"].fillna(product_qty).fillna(1).round().astype(int)

    # Region: no safe way to guess, so label it clearly.
    log.filled_region = int(df["region"].isna().sum())
    df["region"] = df["region"].fillna("Unknown")

    # Category: every product belongs to one category, so look it up.
    missing = df["category"].isna()
    log.filled_category = int(missing.sum())
    if missing.any():
        lookup = (df.dropna(subset=["category"]).groupby("product")["category"]
                  .agg(lambda s: s.value_counts().idxmax()))
        df.loc[missing, "category"] = df.loc[missing, "product"].map(lookup)
        df["category"] = df["category"].fillna("Uncategorized")
    return df


# ---------------------------------------------------------------- pipeline

def clean_sales(df: pd.DataFrame, log: CleaningLog | None = None) -> tuple[pd.DataFrame, CleaningLog]:
    log = log or CleaningLog()
    log.rows_read = len(df)
    df = df.copy()
    if "source_file" not in df.columns:
        df["source_file"] = ""

    df = _tidy_text(df, log)
    df["unit_price"] = _to_number(df["unit_price"])
    df["quantity"] = _to_number(df["quantity"])
    df = _fix_dates(df, log)

    no_key = df["order_id"].isna() | df["product"].isna()
    log.rows_dropped_missing_key = int(no_key.sum())
    df = df.loc[~no_key].copy()

    df = _remove_duplicates(df, log)
    df = _fill_missing(df, log)

    df["revenue"] = (df["quantity"] * df["unit_price"]).round(2)
    df["month"] = df["order_date"].dt.strftime("%Y-%m")
    df = df.sort_values(["order_date", "order_id"]).reset_index(drop=True)[OUTPUT_COLUMNS]

    log.rows_written = len(df)
    return df, log


def save_clean_csv(df: pd.DataFrame, path: str | Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    out = df.copy()
    out["order_date"] = out["order_date"].dt.strftime("%Y-%m-%d")
    out.to_csv(path, index=False)
    return path
