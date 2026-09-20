"""Part 2 - data cleansing for the MOJ sales extract.

Reads the raw extract, applies each cleansing step in turn, writes a log describing what every
step did to the data, and saves the result as cleansed_data.csv.

Usage:
    python clean_data.py
    python clean_data.py --input moj_sales_extract.xlsx --output cleansed_data.csv --log cleaning.log

Only counts are written to the log, never row values.
"""
import argparse
import itertools
import logging
import re
import sys

import numpy as np
import pandas as pd
from hijridate import Hijri

TEXT_COLS = ["RS_CITY_NAME", "NEIGHBORHOOD_NAME", "RS_TYPE", "new_use"]

# Outlier rule: a value is extreme if it lies more than OUTLIER_K IQRs outside the middle 50%
# of the data, measured on a log10 scale (prices and areas span several orders of magnitude).
OUTLIER_K = 3
# Optional manual override, (low, high). Leave as None to use the data-driven rule above.
PRICE_LIMITS = None   # e.g. (100, 100_000)
AREA_LIMITS = None    # e.g. (10, 50_000)

log = logging.getLogger("clean_data")


# ----------------------------------------------------------------------------- helpers
def spelling_key(text):
    """Collapse Arabic spelling variants of the same word onto one comparison key."""
    text = re.sub(r"[ً-ٰٟ]", "", text)     # diacritics
    text = text.replace("ـ", "")                      # tatweel
    text = re.sub(r"[أإآٱ]", "ا", text)                   # alef with hamza / madda / wasla
    text = text.replace("ة", "ه").replace("ى", "ي")       # ta marbuta, alef maqsura
    return re.sub(r"\s+", " ", text).strip().casefold()


# Values the source uses to mean "no value" (compared after spelling_key and removing spaces)
PLACEHOLDERS = {
    spelling_key(p).replace(" ", "")
    for p in ["لا يوجد", "لا توجد", "غير موجود", "غير محدد", "غير معروف",
              "null", "none", "nan", "n/a", "na", "-", "--"]
}


def parse_hijri(value):
    """'1442 - 03 - 02' -> Gregorian Timestamp, or NaT if missing / not a real Hijri date."""
    if pd.isna(value):
        return pd.NaT
    m = re.match(r"^\s*(\d{4})\D+(\d{1,2})\D+(\d{1,2})\s*$", str(value))
    if not m:
        return pd.NaT
    year, month, day = map(int, m.groups())
    try:
        return pd.Timestamp(Hijri(year, month, day).to_gregorian())
    except (ValueError, OverflowError):
        return pd.NaT


def log_fences(s, k=OUTLIER_K):
    """Lower/upper limits from the Tukey rule applied to log10(s)."""
    logs = np.log10(s[s > 0].dropna())
    q1, q3 = logs.quantile([.25, .75])
    return 10 ** (q1 - k * (q3 - q1)), 10 ** (q3 + k * (q3 - q1))


def summary(df):
    """Headline numbers used to show the effect of a step."""
    price, area = df["METER_PRICE"], df["AREA"]
    return (f"mean price {price.mean():,.2f} | median price {price.median():,.2f} | "
            f"total value {(price * area).sum():,.0f}")


def report_rows(before, after):
    removed = len(before) - len(after)
    pct = removed / len(before) * 100 if len(before) else 0
    log.info(f"    rows: {len(before):,} -> {len(after):,}  (removed {removed:,}, {pct:.2f}%)")
    log.info(f"    effect: {summary(after)}")


def as_bool(mask):
    return mask.fillna(False).astype(bool)


# ----------------------------------------------------------------------------- text cleaning
def tidy_text(s):
    """Remove square brackets, collapse repeated spaces, trim, and turn empty strings into missing."""
    s = s.astype("string")
    no_brackets = s.str.replace(r"[\[\]]", "", regex=True)
    tidy = no_brackets.str.replace(r"\s+", " ", regex=True).str.strip()
    tidy = tidy.mask(as_bool(tidy.eq("")))
    return tidy, {
        "values with square brackets": int(as_bool(no_brackets.ne(s)).sum()),
        "values with extra / padded spaces": int(as_bool(tidy.ne(no_brackets)).sum()),
    }


def blank_placeholders(s):
    """Turn 'no value' placeholders (e.g. لايوجد) and purely numeric codes into missing."""
    key = s.map(lambda x: spelling_key(x).replace(" ", ""), na_action="ignore")
    is_placeholder = as_bool(key.isin(PLACEHOLDERS))
    is_numeric = as_bool(s.str.fullmatch(r"\d+([.,]\d+)?"))
    cleaned = s.mask(is_placeholder | is_numeric)
    return cleaned, {"'no value' placeholders": int(is_placeholder.sum()),
                     "purely numeric values": int(is_numeric.sum())}


def unify_spelling(s):
    """Map spelling variants of the same name to its most frequent spelling."""
    key = s.map(spelling_key, na_action="ignore")
    freq = (pd.DataFrame({"raw": s, "key": key}).dropna()
            .groupby(["key", "raw"]).size().rename("n").reset_index())
    canonical = (freq.sort_values(["key", "n", "raw"], ascending=[True, False, True])
                 .drop_duplicates("key").set_index("key")["raw"])
    unified = key.map(canonical).astype("string")
    return unified, {"names with more than one spelling": int((freq.groupby("key")["raw"].nunique() > 1).sum()),
                     "values rewritten to the main spelling": int(as_bool(unified.ne(s)).sum())}


# ----------------------------------------------------------------------------- pipeline
def clean(df):
    df = df.copy()
    n = itertools.count(1)
    log.info(f"Raw data: {len(df):,} rows, {df.shape[1]} columns")
    log.info("")

    # -- numbers ----------------------------------------------------------------------
    log.info(f"[Step {next(n)}] Convert AREA and METER_PRICE to numbers")
    for col in ["AREA", "METER_PRICE"]:
        converted = pd.to_numeric(df[col], errors="coerce")
        log.info(f"    {col}: {int((df[col].notna() & converted.isna()).sum()):,} non-numeric values set to empty")
        df[col] = converted
    log.info(f"    effect on raw data: {summary(df)}")
    log.info("")

    # -- exact duplicates -------------------------------------------------------------
    log.info(f"[Step {next(n)}] Drop rows identical in every column (exact duplicates)")
    before = df
    df = df.drop_duplicates()
    report_rows(before, df)
    log.info("")

    # -- text -------------------------------------------------------------------------
    for title, func in [
        ("Remove square brackets, extra spaces and empty strings from text columns", tidy_text),
        ("Set 'no value' placeholders and purely numeric codes to missing", blank_placeholders),
        ("Unify Arabic spelling variants (hamza, ta marbuta / ha, alef maqsura / ya) to the most frequent spelling",
         unify_spelling),
    ]:
        log.info(f"[Step {next(n)}] {title}")
        for col in TEXT_COLS:
            df[col], counts = func(df[col])
            log.info(f"    {col}: " + "; ".join(f"{v:,} {k}" for k, v in counts.items()))
        log.info("    (no rows removed)")
        log.info("")

    # -- exact duplicates created by the text clean-up --------------------------------
    log.info(f"[Step {next(n)}] Drop exact duplicates that only appeared after the text clean-up")
    before = df
    df = df.drop_duplicates()
    report_rows(before, df)
    log.info("")

    # -- duplicates we do NOT remove: report only ------------------------------------
    log.info(f"[Step {next(n)}] Check SERIAL duplicates (reported only, nothing removed)")
    has_serial = df["SERIAL"].notna()
    shares_serial = has_serial & df["SERIAL"].duplicated(keep=False)
    log.info(f"    rows sharing a SERIAL with another row: {int(shares_serial.sum()):,} "
             f"({int(df.loc[shares_serial, 'SERIAL'].nunique()):,} distinct SERIALs)")
    log.info(f"    rows with a missing SERIAL: {int((~has_serial).sum()):,}")
    log.info(f"    rows identical in everything except SERIAL: "
             f"{int(df.duplicated(subset=[c for c in df.columns if c != 'SERIAL']).sum()):,} extra copies")
    log.info("    Kept on purpose: a deed can cover several properties and there is no deed identifier,")
    log.info("    so these rows may be genuine separate records. Recorded as an open question.")
    log.info("")

    # -- dates ------------------------------------------------------------------------
    log.info(f"[Step {next(n)}] Convert Hijri deed date to Gregorian; drop rows with a missing or invalid date")
    mapping = {value: parse_hijri(value) for value in df["DEED_DATE_H"].dropna().unique()}
    df["deed_date"] = df["DEED_DATE_H"].map(mapping)
    log.info(f"    missing date text: {int(df['DEED_DATE_H'].isna().sum()):,}; "
             f"present but not a valid Hijri date: {int((df['DEED_DATE_H'].notna() & df['deed_date'].isna()).sum()):,}")
    before = df
    df = df[df["deed_date"].notna()]
    report_rows(before, df)
    log.info(f"    Gregorian range kept: {df['deed_date'].min().date()} to {df['deed_date'].max().date()}")
    log.info("")

    # -- missing / zero price or area -------------------------------------------------
    log.info(f"[Step {next(n)}] Drop rows with a missing, zero or negative METER_PRICE or AREA")
    bad_price = ~(df["METER_PRICE"] > 0)
    bad_area = ~(df["AREA"] > 0)
    log.info(f"    bad price: {int(bad_price.sum()):,}; bad area: {int(bad_area.sum()):,}; "
             f"both: {int((bad_price & bad_area).sum()):,}")
    before = df
    df = df[~(bad_price | bad_area)]
    report_rows(before, df)
    log.info("")

    # -- outliers ---------------------------------------------------------------------
    log.info(f"[Step {next(n)}] Detect and remove outliers in METER_PRICE and AREA "
             f"(log10 scale, {OUTLIER_K} x IQR beyond the middle 50%)")
    price_lo, price_hi = PRICE_LIMITS or log_fences(df["METER_PRICE"])
    area_lo, area_hi = AREA_LIMITS or log_fences(df["AREA"])
    log.info(f"    METER_PRICE kept range: {price_lo:,.2f} to {price_hi:,.2f} SAR per m2"
             f"{' (manual override)' if PRICE_LIMITS else ''}")
    log.info(f"    AREA kept range: {area_lo:,.2f} to {area_hi:,.2f} m2{' (manual override)' if AREA_LIMITS else ''}")
    out_price = ~df["METER_PRICE"].between(price_lo, price_hi)
    out_area = ~df["AREA"].between(area_lo, area_hi)
    log.info(f"    outside price range: {int(out_price.sum()):,}; outside area range: {int(out_area.sum()):,}; "
             f"both: {int((out_price & out_area).sum()):,}")
    before = df
    df = df[~(out_price | out_area)]
    report_rows(before, df)
    log.info("")

    # -- derived columns --------------------------------------------------------------
    log.info(f"[Step {next(n)}] Add derived columns: value = AREA x METER_PRICE (SAR) and quarter (e.g. 2020Q1)")
    df = df.assign(value=df["AREA"] * df["METER_PRICE"],
                   quarter=df["deed_date"].dt.to_period("Q").astype(str))
    try:
        df["SERIAL"] = df["SERIAL"].astype("Int64")
    except (TypeError, ValueError):
        log.info("    SERIAL left as-is (not all whole numbers)")
    log.info("")

    df = df.sort_values(["deed_date", "SERIAL"], na_position="last").reset_index(drop=True)
    columns = ["SERIAL", "DEED_DATE_H", "deed_date", "quarter", *TEXT_COLS, "AREA", "METER_PRICE", "value"]
    return df[columns]


def final_report(raw_rows, df):
    log.info("=" * 70)
    log.info(f"Final data: {len(df):,} rows kept of {raw_rows:,} ({len(df) / raw_rows * 100:.1f}%)")
    log.info(f"    {summary(df)}")
    log.info("    Missing values left in text columns (kept, not removed):")
    for col in TEXT_COLS:
        log.info(f"        {col}: {int(df[col].isna().sum()):,}")
    log.info("    Rows per quarter:")
    for quarter, count in df["quarter"].value_counts().sort_index().items():
        log.info(f"        {quarter}: {count:,}")


def main():
    parser = argparse.ArgumentParser(description="Cleanse the MOJ sales extract.")
    parser.add_argument("--input", default="moj_sales_extract.xlsx")
    parser.add_argument("--output", default="cleansed_data.csv")
    parser.add_argument("--log", default="cleaning.log")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(message)s", handlers=[
        logging.FileHandler(args.log, mode="w", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ])
    raw = pd.read_excel(args.input)
    cleaned = clean(raw)
    final_report(len(raw), cleaned)
    # utf-8-sig so Excel opens the Arabic text correctly
    cleaned.to_csv(args.output, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    log.info(f"Saved {args.output}")


if __name__ == "__main__":
    main()
