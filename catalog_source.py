"""
Loads the product catalog from either:
  - a Google Sheet (via its public CSV export link — no credentials needed
    as long as the sheet is shared as "Anyone with the link can view"), or
  - an uploaded .xlsx file.

Required columns: SKU, Barcode (case-insensitive, any spacing).
Short name is OPTIONAL -- if your sheet doesn't have one, a short name is
auto-generated from the SKU text (same abbreviation logic used to build
barcode_filled.xlsx earlier). If your sheet does have a "Short name"
column, that's used as-is instead.
"""

import os
import re
from pathlib import Path

import pandas as pd
from shortname_barcode_gen import build_short_name, build_label_short_name

REQUIRED_COLUMNS = ["sku", "barcode"]
OPTIONAL_SHORT_NAME_COL = "short name"
# "SKU Barcode" sheet calls the product-name column "Item"
COLUMN_ALIASES = {"item": "sku", "item name": "sku"}

# Live, auto-updating SKU Barcode sheet -- the dashboard's default catalog.
# Kept out of the (public) repo: set SKU_SHEET_URL in .streamlit/secrets.toml
# locally, or in the app's Secrets on Streamlit Cloud (app.py reads it).
DEFAULT_SHEET_URL = os.environ.get("SKU_SHEET_URL", "")
# Local copy (not in the repo), used only if the Google Sheet can't be reached
DEFAULT_CATALOG_PATH = Path(__file__).with_name("SKU Barcode.xlsx")


def sheet_url_to_csv_export(url_or_id: str, gid: str = "0") -> str:
    """
    Accepts either a full Google Sheets URL or a bare sheet ID, returns the
    CSV export URL. If the URL already contains a gid, that gid is used;
    otherwise defaults to the first sheet (gid=0).
    """
    url_or_id = url_or_id.strip()

    id_match = re.search(r"/d/([a-zA-Z0-9-_]+)", url_or_id)
    sheet_id = id_match.group(1) if id_match else url_or_id

    gid_match = re.search(r"[?&#]gid=([0-9]+)", url_or_id)
    if gid_match:
        gid = gid_match.group(1)

    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "sku" not in df.columns:
        df = df.rename(columns=COLUMN_ALIASES)
    return df


def _validate_and_extract(df: pd.DataFrame) -> pd.DataFrame:
    df = _normalize_columns(df)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing expected column(s): {', '.join(missing)}. "
            f"Found columns: {', '.join(df.columns)}"
        )

    out = df[REQUIRED_COLUMNS].rename(columns={"sku": "sku", "barcode": "barcode"})
    out = out.dropna(subset=["sku"])
    out["sku"] = out["sku"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    # strip stray trailing dashes/spaces typed into the sheet (e.g. "6953175761310-")
    out["barcode"] = out["barcode"].fillna("").astype(str).str.strip().str.rstrip("- ")

    sheet_short = (df[OPTIONAL_SHORT_NAME_COL].fillna("").astype(str).str.strip()
                   if OPTIONAL_SHORT_NAME_COL in df.columns else pd.Series("", index=df.index))
    blank = sheet_short == ""
    # short_name: 84L barcode labels (<= 26 chars to fit the 46 mm label)
    out["short_name"] = sheet_short.where(~blank, out["sku"].apply(build_short_name))
    # label_short_name: 21L MRP labels -- same sheet value, but blanks get
    # the fuller, readable generated name (e.g. "Kiwi Ears Belle Black 3.5 Mic")
    out["label_short_name"] = sheet_short.where(~blank, out["sku"].apply(build_label_short_name))

    # drop rows with no barcode -- not usable for printing
    out = out[out["barcode"] != ""]
    # same item listed twice with different barcodes is kept (both are real
    # barcodes); only exact repeats are dropped
    out = out.drop_duplicates(subset=["sku", "barcode"], keep="first").reset_index(drop=True)
    return out[["sku", "short_name", "label_short_name", "barcode"]]


def load_from_google_sheet(url_or_id: str) -> pd.DataFrame:
    csv_url = sheet_url_to_csv_export(url_or_id)
    # dtype=str is essential: barcodes are zero-padded digit strings
    # (e.g. "00110001") and pandas will otherwise silently infer an
    # integer column and drop the leading zeros, corrupting every
    # barcode that doesn't already match what's been printed.
    df = pd.read_csv(csv_url, dtype=str)
    return _validate_and_extract(df)


def load_from_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype=str)
    return _validate_and_extract(df)
