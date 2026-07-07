from __future__ import annotations

import argparse
import json
import logging
import os
import random
import sys
import time
from pathlib import Path
from typing import Any, Callable, TypeVar

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DEFAULT_INPUT = Path("output/bank_statement.xlsx")
DEFAULT_CREDENTIALS = Path("credentials.json")

LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

EXPECTED_COLUMNS = [
    "Source PDF",
    "Transaction Date",
    "Value Date",
    "Description",
    "Cheque No/Ref",
    "Deposits",
    "Withdrawals",
    "Running Balance",
]

# Columns used for cross-PDF deduplication
UNIQUE_KEY_COLUMNS = [
    "Transaction Date",
    "Description",
    "Deposits",
    "Withdrawals",
]

MASTER_SHEET_ID = "1B7z7GKp6jPEj0-HjXb9uxL9q5IMueLYTyq6jUYJEZoQ"
MASTER_WORKSHEET_NAME = "Bank_Statement_Master"

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("upload_to_sheets")


# ---------------------------------------------------------------------------
# Google Auth
# ---------------------------------------------------------------------------

GOOGLE_CREDENTIALS_ENV_VAR = "GOOGLE_CREDENTIALS_JSON"


def get_gspread_client(credentials_path: Path) -> gspread.Client:
    """Build an authorized gspread client.

    Tries the local credentials file first (the normal case when running
    on a machine/server with the file present). If that file doesn't
    exist, falls back to the GOOGLE_CREDENTIALS_JSON environment variable
    (the full service-account JSON as a string) — needed for deployments
    such as Vercel, where a secret file can't be committed to the repo
    or placed on a read-only filesystem.
    """
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    if credentials_path.exists():
        creds = Credentials.from_service_account_file(
            str(credentials_path),
            scopes=scope
        )
        return gspread.authorize(creds)

    env_value = os.environ.get(GOOGLE_CREDENTIALS_ENV_VAR)
    if env_value:
        try:
            info = json.loads(env_value)
        except json.JSONDecodeError as exc:
            raise ValueError(
                f"{GOOGLE_CREDENTIALS_ENV_VAR} environment variable is not valid JSON: {exc}"
            )
        creds = Credentials.from_service_account_info(info, scopes=scope)
        return gspread.authorize(creds)

    raise FileNotFoundError(
        f"Credentials file not found: {credentials_path}, and "
        f"{GOOGLE_CREDENTIALS_ENV_VAR} environment variable is not set."
    )


# ---------------------------------------------------------------------------
# Get or Create the Master Worksheet (never clear it)
# ---------------------------------------------------------------------------

def get_or_create_master_worksheet(
    client: gspread.Client,
) -> tuple[gspread.Spreadsheet, gspread.Worksheet]:
    """Open the master spreadsheet and return the master worksheet.

    If the worksheet does not exist yet, create it and write the header row.
    Existing data is NEVER cleared.
    """

    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    existing_titles = [ws.title for ws in spreadsheet.worksheets()]

    if MASTER_WORKSHEET_NAME in existing_titles:
        log.info("Master worksheet exists. Reusing: %s", MASTER_WORKSHEET_NAME)
        worksheet = spreadsheet.worksheet(MASTER_WORKSHEET_NAME)
    else:
        log.info("Creating master worksheet: %s", MASTER_WORKSHEET_NAME)
        worksheet = spreadsheet.add_worksheet(
            title=MASTER_WORKSHEET_NAME,
            rows="5000",
            cols="20",
        )
        # Write header row on brand-new sheet
        worksheet.append_row(EXPECTED_COLUMNS, value_input_option="RAW")

    return spreadsheet, worksheet


# ---------------------------------------------------------------------------
# Load existing sheet data into a DataFrame
# ---------------------------------------------------------------------------

def load_existing_data(worksheet: gspread.Worksheet) -> pd.DataFrame:
    """Read all records from the worksheet into a pandas DataFrame.

    Returns an empty DataFrame (with EXPECTED_COLUMNS columns) if the sheet
    has no data rows (only a header or completely empty).
    """
    try:
        records = worksheet.get_all_records()
    except Exception:
        records = []

    if not records:
        return pd.DataFrame(columns=EXPECTED_COLUMNS)

    df = pd.DataFrame(records)
    
    if "Transaction Date" in df.columns:
        df["Transaction Date"] = df["Transaction Date"].astype(str).str.strip()
    if "Description" in df.columns:
        df["Description"] = df["Description"].astype(str).str.strip().str.upper()
    for num_col in ["Deposits", "Withdrawals"]:
        if num_col in df.columns:
            cleaned = df[num_col].astype(str).str.replace(",", "", regex=False)
            df[num_col] = pd.to_numeric(cleaned, errors="coerce").fillna(0.0)
            
    return df


# ---------------------------------------------------------------------------
# Ensure header row exists
# ---------------------------------------------------------------------------

def ensure_header_row(worksheet: gspread.Worksheet) -> None:
    """If the worksheet is completely empty, write the header row."""
    first_row = worksheet.row_values(1)
    if not first_row or all(v.strip() == "" for v in first_row):
        worksheet.update(range_name="A1", values=[EXPECTED_COLUMNS])
        log.info("Wrote header row to empty worksheet.")


# ---------------------------------------------------------------------------
# Data Validation & Normalization
# ---------------------------------------------------------------------------

def validate_and_normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Apply all data validation and normalization rules to the DataFrame.

    Steps:
      1. Enforce expected column order
      2. Drop fully blank rows
      3. Reject rows missing Transaction Date or Description
      4. Normalize date columns to DD-MMM-YYYY
      5. Normalize numeric columns (remove commas, convert to numeric, fill NaN with 0)
    """

    # 1. Enforce column order ------------------------------------------------
    df = df.reindex(columns=EXPECTED_COLUMNS)

    # 2. Remove fully blank rows ---------------------------------------------
    df.dropna(how="all", inplace=True)

    # 3. Reject incomplete rows (missing Transaction Date or Description) ----
    df = df[
        df["Transaction Date"].notna()
        & (df["Transaction Date"].astype(str).str.strip() != "")
        & df["Description"].notna()
        & (df["Description"].astype(str).str.strip() != "")
    ].copy()

    if df.empty:
        return df

    # 4. Normalize date columns to DD-MMM-YYYY --------------------------------
    for date_col in ["Transaction Date", "Value Date"]:
        if date_col in df.columns:
            parsed = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
            # Keep original value where parsing failed
            formatted = parsed.dt.strftime("%d-%b-%Y")
            df[date_col] = formatted.where(parsed.notna(), df[date_col]).astype(str).str.strip()

    # 5. Normalize numeric columns -------------------------------------------
    for num_col in ["Deposits", "Withdrawals", "Running Balance"]:
        if num_col in df.columns:
            # Remove commas then convert
            cleaned = df[num_col].astype(str).str.replace(",", "", regex=False)
            df[num_col] = pd.to_numeric(cleaned, errors="coerce").fillna(0.0)

    if "Description" in df.columns:
        df["Description"] = df["Description"].astype(str).str.strip().str.upper()

    df.reset_index(drop=True, inplace=True)
    return df


# ---------------------------------------------------------------------------
# Append unique rows (never overwrite)
# ---------------------------------------------------------------------------

def append_unique_rows(
    worksheet: gspread.Worksheet,
    df: pd.DataFrame,
) -> int:
    """Append rows to the bottom of the worksheet.

    Returns the number of rows appended.
    """
    if df.empty:
        return 0

    df_out = df.reindex(columns=EXPECTED_COLUMNS).fillna("").astype(str)

    rows = df_out.values.tolist()

    worksheet.append_rows(rows, value_input_option="RAW")

    return len(rows)


# ---------------------------------------------------------------------------
# Main Upload
# ---------------------------------------------------------------------------

def upload_to_sheets(
    input_path: Path,
    credentials_path: Path,
    source_pdf_name: str,
) -> dict:
    """Upload extracted bank-statement Excel to the master Google Sheet.

    * Uses a SINGLE worksheet: Bank_Statement_Master
    * Adds a "Source PDF" column to track file origin
    * Validates and normalizes data before upload
    * Deduplicates via concat + drop_duplicates for accurate metrics
    * Appends only unique new rows — never overwrites

    Returns:
        The metrics dict (total_rows, new_rows, duplicates_skipped,
        sheet_url) — callers that import this function directly (e.g.
        run_pipeline.py) can use the return value instead of parsing the
        printed JSON line, which remains for CLI/subprocess backward
        compatibility.
    """
    if not input_path.exists():
        raise FileNotFoundError(f"Excel file not found: {input_path}")

    # ── Read the newly extracted Excel ──────────────────────────────────────
    df = pd.read_excel(str(input_path))

    if df.empty:
        metrics = {
            "total_rows": 0,
            "new_rows": 0,
            "duplicates_skipped": 0,
            "sheet_url": "",
        }
        print(json.dumps(metrics), flush=True)
        return metrics

    # ── Add Source PDF column ───────────────────────────────────────────────
    df.insert(0, "Source PDF", source_pdf_name)

    # ── Validate and normalize data ────────────────────────────────────────
    df = validate_and_normalize(df)
    total_rows = len(df)
    log.info("Rows after validation: %d", total_rows)

    # ── Remove B/F rows ────────────────────────────────────────────────────
    bf_mask = df["Description"].astype(str).str.strip().str.upper() == "B/F"
    bf_skipped = int(bf_mask.sum())
    df = df[~bf_mask].copy()

    if df.empty:
        metrics = {
            "total_rows": total_rows,
            "new_rows": 0,
            "duplicates_skipped": bf_skipped,
            "sheet_url": "",
        }
        print(json.dumps(metrics), flush=True)
        return metrics

    # ── Connect to Google Sheets ───────────────────────────────────────────
    client = get_gspread_client(credentials_path)
    spreadsheet, worksheet = get_or_create_master_worksheet(client)

    # ── Ensure header row exists ───────────────────────────────────────────
    ensure_header_row(worksheet)

    # ── Load existing sheet data ───────────────────────────────────────────
    existing_df = load_existing_data(worksheet)
    existing_count = len(existing_df)
    log.info("Existing rows in master sheet: %d", existing_count)

    # ── Left anti-join logic ───────────────────────────────────────────────
    new_unique_df = df.merge(
        existing_df[UNIQUE_KEY_COLUMNS].drop_duplicates(),
        on=UNIQUE_KEY_COLUMNS,
        how="left",
        indicator=True
    )
    
    new_unique_df = new_unique_df[new_unique_df["_merge"] == "left_only"].drop(columns=["_merge"])
    new_unique_df = new_unique_df.drop_duplicates(subset=UNIQUE_KEY_COLUMNS, keep="first")
    
    new_rows = len(new_unique_df)
    duplicates_skipped = (total_rows - new_rows)
    
    log.info(
        "Dedup: existing=%d  new_rows=%d  dupes_skipped=%d",
        existing_count, new_rows, duplicates_skipped,
    )

    if new_rows > 0:
        appended = append_unique_rows(worksheet, new_unique_df)
        log.info("Appended %d new rows to %s", appended, MASTER_WORKSHEET_NAME)
    else:
        log.info("No new rows to append — all duplicates.")

    # ── Build sheet URL ────────────────────────────────────────────────────
    sheet_url = (
        f"https://docs.google.com/spreadsheets/d/"
        f"{spreadsheet.id}/edit#gid={worksheet.id}"
    )

    metrics = {
        "total_rows": total_rows,
        "new_rows": new_rows,
        "duplicates_skipped": duplicates_skipped,
        "sheet_url": sheet_url,
    }

    print(json.dumps(metrics), flush=True)
    return metrics


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Upload bank statement data to Google Sheets (master worksheet)."
    )

    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        default=DEFAULT_INPUT,
    )

    parser.add_argument(
        "-c",
        "--credentials",
        type=Path,
        default=DEFAULT_CREDENTIALS,
    )

    parser.add_argument(
        "-s",
        "--source-pdf",
        default="unknown.pdf",
        help="Original PDF filename for the Source PDF column.",
    )

    # Keep --sheet-title for backward compat but it is now ignored
    parser.add_argument(
        "-t",
        "--sheet-title",
        default=MASTER_WORKSHEET_NAME,
        help="(Ignored) Worksheet name is always Bank_Statement_Master.",
    )

    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    try:
        upload_to_sheets(
            input_path=args.input,
            credentials_path=args.credentials,
            source_pdf_name=args.source_pdf,
        )
    except Exception as exc:
        log.exception(exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())