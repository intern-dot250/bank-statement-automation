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
    "Credits",
    "Debits",
    "Balance",
    "Account Number",
]

# Columns used for cross-PDF deduplication
UNIQUE_KEY_COLUMNS = [
    "Transaction Date",
    "Description",
    "Credits",
    "Debits",
]

# Columns that hold rupee amounts — formatted as real numbers with
# 2-decimal precision. Google Sheets' number formatting doesn't support
# Indian-style 2-2-3 digit grouping (1,23,456.00) — it always renders
# comma groups in fixed 3-digit chunks regardless of pattern or locale
# (confirmed: neither a custom multi-comma pattern nor a bracket-conditional
# pattern nor an en-IN-equivalent locale changes this) — so these display
# with standard Western grouping (1,23,456.00 -> 123,456.00).
NUMERIC_FORMAT_COLUMNS = ["Credits", "Debits", "Balance"]
NUMERIC_CELL_FORMAT = {"type": "NUMBER", "pattern": "#,##0.00"}

MASTER_SHEET_ID = "1B7z7GKp6jPEj0-HjXb9uxL9q5IMueLYTyq6jUYJEZoQ"

# Worksheet/tab names that are reports, not per-account transaction data —
# excluded when combining data across all account tabs (e.g. for
# Summary/Final Report/Validation).
RESERVED_WORKSHEET_NAMES = {"Summary", "Final Report"}

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
# Per-account worksheets (one tab per account, no shared master sheet)
# ---------------------------------------------------------------------------

def get_account_worksheets(spreadsheet: gspread.Spreadsheet) -> list[gspread.Worksheet]:
    """Return every worksheet that holds per-account transaction data —
    i.e. every tab except the reserved report tabs (Summary, Final Report)."""
    return [ws for ws in spreadsheet.worksheets() if ws.title not in RESERVED_WORKSHEET_NAMES]


def load_combined_account_values(spreadsheet: gspread.Spreadsheet) -> list[list[str]]:
    """Read and combine every account worksheet's rows into one grid
    (header + data rows), for use by Summary/Final Report/Validation —
    which need one aggregated view across all accounts, now that there's
    no single master worksheet.

    Each account tab's rows are re-aligned into the first non-empty
    tab's column order (by header name), so minor column-order
    differences between tabs don't corrupt the combined data. Returns
    an empty list if there are no account worksheets/data at all.
    """
    combined_header: list[str] | None = None
    combined_rows: list[list[str]] = []

    for worksheet in get_account_worksheets(spreadsheet):
        values = worksheet.get_all_values()
        if not values:
            continue

        header, *rows = values
        if not rows:
            continue

        if combined_header is None:
            combined_header = header
            combined_rows.extend(rows)
            continue

        column_index = {name: i for i, name in enumerate(header)}
        for row in rows:
            combined_rows.append([
                row[column_index[col]] if col in column_index and column_index[col] < len(row) else ""
                for col in combined_header
            ])

    if combined_header is None:
        return []

    return [combined_header] + combined_rows


def build_account_worksheet_name(bank_name: str, account_number: str) -> str:
    """Build the per-account worksheet/tab name: '<Bank Name> - <last 4 digits>'."""
    last4 = account_number[-4:] if account_number else "0000"
    return f"{bank_name} - {last4}"


def get_or_create_account_worksheet(
    spreadsheet: gspread.Spreadsheet,
    worksheet_name: str,
) -> gspread.Worksheet:
    """Open (or create) the per-account worksheet. Never clears existing data —
    every new statement for this account just appends onto it."""
    existing_titles = [ws.title for ws in spreadsheet.worksheets()]

    if worksheet_name in existing_titles:
        return spreadsheet.worksheet(worksheet_name)

    log.info("Creating account worksheet: %s", worksheet_name)
    worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="5000", cols="20")
    worksheet.append_row(EXPECTED_COLUMNS, value_input_option="RAW")
    apply_numeric_format(worksheet)
    return worksheet


def apply_numeric_format(worksheet: gspread.Worksheet) -> None:
    """Format the Credits/Debits/Balance columns as real 2-decimal
    numbers (see NUMERIC_FORMAT_COLUMNS for why this is Western, not
    Indian, digit grouping). Applied to the whole column (not just
    existing rows), so every future row appended to this sheet is
    formatted too. Failures are logged but never raised — correct data
    with default number formatting is still useful even if the display
    formatting doesn't apply."""
    header = worksheet.row_values(1)
    try:
        for column_name in NUMERIC_FORMAT_COLUMNS:
            if column_name not in header:
                continue
            col_letter = gspread.utils.rowcol_to_a1(1, header.index(column_name) + 1).rstrip("0123456789")
            worksheet.format(f"{col_letter}2:{col_letter}", {"numberFormat": NUMERIC_CELL_FORMAT})
    except Exception as exc:
        log.warning("Could not apply numeric format to %s: %s", worksheet.title, exc)


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
    for num_col in ["Credits", "Debits"]:
        if num_col in df.columns:
            cleaned = df[num_col].astype(str).str.replace(",", "", regex=False)
            df[num_col] = pd.to_numeric(cleaned, errors="coerce").fillna(0.0)

    return df


# ---------------------------------------------------------------------------
# Ensure header row exists
# ---------------------------------------------------------------------------

def ensure_header_row(worksheet: gspread.Worksheet) -> None:
    """If the worksheet is completely empty, write the header row.

    If it already has a header but is missing newer columns (e.g.
    "Account Number", added after this sheet was first created), extend
    the header in place — existing data columns/rows are never touched.
    """
    first_row = worksheet.row_values(1)
    if not first_row or all(v.strip() == "" for v in first_row):
        worksheet.update(range_name="A1", values=[EXPECTED_COLUMNS])
        log.info("Wrote header row to empty worksheet.")
        return

    missing_columns = [c for c in EXPECTED_COLUMNS if c not in first_row]
    if missing_columns:
        start_col = len(first_row) + 1
        worksheet.update(
            range_name=gspread.utils.rowcol_to_a1(1, start_col),
            values=[missing_columns],
        )
        log.info("Extended header row with missing column(s): %s", missing_columns)


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
    for num_col in ["Credits", "Debits", "Balance"]:
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

    Credits/Debits/Balance are sent as actual numbers (not text) so
    Google Sheets' numeric-format grouping (applied to those
    columns — see apply_numeric_format()) actually displays;
    formatting a text string is a no-op. Every other column stays text,
    written under value_input_option="RAW" (Sheets never reinterprets
    string content under RAW, so this carries no formula-injection risk
    even though transaction descriptions are untrusted bank text).

    Returns:
        The number of rows appended.
    """
    if df.empty:
        return 0

    df_out = df.reindex(columns=EXPECTED_COLUMNS)
    for column_name in EXPECTED_COLUMNS:
        if column_name in NUMERIC_FORMAT_COLUMNS:
            df_out[column_name] = pd.to_numeric(df_out[column_name], errors="coerce").fillna(0.0)
        else:
            df_out[column_name] = df_out[column_name].fillna("").astype(str)

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
    account_number: str,
    bank_name: str,
) -> dict:
    """Upload extracted bank-statement Excel to that account's own Google
    Sheet worksheet/tab — e.g. "YES BANK - 2477", created automatically
    if it doesn't exist yet. There is no shared master worksheet; every
    account's transactions live only in its own tab.

    * Adds "Source PDF" and "Account Number" columns
    * Validates and normalizes data before upload
    * Deduplicates against that account's own existing rows
    * Appends only unique new rows — never overwrites

    Returns:
        The metrics dict (total_rows, new_rows, duplicates_skipped,
        sheet_url) — callers that import this function directly (e.g.
        run_pipeline.py) can use the return value instead of parsing the
        printed JSON line, which remains for CLI/subprocess backward
        compatibility.
    """
    if not account_number or not bank_name:
        raise ValueError(
            "account_number and bank_name are both required — every statement "
            "must be routed to a specific account's own worksheet."
        )

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

    # ── Add Source PDF / Account Number columns ─────────────────────────────
    df.insert(0, "Source PDF", source_pdf_name)
    df["Account Number"] = account_number

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
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)
    account_worksheet_name = build_account_worksheet_name(bank_name, account_number)
    worksheet = get_or_create_account_worksheet(spreadsheet, account_worksheet_name)

    # ── Ensure header row exists ───────────────────────────────────────────
    ensure_header_row(worksheet)

    # ── Load existing sheet data ───────────────────────────────────────────
    existing_df = load_existing_data(worksheet)
    existing_count = len(existing_df)
    log.info("Existing rows in %s: %d", account_worksheet_name, existing_count)

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
        log.info("Appended %d new rows to %s", appended, account_worksheet_name)
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
        description="Upload bank statement data to that account's Google Sheets worksheet/tab."
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

    parser.add_argument(
        "-a",
        "--account-number",
        required=True,
        help="Account number this statement belongs to (routes to that account's own tab).",
    )

    parser.add_argument(
        "-b",
        "--bank-name",
        required=True,
        help="Bank name (used in the account tab's name, e.g. 'YES BANK - 2477').",
    )

    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    try:
        upload_to_sheets(
            input_path=args.input,
            credentials_path=args.credentials,
            source_pdf_name=args.source_pdf,
            account_number=args.account_number,
            bank_name=args.bank_name,
        )
    except Exception as exc:
        log.exception(exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())