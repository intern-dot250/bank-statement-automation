"""Phase 2: Classify transactions in the master Google Sheet.

Reads every row from the existing master worksheet (the same sheet that
upload_to_sheets.py writes to), assigns a business ``Head`` and a
human-readable ``Narration`` to each transaction, and writes both values
back into the SAME row. No rows are appended or duplicated.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

import gspread
from gspread.utils import rowcol_to_a1

from heads import get_head
from narration import generate_narration
from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    MASTER_WORKSHEET_NAME,
    get_gspread_client,
)

LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("classify_transactions")

# Columns this script is responsible for adding/populating.
HEAD_COLUMN = "Head"
NARRATION_COLUMN = "Narration"


# ---------------------------------------------------------------------------
# Worksheet access
# ---------------------------------------------------------------------------

def open_master_worksheet(
    client: gspread.Client,
    sheet_id: str,
    worksheet_name: str,
) -> gspread.Worksheet:
    """Open the existing master worksheet.

    Raises:
        gspread.exceptions.WorksheetNotFound: If the worksheet does not exist.
            This script only updates an existing sheet; it never creates one.
    """
    spreadsheet = client.open_by_key(sheet_id)
    return spreadsheet.worksheet(worksheet_name)


# ---------------------------------------------------------------------------
# Header management
# ---------------------------------------------------------------------------

def ensure_head_narration_columns(worksheet: gspread.Worksheet) -> tuple[list[str], int, int]:
    """Ensure Head and Narration columns exist in the header row.

    If either column is missing, it is appended to the end of the header row.

    Returns:
        A tuple of (updated_header_row, head_col_index, narration_col_index),
        where indices are 1-based.
    """
    header_row = worksheet.row_values(1)

    if not header_row:
        raise ValueError("Worksheet has no header row; cannot classify an empty sheet.")

    if HEAD_COLUMN not in header_row:
        header_row.append(HEAD_COLUMN)
        log.info("Head column not found — adding it.")

    if NARRATION_COLUMN not in header_row:
        header_row.append(NARRATION_COLUMN)
        log.info("Narration column not found — adding it.")

    worksheet.update(range_name="A1", values=[header_row])

    head_col_index = header_row.index(HEAD_COLUMN) + 1
    narration_col_index = header_row.index(NARRATION_COLUMN) + 1

    return header_row, head_col_index, narration_col_index


# ---------------------------------------------------------------------------
# Row helpers
# ---------------------------------------------------------------------------

def _get_cell(row: list[str], header_row: list[str], column_name: str) -> str:
    """Safely read a cell value by column name, tolerating short/ragged rows."""
    if column_name not in header_row:
        return ""
    index = header_row.index(column_name)
    if index >= len(row):
        return ""
    return row[index].strip()


def _to_float(value: str) -> float:
    """Safely convert a raw sheet cell value to a float, defaulting to 0.0."""
    cleaned = value.replace(",", "").strip()
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _parse_amount(deposits_raw: str, withdrawals_raw: str) -> float:
    """Return Deposits if positive, otherwise Withdrawals, as a float."""
    deposits = _to_float(deposits_raw)
    if deposits > 0:
        return deposits
    return _to_float(withdrawals_raw)


def _is_row_empty(row: list[str]) -> bool:
    """A row is empty if every cell is blank."""
    return all(cell.strip() == "" for cell in row)


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------

def classify_rows(
    worksheet: gspread.Worksheet,
    header_row: list[str],
    head_col_index: int,
    narration_col_index: int,
) -> int:
    """Classify each data row and write Head + Narration back to the sheet.

    Skips:
      * Fully empty rows
      * Rows without a Description
      * Rows that already have both Head and Narration filled in (idempotent)

    Returns:
        The number of rows updated.
    """
    all_values = worksheet.get_all_values()
    data_rows = all_values[1:]  # exclude header

    updates: list[gspread.cell.Cell] = []
    updated_count = 0

    for offset, row in enumerate(data_rows):
        sheet_row_number = offset + 2  # +1 for header, +1 for 1-based index

        if _is_row_empty(row):
            continue

        description = _get_cell(row, header_row, "Description")
        if not description:
            log.debug("Skipping row %d: no Description.", sheet_row_number)
            continue

        existing_head = _get_cell(row, header_row, HEAD_COLUMN)
        existing_narration = _get_cell(row, header_row, NARRATION_COLUMN)
        if existing_head and existing_narration:
            log.debug("Skipping row %d: already classified.", sheet_row_number)
            continue

        deposits_raw = _get_cell(row, header_row, "Deposits")
        withdrawals_raw = _get_cell(row, header_row, "Withdrawals")
        deposits = _to_float(deposits_raw)
        withdrawals = _to_float(withdrawals_raw)
        amount = _parse_amount(deposits_raw, withdrawals_raw)

        head = get_head(description, deposits, withdrawals)
        narration = generate_narration(description, head, amount)

        updates.append(
            gspread.cell.Cell(
                row=sheet_row_number,
                col=head_col_index,
                value=head,
            )
        )
        updates.append(
            gspread.cell.Cell(
                row=sheet_row_number,
                col=narration_col_index,
                value=narration,
            )
        )
        updated_count += 1

    if updates:
        worksheet.update_cells(updates, value_input_option="RAW")
        log.info("Updated %d row(s) with Head + Narration.", updated_count)
    else:
        log.info("No rows required classification.")

    return updated_count


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def classify_transactions(
    credentials_path: Path,
    sheet_id: str = MASTER_SHEET_ID,
    worksheet_name: str = MASTER_WORKSHEET_NAME,
) -> int:
    """Classify all unclassified transactions in the master Google Sheet.

    Args:
        credentials_path: Path to the Google service-account credentials JSON.
        sheet_id: Spreadsheet ID of the master sheet.
        worksheet_name: Worksheet/tab name within the master sheet.

    Returns:
        Number of rows updated.
    """
    if not credentials_path.exists():
        raise FileNotFoundError(f"Credentials file not found: {credentials_path}")

    client = get_gspread_client(credentials_path)
    worksheet = open_master_worksheet(client, sheet_id, worksheet_name)

    header_row, head_col_index, narration_col_index = ensure_head_narration_columns(worksheet)

    return classify_rows(worksheet, header_row, head_col_index, narration_col_index)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify transactions in the master Google Sheet (Head + Narration)."
    )

    parser.add_argument(
        "-c",
        "--credentials",
        type=Path,
        default=DEFAULT_CREDENTIALS,
    )

    parser.add_argument(
        "--sheet-id",
        default=MASTER_SHEET_ID,
        help="Override the master spreadsheet ID.",
    )

    parser.add_argument(
        "--worksheet-name",
        default=MASTER_WORKSHEET_NAME,
        help="Override the worksheet/tab name.",
    )

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        updated = classify_transactions(
            credentials_path=args.credentials,
            sheet_id=args.sheet_id,
            worksheet_name=args.worksheet_name,
        )
        log.info("Classification complete. Rows updated: %d", updated)
    except Exception as exc:
        log.exception(exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
