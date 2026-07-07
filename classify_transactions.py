"""Phase 2: Classify transactions in the master Google Sheet.

Reads every row from the existing master worksheet (the same sheet that
upload_to_sheets.py writes to), assigns a business ``Head`` and a
human-readable ``Narration`` to each transaction, and writes both values
back into the SAME row. No rows are appended or duplicated.
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Any, Optional

import gspread
from gspread.utils import rowcol_to_a1

from description_parser import parse_description
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
PROJECT_COLUMN = "Project"
HEAD_INCOME_TAX_COLUMN = "Head - Income Tax"
TYPE_RERA_IDW_COLUMN = "Type for RERA IDW"
TCP_HEAD_COLUMN = "TCP Head"

# All columns that must be present, in the order they're appended if missing.
CLASSIFICATION_COLUMNS = [
    HEAD_COLUMN,
    NARRATION_COLUMN,
    PROJECT_COLUMN,
    HEAD_INCOME_TAX_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
]

SCRIPT_DIR = Path(__file__).resolve().parent
HEAD_MAPPING_PATH = SCRIPT_DIR / "config" / "head_mapping.json"

# Value written for any of the 4 mapped fields when head_mapping.json
# itself marks that field "?" (non-deterministic), or when the Head has
# no entry in head_mapping.json at all. Never invented — "?" is the
# literal value head_mapping.json already uses for this exact situation.
UNKNOWN_MAPPING_VALUE = "?"


# ---------------------------------------------------------------------------
# head_mapping.json loading (cached — loaded from disk at most once per process)
# ---------------------------------------------------------------------------

_head_mapping_cache: Optional[dict[str, dict[str, Any]]] = None


def _get_head_mapping() -> dict[str, dict[str, Any]]:
    """Load config/head_mapping.json, caching it after the first read.

    Returns:
        The "heads" mapping dict from head_mapping.json (Head name ->
        {"project", "head_income_tax", "type_rera_idw", "tcp_head"}).
        Returns an empty dict (logged as an error) if the file cannot be
        loaded — callers must treat every Head as unmapped in that case,
        never fabricating values.
    """
    global _head_mapping_cache
    if _head_mapping_cache is None:
        try:
            with open(HEAD_MAPPING_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            _head_mapping_cache = data.get("heads", {})
            log.debug("Loaded head_mapping.json (%d heads).", len(_head_mapping_cache))
        except Exception as exc:
            log.error(
                "Could not load head_mapping.json: %s — Project/Head - Income Tax/"
                "Type for RERA IDW/TCP Head will be written as %r for every row.",
                exc, UNKNOWN_MAPPING_VALUE,
            )
            _head_mapping_cache = {}
    return _head_mapping_cache


def _lookup_head_mapping(head: str) -> dict[str, str]:
    """Look up the 4 mapped fields for a classified Head.

    Values are used EXACTLY as stored in head_mapping.json — never
    calculated, inferred, or modified. If the Head has no entry in
    head_mapping.json, every field is set to "?" (the same convention
    head_mapping.json itself uses for non-deterministic fields) and a
    warning is logged, rather than guessing a value.

    Returns:
        Dict with keys "project", "head_income_tax", "type_rera_idw",
        "tcp_head".
    """
    mapping = _get_head_mapping()
    entry = mapping.get(head)

    if entry is None:
        log.warning(
            "Head %r has no entry in head_mapping.json — writing %r for "
            "Project/Head - Income Tax/Type for RERA IDW/TCP Head.",
            head, UNKNOWN_MAPPING_VALUE,
        )
        return {
            "project": UNKNOWN_MAPPING_VALUE,
            "head_income_tax": UNKNOWN_MAPPING_VALUE,
            "type_rera_idw": UNKNOWN_MAPPING_VALUE,
            "tcp_head": UNKNOWN_MAPPING_VALUE,
        }

    return {
        "project": entry.get("project", UNKNOWN_MAPPING_VALUE),
        "head_income_tax": entry.get("head_income_tax", UNKNOWN_MAPPING_VALUE),
        "type_rera_idw": entry.get("type_rera_idw", UNKNOWN_MAPPING_VALUE),
        "tcp_head": entry.get("tcp_head", UNKNOWN_MAPPING_VALUE),
    }


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

def ensure_classification_columns(worksheet: gspread.Worksheet) -> tuple[list[str], dict[str, int]]:
    """Ensure all classification columns exist in the header row.

    Columns ensured: Head, Narration, Project, Head - Income Tax,
    Type for RERA IDW, TCP Head. Any that are missing are appended to the
    end of the header row (existing columns are never reordered or
    removed).

    Returns:
        A tuple of (updated_header_row, column_indices), where
        column_indices maps each column name in CLASSIFICATION_COLUMNS to
        its 1-based column index.
    """
    header_row = worksheet.row_values(1)

    if not header_row:
        raise ValueError("Worksheet has no header row; cannot classify an empty sheet.")

    for column_name in CLASSIFICATION_COLUMNS:
        if column_name not in header_row:
            header_row.append(column_name)
            log.info("%s column not found — adding it.", column_name)

    worksheet.update(range_name="A1", values=[header_row])

    column_indices = {name: header_row.index(name) + 1 for name in CLASSIFICATION_COLUMNS}

    return header_row, column_indices


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


def _safe_parse_description(description: str, sheet_row_number: int) -> dict | None:
    """Run parse_description() defensively so a parsing issue never
    blocks classification of a row.

    Returns:
        The parsed-fields dict on success (even if every field inside it
        is None — that just means no known pattern matched), or None if
        parse_description() itself raised. Callers must fall back to the
        raw description in either case; get_head()/generate_narration()
        already operate on the raw description string, so this fallback
        is automatic and no transaction is ever skipped.
    """
    try:
        parsed = parse_description(description)
    except Exception as exc:
        log.warning(
            "Row %d: description_parser raised %s — falling back to raw description.",
            sheet_row_number, exc,
        )
        return None

    if any(value is not None for value in parsed.values()):
        log.debug("Row %d: description parsed successfully: %s", sheet_row_number, parsed)
    else:
        log.debug(
            "Row %d: no known pattern matched description — falling back to raw description.",
            sheet_row_number,
        )

    return parsed


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------

def classify_rows(
    worksheet: gspread.Worksheet,
    header_row: list[str],
    column_indices: dict[str, int],
) -> int:
    """Classify each data row and write all classification columns back to
    the sheet: Head, Narration, Project, Head - Income Tax,
    Type for RERA IDW, TCP Head.

    Skips:
      * Fully empty rows
      * Rows without a Description
      * Rows that already have every classification column filled in
        (idempotent — this also lets previously Head/Narration-only rows
        get backfilled with the 4 new columns on the next run, since they
        won't yet have Project/Head - Income Tax/Type for RERA IDW/TCP Head)

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

        already_classified = all(
            _get_cell(row, header_row, column_name)
            for column_name in CLASSIFICATION_COLUMNS
        )
        if already_classified:
            log.debug("Skipping row %d: already fully classified.", sheet_row_number)
            continue

        # Structured parsing step (description_parser.py). Parsing is
        # advisory only: get_head()/generate_narration() below still take
        # the raw description string, so a parse failure or an
        # unrecognized description format never skips the row.
        _safe_parse_description(description, sheet_row_number)

        deposits_raw = _get_cell(row, header_row, "Deposits")
        withdrawals_raw = _get_cell(row, header_row, "Withdrawals")
        deposits = _to_float(deposits_raw)
        withdrawals = _to_float(withdrawals_raw)
        amount = _parse_amount(deposits_raw, withdrawals_raw)

        # Head classification and Narration generation are UNCHANGED.
        head = get_head(description, deposits, withdrawals)
        narration = generate_narration(description, head, amount)

        # New: look up the 4 additional fields from head_mapping.json,
        # using them exactly as stored — no calculation or inference.
        mapping = _lookup_head_mapping(head)

        row_values = {
            HEAD_COLUMN: head,
            NARRATION_COLUMN: narration,
            PROJECT_COLUMN: mapping["project"],
            HEAD_INCOME_TAX_COLUMN: mapping["head_income_tax"],
            TYPE_RERA_IDW_COLUMN: mapping["type_rera_idw"],
            TCP_HEAD_COLUMN: mapping["tcp_head"],
        }

        for column_name, value in row_values.items():
            updates.append(
                gspread.cell.Cell(
                    row=sheet_row_number,
                    col=column_indices[column_name],
                    value=value,
                )
            )
        updated_count += 1

    if updates:
        worksheet.update_cells(updates, value_input_option="RAW")
        log.info("Updated %d row(s) with full classification.", updated_count)
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
    # Credential resolution (file vs GOOGLE_CREDENTIALS_JSON env var fallback)
    # is handled entirely inside get_gspread_client() — no upfront existence
    # check here, since that would bypass the env var fallback it supports.
    client = get_gspread_client(credentials_path)
    worksheet = open_master_worksheet(client, sheet_id, worksheet_name)

    header_row, column_indices = ensure_classification_columns(worksheet)

    return classify_rows(worksheet, header_row, column_indices)


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
