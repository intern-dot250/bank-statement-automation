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
from typing import Any, Optional

import gspread
from gspread.utils import rowcol_to_a1

from description_parser import parse_description
from heads import get_head
from narration import generate_narration
from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
)
import credentials_store

SCRIPT_DIR_FOR_RECORDS = Path(__file__).resolve().parent
_RECORDS_FALLBACK_PATH = SCRIPT_DIR_FOR_RECORDS / "records.json"

# Stage-pair -> Type for RERA IDW label, for internal transfers between two
# of our own tracked accounts. Only pairs confidently confirmed from the
# accounts team's reference sheet are included — any other pair (e.g.
# RERA <-> IDW, which the reference data shows using two DIFFERENT labels
# depending on transaction specifics we can't reliably tell apart from the
# description alone) is intentionally left unmapped, so it falls back to "?".
TRANSFER_STAGE_LABELS: dict[frozenset[str], str] = {
    frozenset({"Master", "Free"}): "Master to Free",
    frozenset({"Master", "RERA"}): "Master 2 RERA",
    frozenset({"Free", "IDW"}): "Free & IDW Loan",
}

# Description prefixes that indicate an incoming payment from an external
# party (as opposed to a transfer between our own tracked accounts).
_INCOMING_PAYMENT_PREFIXES = ("UPI/", "NEFT CR-", "IMPS/", "RTGS CR-", "NET-TPT-", "NET-")


_accounts_by_number_cache: Optional[dict[str, dict[str, Any]]] = None


def _get_accounts_by_number() -> dict[str, dict[str, Any]]:
    """Load account_credentials once per process, keyed by account_number."""
    global _accounts_by_number_cache
    if _accounts_by_number_cache is None:
        accounts = credentials_store.list_credentials(_RECORDS_FALLBACK_PATH)
        _accounts_by_number_cache = {
            acc["account_number"]: acc for acc in accounts if acc.get("account_number")
        }
    return _accounts_by_number_cache


def _find_counterparty_account(description: str, own_account_number: str) -> Optional[dict[str, Any]]:
    """If description mentions one of our OTHER tracked account numbers,
    return that account's record — this reliably signals an internal
    transfer between two of our own accounts (the account number is a
    much stronger signal than company-name matching, since our own
    company's name also legitimately appears in ordinary customer-payment
    descriptions as the beneficiary)."""
    for account_number, account in _get_accounts_by_number().items():
        if account_number != own_account_number and account_number in description:
            return account
    return None


def _looks_like_incoming_payment(description: str) -> bool:
    upper = description.strip().upper()
    return upper.startswith(_INCOMING_PAYMENT_PREFIXES)


def resolve_business_fields(
    account_number: str,
    description: str,
    deposits: float,
    withdrawals: float,
) -> dict[str, Any]:
    """Determine Head/Business Unit/Type for RERA IDW/TCP Head using the
    two most reliable, generalizable rules confirmed from the accounts
    team's reference sheet:

      1. Internal transfer between two of our own tracked accounts
         (detected via a counterparty account number appearing in the
         description) -> Head "Internal", TCP Head "Internal transfer",
         Business Unit = this account's own project, and a Type for
         RERA IDW label looked up by (this account's stage, counterparty's
         stage) when that specific pair is confidently known.
      2. An incoming payment (UPI/NEFT/IMPS/RTGS/NET-TPT) that ISN'T an
         internal transfer -> Head "Collection", TCP Head "Credit- no
         effect", Type for RERA IDW "Customer Collection".

    Anything else returns head=None (caller falls back to the existing
    get_head() heuristic) with business_unit/type_rera_idw/tcp_head all
    "?", per the explicit instruction to leave fields blank/unknown
    rather than guess.

    Returns:
        Dict with keys "head" (str or None), "business_unit",
        "type_rera_idw", "tcp_head".
    """
    accounts = _get_accounts_by_number()
    own_account = accounts.get(account_number, {})
    own_business_unit = own_account.get("business_unit") or UNKNOWN_MAPPING_VALUE
    own_stage = own_account.get("account_stage")

    counterparty = _find_counterparty_account(description, account_number)
    if counterparty is not None:
        counterparty_stage = counterparty.get("account_stage")
        type_rera_idw = UNKNOWN_MAPPING_VALUE
        if own_stage and counterparty_stage:
            type_rera_idw = TRANSFER_STAGE_LABELS.get(
                frozenset({own_stage, counterparty_stage}), UNKNOWN_MAPPING_VALUE
            )
        return {
            "head": "Internal",
            "business_unit": own_business_unit,
            "type_rera_idw": type_rera_idw,
            "tcp_head": "Internal transfer",
        }

    if deposits > 0 and _looks_like_incoming_payment(description):
        return {
            "head": "Collection",
            "business_unit": own_business_unit,
            "type_rera_idw": "Customer Collection",
            "tcp_head": "Credit- no effect",
        }

    return {
        "head": None,
        "business_unit": UNKNOWN_MAPPING_VALUE,
        "type_rera_idw": UNKNOWN_MAPPING_VALUE,
        "tcp_head": UNKNOWN_MAPPING_VALUE,
    }

LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("classify_transactions")

# Columns this script is responsible for adding/populating.
BUSINESS_UNIT_COLUMN = "Business Unit"
HEAD_COLUMN = "Head"
TYPE_RERA_IDW_COLUMN = "Type for RERA IDW"
TCP_HEAD_COLUMN = "TCP Head"
NARRATION_COLUMN = "Narration"

# All columns that must be present, in the order they're appended if missing.
# Matches the accounts department's own sheet format: Business Unit | Head |
# Type for RERA IDW | TCP Head | Narration.
CLASSIFICATION_COLUMNS = [
    BUSINESS_UNIT_COLUMN,
    HEAD_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
    NARRATION_COLUMN,
]

SCRIPT_DIR = Path(__file__).resolve().parent

# Value written for Business Unit/Type for RERA IDW/TCP Head whenever we
# aren't confident enough to fill them in from a known rule — never
# invented/guessed.
UNKNOWN_MAPPING_VALUE = "?"


# ---------------------------------------------------------------------------
# Worksheet access
# ---------------------------------------------------------------------------

def open_account_worksheet(
    client: gspread.Client,
    sheet_id: str,
    worksheet_name: str,
) -> gspread.Worksheet:
    """Open the given account's worksheet/tab.

    Raises:
        gspread.exceptions.WorksheetNotFound: If the worksheet does not exist.
            This script only updates an existing sheet; it never creates one
            (upload_to_sheets.py is responsible for creating account tabs).
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
        account_number = _get_cell(row, header_row, "Account Number")

        # Try the confident, generalizable business rules first (internal
        # transfer between our own tracked accounts, or an incoming
        # customer payment). Falls back to the existing get_head()
        # heuristic — with business_unit/type_rera_idw/tcp_head left as
        # "?" — for anything those two rules don't confidently cover.
        resolved = resolve_business_fields(account_number, description, deposits, withdrawals)
        head = resolved["head"] or get_head(description, deposits, withdrawals)
        narration = generate_narration(description, head, amount)

        row_values = {
            BUSINESS_UNIT_COLUMN: resolved["business_unit"],
            HEAD_COLUMN: head,
            TYPE_RERA_IDW_COLUMN: resolved["type_rera_idw"],
            TCP_HEAD_COLUMN: resolved["tcp_head"],
            NARRATION_COLUMN: narration,
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
    worksheet_name: str,
    sheet_id: str = MASTER_SHEET_ID,
) -> int:
    """Classify all unclassified transactions in one account's worksheet/tab.

    Args:
        credentials_path: Path to the Google service-account credentials JSON.
        worksheet_name: The account's worksheet/tab name (e.g. "YES BANK - 2477").
        sheet_id: Spreadsheet ID containing the account tabs.

    Returns:
        Number of rows updated.
    """
    # Credential resolution (file vs GOOGLE_CREDENTIALS_JSON env var fallback)
    # is handled entirely inside get_gspread_client() — no upfront existence
    # check here, since that would bypass the env var fallback it supports.
    client = get_gspread_client(credentials_path)
    worksheet = open_account_worksheet(client, sheet_id, worksheet_name)

    header_row, column_indices = ensure_classification_columns(worksheet)

    return classify_rows(worksheet, header_row, column_indices)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify transactions in one account's Google Sheet tab (Head + Narration)."
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
        help="Override the spreadsheet ID.",
    )

    parser.add_argument(
        "--worksheet-name",
        required=True,
        help="The account's worksheet/tab name (e.g. 'YES BANK - 2477').",
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
