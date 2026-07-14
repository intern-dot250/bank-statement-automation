"""Backfill CONFIDENCE, REASON, and APPROVAL 1/2/3 columns on already-classified rows.

Reads every account tab, finds rows that have HEAD filled but are missing
CONFIDENCE, and re-runs resolve_business_fields() to write the new columns.
Never overwrites APPROVAL columns that already have a value.

Usage:
    py -3 backfill_new_columns.py
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __future__ import annotations

from pathlib import Path

import gspread

from classify_transactions import (
    BUSINESS_UNIT_COLUMN,
    HEAD_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
    NARRATION_COLUMN,
    CONFIDENCE_COLUMN,
    REASON_COLUMN,
    APPROVAL_1_COLUMN,
    APPROVAL_2_COLUMN,
    APPROVAL_3_COLUMN,
    CLASSIFICATION_COLUMNS,
    UNKNOWN_MAPPING_VALUE,
    resolve_business_fields,
    ensure_classification_columns,
    _build_reason_text,
    _get_cell,
    _is_row_empty,
    _to_float,
)
from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
    get_account_worksheets,
)


def backfill_worksheet(worksheet: gspread.Worksheet) -> int:
    """Backfill new columns on already-classified rows. Returns count updated."""
    # Ensure new columns exist in header
    header_row, col_idx = ensure_classification_columns(worksheet)

    all_values = worksheet.get_all_values()
    if len(all_values) < 2:
        return 0

    hdr = all_values[0]
    updates: list[gspread.cell.Cell] = []
    count = 0

    for offset, row in enumerate(all_values[1:]):
        sheet_row = offset + 2

        if _is_row_empty(row):
            continue

        head = _get_cell(row, hdr, HEAD_COLUMN)
        if not head or head == "":
            continue  # not classified yet

        # Skip if CONFIDENCE already filled — already backfilled
        if _get_cell(row, hdr, CONFIDENCE_COLUMN):
            continue

        description = _get_cell(row, hdr, "DESCRIPTION")
        if not description:
            continue

        account_number = _get_cell(row, hdr, "Account Number")
        deposits = _to_float(_get_cell(row, hdr, "CREDITS"))
        withdrawals = _to_float(_get_cell(row, hdr, "DEBITS"))

        resolved = resolve_business_fields(account_number, description, deposits, withdrawals)

        # Use existing head from sheet (already verified), not re-derived head
        display_head = head
        reason_text = _build_reason_text(display_head, resolved)
        confidence = resolved.get("confidence", "Low")

        new_values = {
            CONFIDENCE_COLUMN: confidence,
            REASON_COLUMN: reason_text,
        }

        # Only write approval columns if blank
        for col in (APPROVAL_1_COLUMN, APPROVAL_2_COLUMN, APPROVAL_3_COLUMN):
            if not _get_cell(row, hdr, col):
                new_values[col] = ""

        for col_name, value in new_values.items():
            if col_name in col_idx:
                updates.append(
                    gspread.cell.Cell(row=sheet_row, col=col_idx[col_name], value=value)
                )

        count += 1

    if updates:
        worksheet.update_cells(updates, value_input_option="RAW")

    return count


def main() -> None:
    print("=" * 60)
    print("Backfilling CONFIDENCE / REASON / APPROVAL columns")
    print("=" * 60)
    print()

    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    total = 0
    for ws in get_account_worksheets(spreadsheet):
        print(f"  Processing: {ws.title} ...", end=" ", flush=True)
        n = backfill_worksheet(ws)
        print(f"{n} rows updated.")
        total += n

    print()
    print(f"Done. {total} rows backfilled across all accounts.")
    print("=" * 60)


if __name__ == "__main__":
    main()
