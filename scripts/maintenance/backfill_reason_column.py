"""One-off: add the REASON FOR ? column (if missing) and populate it for
every row across all account tabs — including rows that are already
fully classified and would normally be skipped by classify_rows()'s
idempotency check, since that check deliberately doesn't require this
new column to be filled in.

For a row with no "?" anywhere, the Reason column is written blank.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from __future__ import annotations

from classify_transactions import (
    BUSINESS_UNIT_COLUMN,
    HEAD_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
    REASON_COLUMN,
    UNKNOWN_MAPPING_VALUE,
    resolve_business_fields,
    _build_reason_text,
    _get_cell,
    _is_row_empty,
    _to_float,
)
from heads import get_head
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client, get_account_worksheets
import gspread


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    for worksheet in get_account_worksheets(spreadsheet):
        header_row = worksheet.row_values(1)
        if not header_row:
            continue

        if REASON_COLUMN not in header_row:
            header_row.append(REASON_COLUMN)
            worksheet.update(range_name="A1", values=[header_row])
            print(f"[HEADER] {worksheet.title}: added {REASON_COLUMN} column.")

        reason_col_index = header_row.index(REASON_COLUMN) + 1

        all_values = worksheet.get_all_values()
        data_rows = all_values[1:]
        updates: list[gspread.cell.Cell] = []
        with_reason = 0
        blank = 0

        for offset, row in enumerate(data_rows):
            sheet_row_number = offset + 2

            if _is_row_empty(row):
                continue

            description = _get_cell(row, header_row, "DESCRIPTION")
            if not description:
                continue

            current_head = _get_cell(row, header_row, HEAD_COLUMN)
            if not current_head:
                # Not yet classified at all — leave for classify_transactions.py.
                continue

            deposits_raw = _get_cell(row, header_row, "CREDITS")
            withdrawals_raw = _get_cell(row, header_row, "DEBITS")
            deposits = _to_float(deposits_raw)
            withdrawals = _to_float(withdrawals_raw)
            account_number = _get_cell(row, header_row, "Account Number")

            resolved = resolve_business_fields(account_number, description, deposits, withdrawals)

            # Reconstruct the SAME display_head this row was actually
            # classified with, so the reason lines up with what's on the
            # sheet right now (rather than recomputing get_head(), which
            # for a row already saved as "?" would just be re-deriving
            # the same "Others" fallback anyway).
            display_head = current_head

            reason_text = _build_reason_text(display_head, resolved)

            existing_reason = _get_cell(row, header_row, REASON_COLUMN)
            if existing_reason == reason_text:
                continue

            updates.append(
                gspread.cell.Cell(row=sheet_row_number, col=reason_col_index, value=reason_text)
            )
            if reason_text:
                with_reason += 1
            else:
                blank += 1

        if updates:
            worksheet.update_cells(updates, value_input_option="RAW")

        print(f"[OK] {worksheet.title}: {with_reason} row(s) with a reason, {blank} row(s) blank (no ?).")


if __name__ == "__main__":
    main()
