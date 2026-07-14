"""One-off: regenerate the Narration column for ALL already-classified rows
in every account worksheet, using the new accounts-team-format narration
logic. classify_rows() in classify_transactions.py deliberately skips rows
that are already fully classified (idempotent for normal runs), so this
script bypasses that check to backfill Narration on existing data without
touching Business Unit/Head/Type for RERA IDW/TCP Head (those keep
whatever value is already in the sheet, including any manual corrections
the accounts team has already made).
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __future__ import annotations

from pathlib import Path

from classify_transactions import (
    CLASSIFICATION_COLUMNS,
    NARRATION_COLUMN,
    _get_cell,
    _is_row_empty,
    _to_float,
    _parse_amount,
)
from narration import generate_narration
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client
import gspread

SCRIPT_DIR = Path(__file__).resolve().parent


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    skip_titles = {"Summary", "Final Report", "Validation"}
    worksheets = [ws for ws in spreadsheet.worksheets() if ws.title not in skip_titles]

    for worksheet in worksheets:
        all_values = worksheet.get_all_values()
        if not all_values:
            continue

        header_row = all_values[0]
        if NARRATION_COLUMN not in header_row:
            print(f"[SKIP] {worksheet.title}: no Narration column.")
            continue

        narration_col = header_row.index(NARRATION_COLUMN) + 1
        data_rows = all_values[1:]
        updates: list[gspread.cell.Cell] = []

        for offset, row in enumerate(data_rows):
            sheet_row_number = offset + 2

            if _is_row_empty(row):
                continue

            description = _get_cell(row, header_row, "DESCRIPTION")
            if not description:
                continue

            head = _get_cell(row, header_row, "HEAD")
            if not head:
                continue  # not yet classified at all — leave for classify_transactions.py

            business_unit = _get_cell(row, header_row, "BUSINESS UNIT") or "?"
            type_rera_idw = _get_cell(row, header_row, "TYPE FOR RERA IDW") or "?"
            deposits_raw = _get_cell(row, header_row, "CREDITS")
            withdrawals_raw = _get_cell(row, header_row, "DEBITS")
            deposits = _to_float(deposits_raw)
            withdrawals = _to_float(withdrawals_raw)
            amount = _parse_amount(deposits_raw, withdrawals_raw)
            account_number = _get_cell(row, header_row, "Account Number")
            reference_value = _get_cell(row, header_row, "Cheque No/Ref") or None

            new_narration = generate_narration(
                description,
                head,
                amount,
                business_unit=business_unit,
                type_rera_idw=type_rera_idw,
                deposits=deposits,
                withdrawals=withdrawals,
                own_account_number=account_number,
                reference=reference_value,
            )

            old_narration = _get_cell(row, header_row, NARRATION_COLUMN)
            if new_narration != old_narration:
                updates.append(
                    gspread.cell.Cell(row=sheet_row_number, col=narration_col, value=new_narration)
                )

        if updates:
            worksheet.update_cells(updates, value_input_option="RAW")
            print(f"[OK] {worksheet.title}: updated {len(updates)} narration(s).")
        else:
            print(f"[OK] {worksheet.title}: nothing to update.")


if __name__ == "__main__":
    main()
