"""One-off: rename "RERA IDW New" -> "RERA 2 IDW" in the TYPE FOR RERA IDW
column on every account worksheet tab.

The accounts team renamed this label on 2026-08-07 (see
classify_transactions.py's _AMBIGUOUS_STAGE_PAIRS handling, which now
writes "RERA 2 IDW" for all future uploads). This script updates
already-uploaded rows so historical data matches the new label too.

Safe to re-run: a cell that doesn't say "RERA IDW New" is left untouched.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import gspread

from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client, get_account_worksheets

OLD_LABEL = "RERA IDW New"
NEW_LABEL = "RERA 2 IDW"


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    total = 0
    for worksheet in get_account_worksheets(spreadsheet):
        all_values = worksheet.get_all_values()
        if len(all_values) < 2:
            print(f"[SKIP] {worksheet.title}: no data rows.")
            continue

        header = all_values[0]
        if "TYPE FOR RERA IDW" not in header:
            print(f"[SKIP] {worksheet.title}: no TYPE FOR RERA IDW column.")
            continue

        col_idx = header.index("TYPE FOR RERA IDW") + 1  # 1-based

        updates = []
        for offset, row in enumerate(all_values[1:]):
            sheet_row = offset + 2
            value = row[col_idx - 1].strip() if len(row) >= col_idx else ""
            if value == OLD_LABEL:
                updates.append(gspread.cell.Cell(row=sheet_row, col=col_idx, value=NEW_LABEL))

        if not updates:
            print(f"[SKIP] {worksheet.title}: no '{OLD_LABEL}' values.")
            continue

        worksheet.update_cells(updates, value_input_option="RAW")
        print(f"[OK] {worksheet.title}: renamed {len(updates)} cell(s) to '{NEW_LABEL}'.")
        total += len(updates)

    print(f"\nDone. {total} total cell(s) renamed across all accounts.")


if __name__ == "__main__":
    main()
