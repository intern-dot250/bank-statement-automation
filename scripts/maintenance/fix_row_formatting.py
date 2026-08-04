"""One-off: apply the border/wrap-text grid formatting (see
apply_row_grid_format() in upload_to_sheets.py) to every existing data
row on every account worksheet tab, so rows uploaded before this
formatting fix was added can also be safely copy-pasted into the
accounts team's sheet without losing their grid/wrap look.

Safe to re-run: every row just gets the same formatting reapplied
regardless of its current state.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
    get_account_worksheets,
    apply_row_grid_format,
)


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    for worksheet in get_account_worksheets(spreadsheet):
        all_values = worksheet.get_all_values()
        if len(all_values) < 2:
            print(f"[SKIP] {worksheet.title}: no data rows.")
            continue

        last_row = len(all_values)
        apply_row_grid_format(worksheet, start_row=2, end_row=last_row)
        print(f"[OK] {worksheet.title}: applied grid/wrap format to {last_row - 1} row(s).")


if __name__ == "__main__":
    main()
