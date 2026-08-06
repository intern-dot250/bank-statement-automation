"""One-off: insert "Balance (AI)" and "Check" columns immediately after
NARRATION on every existing account worksheet tab, matching the position
upload_to_sheets.py's EXPECTED_COLUMNS now declares for all future uploads.

Without this migration, the next real upload to an already-existing tab
would let upload_to_sheets.ensure_header_row() silently "auto-heal" the
missing columns by tacking them onto the very END of the header instead
(after REASON/APPROVAL 1-3) - misaligning that tab from both the intended
layout and from brand-new tabs.

Each existing row's already-correct BALANCE value is copied into the new
Balance (AI) column - BALANCE itself is not modified. This is required,
not cosmetic: upload_to_sheets.py's duplicate-detection key now uses
"Balance (AI)" instead of "BALANCE" (BALANCE is a live, accounts-team-
editable formula going forward, no longer a stable anchor). Leaving
Balance (AI) blank on already-uploaded rows would make re-uploading an
already-processed statement fail to be recognised as a duplicate.

No formula is retrofitted onto old rows' BALANCE/Check - only rows
appended after this migration get the new live-formula treatment, per
"only new uploads follow the new behaviour."

Safe to re-run: any tab that already has a "Balance (AI)" column is
skipped.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import gspread

from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
    get_account_worksheets,
    apply_numeric_format,
)


def migrate_worksheet(worksheet: gspread.Worksheet) -> bool:
    """Returns True if this tab was migrated, False if skipped."""
    header = worksheet.row_values(1)
    if not header:
        print(f"[SKIP] {worksheet.title}: empty tab (no header).")
        return False
    if "Balance (AI)" in header:
        print(f"[SKIP] {worksheet.title}: already migrated.")
        return False
    if "NARRATION" not in header or "BALANCE" not in header:
        print(f"[SKIP] {worksheet.title}: missing NARRATION/BALANCE column.")
        return False

    narration_idx = header.index("NARRATION")  # 0-based
    balance_idx = header.index("BALANCE")  # 0-based
    insert_at_col = narration_idx + 2  # 1-based column to insert before

    # Unformatted so a comma-grouped display string (e.g. "40,312") isn't
    # copied as literal text - we want the real number, same as
    # load_existing_data()'s own dedup-safe read elsewhere in this project.
    all_values = worksheet.get_all_values(
        value_render_option=gspread.utils.ValueRenderOption.unformatted
    )
    data_rows = all_values[1:]

    balance_ai_column = ["Balance (AI)"] + [
        row[balance_idx] if balance_idx < len(row) else "" for row in data_rows
    ]
    check_column = ["Check"] + [""] * len(data_rows)

    worksheet.insert_cols(
        [balance_ai_column, check_column],
        col=insert_at_col,
        value_input_option=gspread.utils.ValueInputOption.raw,
    )
    apply_numeric_format(worksheet)

    print(f"[OK] {worksheet.title}: inserted Balance (AI)/Check after NARRATION, "
          f"backfilled Balance (AI) for {len(data_rows)} existing row(s).")
    return True


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    migrated = 0
    for worksheet in get_account_worksheets(spreadsheet):
        if migrate_worksheet(worksheet):
            migrated += 1

    print(f"\nDone. {migrated} tab(s) migrated.")


if __name__ == "__main__":
    main()
