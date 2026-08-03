"""One-off: rewrite the existing QTR column formula on every account
worksheet tab so it displays "Q1".."Q4" (accounts team's own convention)
instead of the bare number 1-4.

QTR is a live formula (see upload_to_sheets.py's append_unique_rows()),
not a static value, so this only needs to overwrite the formula string
for every existing data row - it recalculates from each row's own MONTH
cell (column C) immediately, no need to re-derive anything from TXN DATE.

Safe to re-run: every row gets the same corrected formula regardless of
its current value.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client, get_account_worksheets


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    for worksheet in get_account_worksheets(spreadsheet):
        all_values = worksheet.get_all_values()
        if len(all_values) < 2:
            print(f"[SKIP] {worksheet.title}: no data rows.")
            continue

        header = all_values[0]
        if "QTR" not in header or "MONTH" not in header:
            print(f"[SKIP] {worksheet.title}: missing QTR/MONTH column.")
            continue

        qtr_col_letter = chr(ord("A") + header.index("QTR"))
        month_col_letter = chr(ord("A") + header.index("MONTH"))
        last_row = len(all_values)

        formulas = [
            [f'=IFERROR("Q"&(INT(MOD({month_col_letter}{r}-4,12)/3)+1),"")']
            for r in range(2, last_row + 1)
        ]

        worksheet.update(
            range_name=f"{qtr_col_letter}2:{qtr_col_letter}{last_row}",
            values=formulas,
            value_input_option="USER_ENTERED",
        )
        print(f"[OK] {worksheet.title}: rewrote QTR formula for {len(formulas)} row(s).")


if __name__ == "__main__":
    main()
