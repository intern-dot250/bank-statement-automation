"""One-off: refresh the "Check" column's formula on every existing account
worksheet tab, in every company's spreadsheet, to wrap it in ROUND(...,2).

Balance (AI) - BALANCE can leave a sub-paise floating-point residue (e.g.
1.89e-9) even when the two values are truly equal, since BALANCE is a
chain of previous+credit-debit additions - see upload_to_sheets.py's
append_unique_rows() docstring. Rounding to paise stops a real match from
showing as a spurious non-zero Check value.

Safe to re-run: only a cell whose formula still holds exactly the old,
un-rounded pattern is rewritten - anything else (already rounded, or
edited/replaced by a human) is left untouched.
"""

import re
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import gspread

from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
    get_account_worksheets,
    get_spreadsheet_id_for_company,
)
import company_sheets_store

_OLD_FORMULA_RE = re.compile(
    r'^=IFERROR\(([A-Z]+)(\d+)-([A-Z]+)\2,""\)$'
)


def fix_worksheet(worksheet: gspread.Worksheet) -> int:
    """Returns the number of Check cells rewritten in this tab."""
    header = worksheet.row_values(1)
    if "Check" not in header or "Balance (AI)" not in header:
        return 0

    check_col_idx = header.index("Check") + 1
    check_col_letter = gspread.utils.rowcol_to_a1(1, check_col_idx).rstrip("0123456789")
    balance_ai_col_letter = gspread.utils.rowcol_to_a1(1, header.index("Balance (AI)") + 1).rstrip("0123456789")

    all_values = worksheet.get_all_values()
    last_row = len(all_values)
    if last_row < 2:
        return 0

    check_formulas = worksheet.get(
        f"{check_col_letter}2:{check_col_letter}{last_row}",
        value_render_option=gspread.utils.ValueRenderOption.formula,
    )

    updates = []
    for offset, row_values in enumerate(check_formulas):
        formula = row_values[0] if row_values else ""
        match = _OLD_FORMULA_RE.match(formula)
        if not match:
            continue
        balance_ai_col, row, balance_col = match.group(1), match.group(2), match.group(3)
        if balance_ai_col != balance_ai_col_letter:
            continue  # not actually our Balance (AI)/BALANCE pattern - leave alone
        new_formula = f'=IFERROR(ROUND({balance_ai_col}{row}-{balance_col}{row},2),"")'
        updates.append(gspread.cell.Cell(row=int(row), col=check_col_idx, value=new_formula))

    if updates:
        worksheet.update_cells(updates, value_input_option="USER_ENTERED")

    return len(updates)


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)

    spreadsheet_ids = {MASTER_SHEET_ID}
    for row in company_sheets_store.list_company_sheets():
        spreadsheet_ids.add(get_spreadsheet_id_for_company(row.get("company")))

    total = 0
    for spreadsheet_id in spreadsheet_ids:
        spreadsheet = client.open_by_key(spreadsheet_id)
        for worksheet in get_account_worksheets(spreadsheet):
            n = fix_worksheet(worksheet)
            if n:
                print(f"[OK] {spreadsheet.title} / {worksheet.title}: rounded {n} Check formula(s).")
                total += n
            else:
                print(f"[SKIP] {spreadsheet.title} / {worksheet.title}: nothing to fix.")

    print(f"\nDone. {total} total Check formula(s) rounded across all companies.")


if __name__ == "__main__":
    main()
