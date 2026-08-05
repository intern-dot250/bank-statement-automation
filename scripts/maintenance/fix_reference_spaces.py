"""One-off: strip stray internal spaces from the REFERENCE column on every
account worksheet tab.

extract_statement.py used to join every field's bucketed words with a
space, including "reference" - but a bank-generated reference code (e.g.
"YESME6216001852500") is always one contiguous token; the space was an
extraction artifact (pdfplumber sometimes splits one such code into more
than one "word", or a statement wraps it across physical PDF lines).
That's now fixed for future extractions (see _bucket_line() and the
continuation-line merge in extract_transactions_from_pdf()); this script
corrects already-uploaded rows by removing any spaces already sitting in
REFERENCE, so historical rows match the same no-space convention.

Safe to re-run: a value with no spaces is left untouched.
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
        if "REFERENCE" not in header:
            print(f"[SKIP] {worksheet.title}: no REFERENCE column.")
            continue

        ref_col_index = header.index("REFERENCE")
        ref_col_letter = chr(ord("A") + ref_col_index)
        last_row = len(all_values)

        original = [row[ref_col_index] if ref_col_index < len(row) else "" for row in all_values[1:]]
        fixed = [value.replace(" ", "") for value in original]

        if fixed == original:
            print(f"[SKIP] {worksheet.title}: no spaced REFERENCE values.")
            continue

        worksheet.update(
            range_name=f"{ref_col_letter}2:{ref_col_letter}{last_row}",
            values=[[value] for value in fixed],
            value_input_option="RAW",
        )
        changed = sum(1 for o, f in zip(original, fixed) if o != f)
        print(f"[OK] {worksheet.title}: removed spaces from {changed} REFERENCE value(s).")


if __name__ == "__main__":
    main()
