"""Fills blank QTR and MONTH columns in all account worksheets
by deriving values from the TXN DATE column.

QTR uses Indian Financial Year (April–March):
  Q1 = Apr May Jun | Q2 = Jul Aug Sep | Q3 = Oct Nov Dec | Q4 = Jan Feb Mar

MONTH = numeric month number (1–12).

Only updates rows where QTR or MONTH is currently blank.
Never overwrites a cell that already has a value.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __future__ import annotations

from datetime import datetime

import gspread
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client, get_account_worksheets

DATE_FORMATS = ["%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]

MONTH_TO_QTR = {
    4: "1", 5: "1", 6: "1",
    7: "2", 8: "2", 9: "2",
    10: "3", 11: "3", 12: "3",
    1: "4", 2: "4", 3: "4",
}


def parse_date(raw: str) -> datetime | None:
    raw = raw.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    for worksheet in get_account_worksheets(spreadsheet):
        all_values = worksheet.get_all_values()
        if not all_values:
            continue

        header = all_values[0]

        if "TXN DATE" not in header or "QTR" not in header or "MONTH" not in header:
            print(f"[SKIP] {worksheet.title}: missing required columns.")
            continue

        txn_date_col = header.index("TXN DATE")
        qtr_col      = header.index("QTR")
        month_col    = header.index("MONTH")

        updates: list[gspread.cell.Cell] = []
        filled = 0
        skipped = 0

        for offset, row in enumerate(all_values[1:], 2):
            # Pad short rows
            while len(row) <= max(txn_date_col, qtr_col, month_col):
                row.append("")

            txn_date_raw = row[txn_date_col].strip()
            current_qtr  = row[qtr_col].strip()
            current_month = row[month_col].strip()

            if not txn_date_raw:
                continue

            # Skip rows where both are already in the correct format
            qtr_ok = current_qtr and not current_qtr.startswith("Q")
            if qtr_ok and current_month:
                skipped += 1
                continue

            dt = parse_date(txn_date_raw)
            if dt is None:
                print(f"  [WARN] {worksheet.title} row {offset}: cannot parse date '{txn_date_raw}'")
                continue

            month_num = dt.month
            qtr       = MONTH_TO_QTR[month_num]

            if not qtr_ok:
                updates.append(gspread.cell.Cell(row=offset, col=qtr_col + 1, value=qtr))
            if not current_month:
                updates.append(gspread.cell.Cell(row=offset, col=month_col + 1, value=month_num))

            filled += 1

        if updates:
            worksheet.update_cells(updates, value_input_option="RAW")

        print(f"[OK] {worksheet.title}: filled {filled} row(s), skipped {skipped} already-filled row(s).")


if __name__ == "__main__":
    main()
