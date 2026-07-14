"""Backfill QTR and MONTH columns on existing rows that have TXN DATE."""

import pandas as pd
import gspread
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client, get_account_worksheets


def get_qtr(month: int) -> int:
    if month in (4, 5, 6): return 1
    if month in (7, 8, 9): return 2
    if month in (10, 11, 12): return 3
    return 4  # Jan, Feb, Mar


client = get_gspread_client(DEFAULT_CREDENTIALS)
spreadsheet = client.open_by_key(MASTER_SHEET_ID)

for ws in get_account_worksheets(spreadsheet):
    all_values = ws.get_all_values()
    if not all_values:
        continue
    hdr = all_values[0]
    if "TXN DATE" not in hdr or "QTR" not in hdr or "MONTH" not in hdr:
        print(f"  {ws.title}: missing columns, skipped")
        continue

    date_col = hdr.index("TXN DATE")
    qtr_col = hdr.index("QTR") + 1
    month_col = hdr.index("MONTH") + 1

    updates = []
    for offset, row in enumerate(all_values[1:]):
        sheet_row = offset + 2
        if not row or not row[date_col].strip():
            continue
        # Skip if already filled
        qtr_val = row[hdr.index("QTR")].strip() if len(row) > hdr.index("QTR") else ""
        if qtr_val:
            continue
        try:
            dt = pd.to_datetime(row[date_col], dayfirst=True)
        except Exception:
            continue
        m = dt.month
        updates.append(gspread.cell.Cell(row=sheet_row, col=qtr_col, value=get_qtr(m)))
        updates.append(gspread.cell.Cell(row=sheet_row, col=month_col, value=m))

    if updates:
        ws.update_cells(updates, value_input_option="RAW")
        print(f"  {ws.title}: {len(updates)//2} rows updated")
    else:
        print(f"  {ws.title}: already filled, skipped")

print("Done.")
