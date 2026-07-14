"""Fix old 'Rera 2 IDW' and 'Rera to IDW' values in the sheet.

Patches:
  TYPE FOR RERA IDW: "Rera 2 IDW" -> "RERA IDW New"
  TCP Head:          "Rera to IDW" -> "Internal transfer"

Run once:
    py -3 backfill_rera_idw.py
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import gspread
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client, get_account_worksheets

client = get_gspread_client(DEFAULT_CREDENTIALS)
spreadsheet = client.open_by_key(MASTER_SHEET_ID)

total = 0

for ws in get_account_worksheets(spreadsheet):
    all_values = ws.get_all_values()
    if not all_values:
        continue
    hdr = all_values[0]

    type_col_idx = hdr.index("TYPE FOR RERA IDW") + 1 if "TYPE FOR RERA IDW" in hdr else None
    tcp_col_idx = hdr.index("TCP Head") + 1 if "TCP Head" in hdr else None

    if not type_col_idx and not tcp_col_idx:
        continue

    updates = []
    for offset, row in enumerate(all_values[1:]):
        sheet_row = offset + 2

        if type_col_idx:
            type_val = row[type_col_idx - 1].strip() if len(row) >= type_col_idx else ""
            if type_val == "Rera 2 IDW":
                updates.append(gspread.cell.Cell(row=sheet_row, col=type_col_idx, value="RERA IDW New"))

        if tcp_col_idx:
            tcp_val = row[tcp_col_idx - 1].strip() if len(row) >= tcp_col_idx else ""
            if tcp_val == "Rera to IDW":
                updates.append(gspread.cell.Cell(row=sheet_row, col=tcp_col_idx, value="Internal transfer"))

    if updates:
        ws.update_cells(updates, value_input_option="RAW")
        changed = len(updates)
        print(f"  {ws.title}: {changed} cell(s) updated")
        total += changed
    else:
        print(f"  {ws.title}: nothing to fix")

print(f"\nDone. {total} total cells patched.")
