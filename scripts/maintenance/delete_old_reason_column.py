"""Delete the old 'REASON FOR ?' column from every account tab."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client, get_account_worksheets

client = get_gspread_client(DEFAULT_CREDENTIALS)
spreadsheet = client.open_by_key(MASTER_SHEET_ID)

for ws in get_account_worksheets(spreadsheet):
    hdr = ws.row_values(1)
    if "REASON FOR ?" in hdr:
        col = hdr.index("REASON FOR ?") + 1
        ws.delete_columns(col)
        print(f"  {ws.title}: deleted REASON FOR ? (was col {col})")
    else:
        print(f"  {ws.title}: not found, skipped")

print("Done.")
