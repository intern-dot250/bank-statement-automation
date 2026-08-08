"""One-off: mark every existing AMB Beneficiary Master row as "Confirmed",
and add a Confirmed/Pending dropdown (data validation) on the STATUS
column for future rows.

Usage:
    py -3 scripts/maintenance/set_amb_beneficiary_status.py            (dry run, default)
    py -3 scripts/maintenance/set_amb_beneficiary_status.py --write    (writes to the AMB sheet)
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import gspread

from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS

AMB_SHEET_ID = "1kVMuah99dU8g3q9zsHxiTtBh7l7E8xIPrFoSmGQpywc"
STATUS_OPTIONS = ["Confirmed", "Pending"]
# Generous headroom so future manually-added rows also get the dropdown.
VALIDATION_LAST_ROW = 1000


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()

    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(AMB_SHEET_ID)
    ws = ss.worksheet("Beneficiary Master")
    all_values = ws.get_all_values()
    header = all_values[0]
    status_idx = header.index("STATUS")
    data_rows = all_values[1:]

    to_set = [i + 2 for i, row in enumerate(data_rows) if len(row) <= status_idx or not row[status_idx].strip()]
    print(f"=== {len(to_set)} of {len(data_rows)} rows will be set to 'Confirmed' ===")
    print(f"=== Dropdown (Confirmed/Pending) will be applied to STATUS column, rows 2-{VALIDATION_LAST_ROW} ===")

    if not args.write:
        print("\nDry run only - no changes written. Re-run with --write to apply.")
        return

    if to_set:
        cells = [gspread.cell.Cell(row=r, col=status_idx + 1, value="Confirmed") for r in to_set]
        ws.update_cells(cells, value_input_option="RAW")
        print(f"\nSet STATUS='Confirmed' on {len(to_set)} rows.")

    ss.batch_update({
        "requests": [{
            "setDataValidation": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": 1,  # row 2 (0-indexed, skip header)
                    "endRowIndex": VALIDATION_LAST_ROW,
                    "startColumnIndex": status_idx,
                    "endColumnIndex": status_idx + 1,
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": v} for v in STATUS_OPTIONS],
                    },
                    "showCustomUi": True,
                    "strict": True,
                },
            }
        }]
    })
    print("Added Confirmed/Pending dropdown to STATUS column.")


if __name__ == "__main__":
    main()
