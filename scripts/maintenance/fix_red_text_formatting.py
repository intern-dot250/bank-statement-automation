"""One-off: fix the unverified-row red-text formatting on the live sheet.

The previous version of _mark_rows_unverified() colored a min..max column
span (Business Unit through Narration), which swept in every blank column
in between (SUB HEAD, RECO, CONCERN, CUST ID, APT#, ACC REMARKS, CRM
REMARKS) as well as leaving stray red formatting on columns from earlier
layout iterations. This script, for every account worksheet:

  1. Resets ALL columns' text color to default (black) for every data row.
  2. Re-applies red text ONLY to Business Unit, Head, Type for RERA IDW,
     TCP Head, and Narration, and ONLY on rows that are actually fully
     classified (i.e. have a Head value) — matching the corrected logic
     in classify_transactions.py.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from __future__ import annotations

from classify_transactions import (
    BUSINESS_UNIT_COLUMN,
    HEAD_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
    NARRATION_COLUMN,
    UNVERIFIED_TEXT_COLOR,
    _get_cell,
    _is_row_empty,
)
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client

DEFAULT_TEXT_COLOR = {"red": 0.0, "green": 0.0, "blue": 0.0}

TARGET_COLUMNS = [
    BUSINESS_UNIT_COLUMN, HEAD_COLUMN, TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN, NARRATION_COLUMN,
]


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    skip_titles = {"Summary", "Final Report", "Validation"}
    worksheets = [ws for ws in spreadsheet.worksheets() if ws.title not in skip_titles]

    for worksheet in worksheets:
        all_values = worksheet.get_all_values()
        if not all_values:
            continue

        header_row = all_values[0]
        data_rows = all_values[1:]
        num_rows = len(data_rows)
        num_cols = len(header_row)

        # 1. Reset every column to default black text for all data rows.
        reset_request = {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 1,  # row 2 (0-based, inclusive) — skip header
                    "endRowIndex": 1 + num_rows,
                    "startColumnIndex": 0,
                    "endColumnIndex": num_cols,
                },
                "cell": {
                    "userEnteredFormat": {"textFormat": {"foregroundColor": DEFAULT_TEXT_COLOR}}
                },
                "fields": "userEnteredFormat.textFormat.foregroundColor",
            }
        }

        # 2. Re-apply red ONLY to the 5 target columns, ONLY on rows that
        # are actually classified (have a Head value).
        classified_rows = []
        for offset, row in enumerate(data_rows):
            if _is_row_empty(row):
                continue
            if _get_cell(row, header_row, HEAD_COLUMN):
                classified_rows.append(offset + 2)  # 1-based sheet row number

        red_requests = [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": header_row.index(column_name),
                        "endColumnIndex": header_row.index(column_name) + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {"textFormat": {"foregroundColor": UNVERIFIED_TEXT_COLOR}}
                    },
                    "fields": "userEnteredFormat.textFormat.foregroundColor",
                }
            }
            for row in classified_rows
            for column_name in TARGET_COLUMNS
            if column_name in header_row
        ]

        spreadsheet.batch_update({"requests": [reset_request]})
        if red_requests:
            spreadsheet.batch_update({"requests": red_requests})

        print(f"[OK] {worksheet.title}: reset {num_rows} row(s), re-colored {len(classified_rows)} classified row(s).")


if __name__ == "__main__":
    main()
