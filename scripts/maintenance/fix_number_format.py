"""Converts DEBITS, CREDITS, and BALANCE columns from text format
(e.g. "1,03,415" or "2,970") to actual numeric values across all
account worksheets so copy-paste into other sheets works without errors.

Skips cells that are already numeric or empty.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from __future__ import annotations

import gspread
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client, get_account_worksheets

TARGET_COLUMNS = ["DEBITS", "CREDITS", "BALANCE", "SL#", "QTR", "MONTH"]


def to_number(raw: str) -> float | None:
    """Strip Indian-format commas and convert to float. Returns None if blank or unparseable."""
    cleaned = raw.strip().replace(",", "")
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    for worksheet in get_account_worksheets(spreadsheet):
        all_values = worksheet.get_all_values()
        if not all_values:
            continue

        header = all_values[0]
        missing = [c for c in TARGET_COLUMNS if c not in header]
        if missing:
            print(f"[SKIP] {worksheet.title}: missing columns {missing}")
            continue

        col_indices = {c: header.index(c) for c in TARGET_COLUMNS}
        updates: list[gspread.cell.Cell] = []
        converted = 0

        for offset, row in enumerate(all_values[1:], 2):
            for col_name, col_idx in col_indices.items():
                if col_idx >= len(row):
                    continue
                raw = row[col_idx]
                if not raw.strip():
                    continue
                num = to_number(raw)
                if num is not None:
                    value = int(num) if num == int(num) else num
                    updates.append(
                        gspread.cell.Cell(row=offset, col=col_idx + 1, value=value)
                    )
                    converted += 1

        if updates:
            worksheet.update_cells(updates, value_input_option="USER_ENTERED")

        print(f"[OK] {worksheet.title}: converted {converted} cell(s) to number format.")


if __name__ == "__main__":
    main()
