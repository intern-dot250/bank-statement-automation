"""One-off: re-classify only the AMB rows whose HEAD/BUSINESS UNIT/TYPE FOR
RERA IDW/TCP Head would come out differently under the current
classify_transactions.py rules than what's already written in the sheet
(from before this session's AMB rule fixes).

Never touches a row whose current values already match what the rules
would produce now. Clears just the 5 classification columns (Business
Unit, Head, Type for RERA IDW, TCP Head, Narration) for the changed rows,
then calls the real classify_rows() so Narration/Reason/Confidence and the
red "unverified" coloring all get regenerated consistently through the
normal pipeline — this script never writes classification values itself.

Usage:
    py -3 scripts/maintenance/reclassify_amb_changed_rows.py            (dry run, default)
    py -3 scripts/maintenance/reclassify_amb_changed_rows.py --write    (clears + reclassifies changed rows)
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import gspread

from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS
from classify_transactions import (
    resolve_business_fields,
    classify_rows,
    ensure_classification_columns,
    CLASSIFICATION_COLUMNS,
)

AMB_SHEET_ID = "1kVMuah99dU8g3q9zsHxiTtBh7l7E8xIPrFoSmGQpywc"
TABS = ["IDW KVB-6535", "KVB FREE 1050"]


def find_changed_rows(spreadsheet: gspread.Spreadsheet, tab: str) -> list[dict]:
    ws = spreadsheet.worksheet(tab)
    all_values = ws.get_all_values()
    header = all_values[0]

    idx = {c: header.index(c) for c in (
        "DESCRIPTION", "DEBITS", "CREDITS", "Account Number",
        "BUSINESS UNIT", "HEAD", "TYPE FOR RERA IDW", "TCP Head",
    ) if c in header}

    changed = []
    for row_num, row in enumerate(all_values[1:], start=2):
        if len(row) <= idx["DESCRIPTION"] or not row[idx["DESCRIPTION"]].strip():
            continue
        description = row[idx["DESCRIPTION"]]
        account_number = row[idx["Account Number"]] if len(row) > idx["Account Number"] else ""
        if not account_number:
            continue

        def cell(col):
            return row[idx[col]] if len(row) > idx[col] else ""

        deposits = float((cell("CREDITS") or "0").replace(",", "") or 0)
        withdrawals = float((cell("DEBITS") or "0").replace(",", "") or 0)

        resolved = resolve_business_fields(account_number, description, deposits, withdrawals, spreadsheet=spreadsheet)

        current = {
            "BUSINESS UNIT": cell("BUSINESS UNIT"),
            "HEAD": cell("HEAD"),
            "TYPE FOR RERA IDW": cell("TYPE FOR RERA IDW"),
            "TCP Head": cell("TCP Head"),
        }
        new = {
            "BUSINESS UNIT": resolved["business_unit"],
            "HEAD": resolved["head"] or current["HEAD"],  # None means "fall back to heads.py heuristic" — can't predict that here, so treat as unchanged unless BU/Type/TCP differ
            "TYPE FOR RERA IDW": resolved["type_rera_idw"],
            "TCP Head": resolved["tcp_head"],
        }

        if current != new:
            changed.append({
                "row_num": row_num,
                "description": description,
                "current": current,
                "new": new,
            })

    return changed


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()

    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(AMB_SHEET_ID)

    for tab in TABS:
        changed = find_changed_rows(spreadsheet, tab)
        print(f"=== {tab}: {len(changed)} rows would change ===")
        for c in changed:
            print(f"  row {c['row_num']}: {c['description'][:70]!r}")
            for field in ("BUSINESS UNIT", "HEAD", "TYPE FOR RERA IDW", "TCP Head"):
                if c["current"][field] != c["new"][field]:
                    print(f"    {field}: {c['current'][field]!r} -> {c['new'][field]!r}")
        print()

        if not args.write or not changed:
            continue

        ws = spreadsheet.worksheet(tab)
        header_row, column_indices = ensure_classification_columns(ws)
        clear_cells = []
        for c in changed:
            for col in ("BUSINESS UNIT", "HEAD", "TYPE FOR RERA IDW", "TCP Head", "NARRATION"):
                if col in column_indices:
                    clear_cells.append(gspread.cell.Cell(row=c["row_num"], col=column_indices[col], value=""))
        ws.update_cells(clear_cells, value_input_option="RAW")
        print(f"Cleared {len(changed)} rows' classification columns in {tab}; running classify_rows()...")
        updated = classify_rows(ws, header_row, column_indices)
        print(f"classify_rows() updated {updated} row(s) in {tab}.")

    if not args.write:
        print("Dry run only - no changes written. Re-run with --write to apply.")


if __name__ == "__main__":
    main()
