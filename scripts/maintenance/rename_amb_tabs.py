"""One-off: rename AMB sheet tabs to match the accounts department's own
naming convention (from their duplicate copy of this workbook), instead of
our default '<Bank Name> - <last 4 digits>' pattern.

Renaming a worksheet only changes its title — all rows, formulas, and
formatting inside are untouched.

Usage:
    py -3 scripts/maintenance/rename_amb_tabs.py            (dry run, default)
    py -3 scripts/maintenance/rename_amb_tabs.py --write    (renames the tabs)
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS

AMB_SHEET_ID = "1kVMuah99dU8g3q9zsHxiTtBh7l7E8xIPrFoSmGQpywc"

RENAMES = {
    "Karur Vysya Bank - 6375": "AMB MASTER KVB-6375",
    "Bank of Maharashtra - 2667": "AMB RERA BOM 667",
    "Karur Vysya Bank - 6535": "IDW KVB-6535",
    "Karur Vysya Bank - 1050": "KVB FREE 1050",
    "Axis Bank - 2280": "AMB CA AXIS-280",
}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()

    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(AMB_SHEET_ID)
    existing_titles = {ws.title: ws for ws in ss.worksheets()}

    print("=== Planned renames ===")
    plan = []
    for old_name, new_name in RENAMES.items():
        if old_name not in existing_titles:
            print(f"SKIP (not found): {old_name!r}")
            continue
        row_count = existing_titles[old_name].row_count
        print(f"{old_name!r} -> {new_name!r} (sheet reports {row_count} rows)")
        plan.append((old_name, new_name))

    if not args.write:
        print("\nDry run only - no changes written. Re-run with --write to apply.")
        return

    for old_name, new_name in plan:
        ws = existing_titles[old_name]
        before_values = ws.get_all_values()
        ws.update_title(new_name)
        after_values = ws.get_all_values()
        assert len(before_values) == len(after_values), f"Row count changed for {old_name}!"
        print(f"Renamed {old_name!r} -> {new_name!r} (row count unchanged: {len(after_values)})")

    print("\nFinal tab list:")
    for ws in ss.worksheets():
        print(f"  {ws.title!r}")


if __name__ == "__main__":
    main()
