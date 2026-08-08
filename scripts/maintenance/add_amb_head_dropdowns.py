"""One-off: add a dropdown (data validation) to the Head 1 / Head 2 /
Head 3 columns of the AMB Beneficiary Master, listing AMB's established
Head vocabulary (config/heads_config.json) plus any extra values already
present in the sheet that aren't in that canonical list (so existing data
isn't invalidated by a strict dropdown).

Usage:
    py -3 scripts/maintenance/add_amb_head_dropdowns.py            (dry run, default)
    py -3 scripts/maintenance/add_amb_head_dropdowns.py --write    (writes to the AMB sheet)
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS

AMB_SHEET_ID = "1kVMuah99dU8g3q9zsHxiTtBh7l7E8xIPrFoSmGQpywc"
HEADS_CONFIG_PATH = Path(__file__).resolve().parent.parent.parent / "config" / "heads_config.json"
HEAD_COLUMNS = ["Head 1", "Head 2", "Head 3"]
# Generous headroom so future manually-added rows also get the dropdown.
VALIDATION_LAST_ROW = 1000


def load_canonical_heads() -> list[str]:
    with open(HEADS_CONFIG_PATH, encoding="utf-8") as f:
        data = json.load(f)
    return list(data["heads"].keys())


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()

    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(AMB_SHEET_ID)
    ws = ss.worksheet("Beneficiary Master")
    all_values = ws.get_all_values()
    header = all_values[0]
    data_rows = all_values[1:]

    canonical = load_canonical_heads()
    head_idxs = [header.index(col) for col in HEAD_COLUMNS]

    existing_values: set[str] = set()
    for row in data_rows:
        for idx in head_idxs:
            if len(row) > idx and row[idx].strip():
                existing_values.add(row[idx].strip())

    extras = sorted(v for v in existing_values if v not in canonical)
    options = canonical + extras

    print(f"=== {len(canonical)} canonical heads + {len(extras)} extra existing values = {len(options)} dropdown options ===")
    if extras:
        print(f"Extra values found in sheet not in heads_config.json: {extras}")
    print(f"=== Dropdown will be applied to columns {HEAD_COLUMNS}, rows 2-{VALIDATION_LAST_ROW} ===")

    if not args.write:
        print("\nDry run only - no changes written. Re-run with --write to apply.")
        return

    requests = []
    for idx in head_idxs:
        requests.append({
            "setDataValidation": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": 1,  # row 2 (0-indexed, skip header)
                    "endRowIndex": VALIDATION_LAST_ROW,
                    "startColumnIndex": idx,
                    "endColumnIndex": idx + 1,
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": v} for v in options],
                    },
                    "showCustomUi": True,
                    "strict": True,
                },
            }
        })

    ss.batch_update({"requests": requests})
    print(f"Added dropdown to {', '.join(HEAD_COLUMNS)}.")


if __name__ == "__main__":
    main()
