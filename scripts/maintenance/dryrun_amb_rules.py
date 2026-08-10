"""Dry-run the new AMB-specific classification rules
(classify_transactions._resolve_amb_business_fields) against the 391 real
reference transactions already dumped to C:/tmp/amb_rules_rows.json,
comparing against the accounts department's own ground-truth
Head/Business Unit/Type for RERA IDW/TCP Head. Read-only — does not touch
any live Google Sheet.

Usage:
    py -3 scripts/maintenance/dryrun_amb_rules.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from classify_transactions import _resolve_amb_business_fields, _AMB_STAGE_BY_ACCOUNT
from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS

AMB_SHEET_ID = "1kVMuah99dU8g3q9zsHxiTtBh7l7E8xIPrFoSmGQpywc"

ACCOUNT_NUMBER_BY_TAB = {
    "AMB MASTER KVB-6375": "411435000006375",
    "AMB RERA BOM 667": "60073932667",
    "IDW KVB-6535": "411435000006535",
    "KVB FREE 1050": "4114115000001050",
    "AMB CA AXIS-280": "925020010722280",  # not ruled — expect all skipped
}


def main() -> None:
    rows = json.load(open("C:/tmp/amb_rules_rows.json", encoding="utf-8"))

    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(AMB_SHEET_ID)

    for tab, account_number in ACCOUNT_NUMBER_BY_TAB.items():
        sub = [r for r in rows if r["TAB"] == tab and r["HEAD"] not in (None, "x")]
        matched = 0
        mismatched = []
        skipped = 0
        for r in sub:
            deposits = r["CREDITS"] or 0
            withdrawals = r["DEBITS"] or 0
            result = _resolve_amb_business_fields(
                account_number, r["DESCRIPTION"] or "", deposits, withdrawals, spreadsheet=spreadsheet
            )
            if result is None:
                skipped += 1
                continue
            expected = {
                "head": r["HEAD"],
                "business_unit": r["BUSINESS UNIT"],
                "type_rera_idw": r["TYPE FOR RERA IDW"],
                "tcp_head": r["TCP Head"] or "",
            }
            got = {
                "head": result["head"],
                "business_unit": result["business_unit"],
                "type_rera_idw": result["type_rera_idw"],
                "tcp_head": result["tcp_head"] or "",
            }
            if got == expected:
                matched += 1
            else:
                mismatched.append((r["DESCRIPTION"], expected, got))

        print(f"=== {tab} ({account_number}) — {len(sub)} rows ===")
        print(f"  matched={matched}  skipped(no-rule, fell through)={skipped}  mismatched={len(mismatched)}")
        for desc, expected, got in mismatched[:15]:
            print(f"  MISMATCH: {desc!r:.90}")
            print(f"    expected={expected}")
            print(f"    got     ={got}")
        print()


if __name__ == "__main__":
    main()
