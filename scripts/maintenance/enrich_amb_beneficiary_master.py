"""One-off: enrich the AMB Beneficiary Master with Account Number / IFSC
Code / Bank Name, using the accounts team's "AMB Bank Details" reference
sheet, and add new beneficiary rows for parties not already covered.

Usage:
    py -3 scripts/maintenance/enrich_amb_beneficiary_master.py            (dry run, default)
    py -3 scripts/maintenance/enrich_amb_beneficiary_master.py --write    (writes to the AMB sheet)
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import gspread

from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS

AMB_SHEET_ID = "1kVMuah99dU8g3q9zsHxiTtBh7l7E8xIPrFoSmGQpywc"
BANK_DETAILS_SHEET_ID = "1P0jkt2PL69BfbODlTA0RnZ1-th7w44eWAjaDWtkihbg"

BENEFICIARY_MASTER_COLUMNS = [
    "BENEFICIARY NAME", "Head 1", "Head 2", "Head 3", "NOTES", "ADDED BY",
    "DATE ADDED", "STATUS", "ACCOUNT NUMBER", "IFSC CODE", "BANK NAME", "Company",
]

# Loose "Type" -> this project's actual established Head vocabulary
# (config/heads_config.json / the real AMB ledger's own Head values) - not
# an exact mapping, per the user's own instruction, just the closest fit.
_TYPE_TO_HEAD = {
    "vendor": "Vendor",
    "professional": "Professional",
    "customer": "Collection",  # matches heads_config.json's own Collection->party_types=[Customer]
    "cntr- ho": "Contractor",
    "cntr- site": "Contractor",
    "broker": "Commission",
    "legal": "Legal & Proff.",
    # "employee" deliberately omitted - ambiguous between Salary-HO/
    # Salary-Site with no signal in this sheet to disambiguate; left blank.
}


def _fuzzy_key(name: str) -> str:
    text = name.upper()
    text = re.sub(r"\bPRIVATE LIMITED\b", "PVTLTD", text)
    text = re.sub(r"\bPVT\.?\s*LTD\.?\b", "PVTLTD", text)
    text = re.sub(r"\bLIMITED\b", "LTD", text)
    return re.sub(r"[^A-Z0-9]", "", text)


def _names_match(a: str, b: str) -> bool:
    """Exact match, or a very high whole-string similarity ratio (typo-
    level differences only) - deliberately NOT prefix matching. This
    script matches two INDEPENDENT name sources (a flat reference list vs
    the existing Beneficiary Master), unlike the earlier extraction
    phase's fuzzy merge, which only ever merged names already verified to
    be bank-format truncations of the SAME transaction data.

    Prefix matching is dangerous here: "Deepak Singh" is a real prefix of
    "Deepak Singh Saini", but they are two different people with
    different account numbers - confirmed by a real false-merge caught in
    review (ratio 0.815, well below the threshold below). A pure ratio
    check on the full string catches genuine source-sheet typos instead
    ("Raj Tyre"/"Raj Tyres" 0.933, "Costify"/"Contify" 0.964,
    "Misthan"/"Mishthan" 0.976) without conflating different people."""
    import difflib

    a_raw = re.sub(r"[^A-Z0-9]", "", a.upper())
    b_raw = re.sub(r"[^A-Z0-9]", "", b.upper())
    if a_raw == b_raw:
        return True
    a_norm, b_norm = _fuzzy_key(a), _fuzzy_key(b)
    if a_norm == b_norm:
        return True
    if min(len(a_raw), len(b_raw)) < 5:
        return False  # too short for ratio-based matching to be reliable
    ratio = max(
        difflib.SequenceMatcher(None, a_raw, b_raw).ratio(),
        difflib.SequenceMatcher(None, a_norm, b_norm).ratio(),
    )
    return ratio >= 0.92


def load_bank_details() -> list[dict[str, str]]:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(BANK_DETAILS_SHEET_ID)
    ws = ss.worksheet("Sheet1")
    all_values = ws.get_all_values()
    header = all_values[1]
    idx = {col: header.index(col) for col in ("Co", "Party Name", "Type", "Account Number", "ISFC Code", "Bank Name")}

    records = []
    for row in all_values[2:]:
        if len(row) <= idx["Party Name"]:
            continue
        co = row[idx["Co"]].strip()
        if not co.upper().startswith("AMB"):
            continue
        party = row[idx["Party Name"]].strip()
        if not party:
            continue
        records.append({
            "party_name": party,
            "type": row[idx["Type"]].strip() if len(row) > idx["Type"] else "",
            "account_number": row[idx["Account Number"]].strip() if len(row) > idx["Account Number"] else "",
            "ifsc_code": row[idx["ISFC Code"]].strip() if len(row) > idx["ISFC Code"] else "",
            "bank_name": row[idx["Bank Name"]].strip() if len(row) > idx["Bank Name"] else "",
        })
    return records


def load_beneficiary_master() -> tuple[gspread.Worksheet, list[list[str]]]:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(AMB_SHEET_ID)
    ws = ss.worksheet("Beneficiary Master")
    all_values = ws.get_all_values()
    return ws, all_values


def build_plan() -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    """Returns (updates, new_rows, header)."""
    ws, all_values = load_beneficiary_master()
    header = all_values[0]
    data_rows = all_values[1:]
    name_idx = header.index("BENEFICIARY NAME")
    acct_idx = header.index("ACCOUNT NUMBER")
    ifsc_idx = header.index("IFSC CODE")
    bank_idx = header.index("BANK NAME")

    bank_details = load_bank_details()

    updates = []  # {"sheet_row": int, "name": str, "account_number", "ifsc_code", "bank_name"}
    new_rows = []  # {"name": str, "head1": str, "account_number", "ifsc_code", "bank_name", "type"}
    matched_existing_rows: set[int] = set()

    for record in bank_details:
        matched_row = None
        for offset, row in enumerate(data_rows):
            if offset in matched_existing_rows:
                continue
            existing_name = row[name_idx] if len(row) > name_idx else ""
            if existing_name and _names_match(record["party_name"], existing_name):
                matched_row = offset
                break

        if matched_row is not None:
            matched_existing_rows.add(matched_row)
            row = data_rows[matched_row]
            existing_acct = row[acct_idx] if len(row) > acct_idx else ""
            existing_ifsc = row[ifsc_idx] if len(row) > ifsc_idx else ""
            existing_bank = row[bank_idx] if len(row) > bank_idx else ""
            if existing_acct or existing_ifsc or existing_bank:
                continue  # never overwrite an existing value
            updates.append({
                "sheet_row": matched_row + 2,  # +1 header, +1 1-indexed
                "existing_name": row[name_idx],
                "account_number": record["account_number"],
                "ifsc_code": record["ifsc_code"],
                "bank_name": record["bank_name"],
            })
        else:
            # Also check against new_rows already queued this run - the
            # source sheet itself has duplicate party rows (e.g. the same
            # person listed once per account-stage tag), which would
            # otherwise become two identical Beneficiary Master entries.
            if any(_names_match(record["party_name"], r["name"]) for r in new_rows):
                continue
            head1 = _TYPE_TO_HEAD.get(record["type"].strip().lower(), "")
            new_rows.append({
                "name": record["party_name"].upper(),
                "head1": head1,
                "type": record["type"],
                "account_number": record["account_number"],
                "ifsc_code": record["ifsc_code"],
                "bank_name": record["bank_name"],
            })

    return updates, new_rows, header


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()

    updates, new_rows, header = build_plan()

    print(f"=== {len(updates)} existing beneficiaries to enrich ===\n")
    for u in updates:
        print(f"{u['existing_name']:<40} <- Acct={u['account_number']:<18} IFSC={u['ifsc_code']:<14} Bank={u['bank_name']}")

    print(f"\n=== {len(new_rows)} new beneficiaries to add ===\n")
    for r in new_rows:
        head_display = r["head1"] or "(BLANK - ambiguous type: " + r["type"] + ")"
        print(f"{r['name']:<40} Head1={head_display:<35} Acct={r['account_number']:<18} IFSC={r['ifsc_code']:<14} Bank={r['bank_name']}")

    if not args.write:
        print("\nDry run only - no changes written. Re-run with --write to apply.")
        return

    ws, _ = load_beneficiary_master()
    acct_idx = header.index("ACCOUNT NUMBER")
    ifsc_idx = header.index("IFSC CODE")
    bank_idx = header.index("BANK NAME")

    cell_updates = []
    for u in updates:
        cell_updates.append(gspread.cell.Cell(row=u["sheet_row"], col=acct_idx + 1, value=u["account_number"]))
        cell_updates.append(gspread.cell.Cell(row=u["sheet_row"], col=ifsc_idx + 1, value=u["ifsc_code"]))
        cell_updates.append(gspread.cell.Cell(row=u["sheet_row"], col=bank_idx + 1, value=u["bank_name"]))
    if cell_updates:
        ws.update_cells(cell_updates, value_input_option="RAW")
        print(f"\nUpdated {len(updates)} existing rows with account details.")

    if new_rows:
        values = []
        for r in new_rows:
            row = [""] * len(BENEFICIARY_MASTER_COLUMNS)
            row[header.index("BENEFICIARY NAME")] = r["name"]
            row[header.index("Head 1")] = r["head1"]
            row[header.index("ACCOUNT NUMBER")] = r["account_number"]
            row[header.index("IFSC CODE")] = r["ifsc_code"]
            row[header.index("BANK NAME")] = r["bank_name"]
            values.append(row)
        ws.append_rows(values, value_input_option="RAW")
        print(f"Added {len(new_rows)} new beneficiary rows.")


if __name__ == "__main__":
    main()
