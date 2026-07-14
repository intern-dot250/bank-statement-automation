"""Phase 1: Build the Beneficiary Master tab in Google Sheet.

Scans all classified rows across every account tab, extracts the
beneficiary name from the DESCRIPTION field, and writes a unique
(BENEFICIARY NAME, HEAD) pair per row into a new "Beneficiary Master"
tab — ready for the accounts team to review and edit directly.

Description parsing logic
--------------------------
NEFT/TPT format:
  YIB-NEFT-{ref}-{NAME}-{IFSC}-{ROLE}-{BANK}
  Split by "-", name = segment[3]

IMPS format:
  IMPS/NA/{ref}/{RRN}/{PC}/{BANK}/{NAME}/{NAME ROLE}
  Split by "/", name = segment[-2] (second-to-last)

CHQ DEP / other:
  Name cannot be reliably extracted — these rows are skipped.

Role keywords at the end of a name are stripped
(e.g. "YOGESH SINGH IMPREST" → "YOGESH SINGH").
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __future__ import annotations

import datetime
import json
import re
from pathlib import Path

import gspread

from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
    get_account_worksheets,
)

# Heads where the transaction does not have a single named beneficiary
SKIP_HEADS = {
    "Internal", "Collection", "Cancellation",
    "?", "", "HO-Admin",
}

# Role keywords that sometimes appear at the end of a name in the description
_ROLE_SUFFIX = re.compile(
    r"\s+(IMPREST|SALARY|CONTRACTOR|PROFESSIONAL|VENDOR|ADVANCE|REFUND)$",
    re.IGNORECASE,
)

# IFSC code pattern (uppercase letters + digits, 11 chars)
_IFSC_LIKE = re.compile(r"^[A-Z]{4}0[A-Z0-9]{6}$")

MASTER_TAB_NAME = "Beneficiary Master"
MASTER_HEADERS = [
    "BENEFICIARY NAME",
    "HEAD",
    "NOTES",
    "ADDED BY",
    "DATE ADDED",
]

TODAY = datetime.date.today().strftime("%d-%b-%Y")

CACHE_FILE = Path(__file__).resolve().parent / "beneficiary_master_cache.json"


def _extract_name_neft(desc: str) -> str | None:
    """Extract name from YIB-NEFT-... or YIB-TPT-... descriptions."""
    parts = desc.split("-")
    # Expect at least: YIB | method | ref | NAME | ...
    if len(parts) < 4:
        return None
    name = parts[3].strip()
    # Skip if it looks like an IFSC or a reference number
    if not name or _IFSC_LIKE.match(name) or name.isdigit():
        return None
    name = _ROLE_SUFFIX.sub("", name).strip()
    return name or None


def _extract_name_imps(desc: str) -> str | None:
    """Extract name from IMPS/... descriptions."""
    parts = desc.split("/")
    # Second-to-last segment is typically the person's name
    if len(parts) < 3:
        return None
    name = parts[-2].strip()
    if not name or name.upper().startswith("RRN") or name.isdigit():
        return None
    name = _ROLE_SUFFIX.sub("", name).strip()
    return name or None


def extract_beneficiary(desc: str) -> str | None:
    """Return the beneficiary name from a bank description, or None."""
    upper = desc.upper()
    if upper.startswith("YIB-NEFT") or upper.startswith("YIB-TPT"):
        return _extract_name_neft(desc)
    if upper.startswith("IMPS/"):
        return _extract_name_imps(desc)
    return None  # CHQ DEP, B/F, and other formats skipped


def collect_beneficiaries(spreadsheet: gspread.Spreadsheet) -> dict[tuple[str, str], int]:
    """Scan all account tabs and return {(name, head): count} of classified rows."""
    results: dict[tuple[str, str], int] = {}

    for ws in get_account_worksheets(spreadsheet):
        rows = ws.get_all_values()
        if not rows:
            continue
        hdr = rows[0]
        if "DESCRIPTION" not in hdr or "HEAD" not in hdr:
            continue
        di = hdr.index("DESCRIPTION")
        hi = hdr.index("HEAD")

        for row in rows[1:]:
            if len(row) <= max(di, hi):
                continue
            desc = row[di].strip()
            head = row[hi].strip()

            if not desc or head in SKIP_HEADS:
                continue

            name = extract_beneficiary(desc)
            if not name:
                continue

            key = (name.upper(), head)
            results[key] = results.get(key, 0) + 1

    return results


def get_or_create_master_tab(spreadsheet: gspread.Spreadsheet) -> gspread.Worksheet:
    """Return the Beneficiary Master worksheet, creating it if needed."""
    existing = {ws.title: ws for ws in spreadsheet.worksheets()}
    if MASTER_TAB_NAME in existing:
        return existing[MASTER_TAB_NAME]

    ws = spreadsheet.add_worksheet(title=MASTER_TAB_NAME, rows="2000", cols="10")
    ws.append_row(MASTER_HEADERS, value_input_option="RAW")

    # Bold header
    ws.format("A1:E1", {"textFormat": {"bold": True}})
    print(f"[CREATE] '{MASTER_TAB_NAME}' tab created.")
    return ws


def build_master(spreadsheet: gspread.Spreadsheet) -> None:
    beneficiaries = collect_beneficiaries(spreadsheet)
    if not beneficiaries:
        print("[WARN] No classifiable beneficiaries found.")
        return

    ws = get_or_create_master_tab(spreadsheet)

    # Read what's already there to avoid duplicates
    existing_rows = ws.get_all_values()
    existing_keys: set[tuple[str, str]] = set()
    if len(existing_rows) > 1:
        try:
            hi_name = existing_rows[0].index("BENEFICIARY NAME")
            hi_head = existing_rows[0].index("HEAD")
            for row in existing_rows[1:]:
                if len(row) > max(hi_name, hi_head):
                    existing_keys.add((row[hi_name].upper(), row[hi_head]))
        except ValueError:
            pass

    new_rows = []
    for (name, head), count in sorted(beneficiaries.items()):
        if (name, head) in existing_keys:
            continue
        notes = f"Auto-extracted ({count} txn)" if count > 1 else "Auto-extracted"
        new_rows.append([name, head, notes, "System", TODAY])

    if new_rows:
        ws.append_rows(new_rows, value_input_option="RAW")
        print(f"[OK] Added {len(new_rows)} beneficiaries to '{MASTER_TAB_NAME}'.")
    else:
        print("[OK] No new beneficiaries to add — tab is already up to date.")

    print(f"     Total unique (name, head) pairs: {len(beneficiaries)}")
    print(f"     Already existed: {len(existing_keys)}")
    print(f"     Newly added: {len(new_rows)}")


def export_cache(spreadsheet: gspread.Spreadsheet) -> None:
    """Write the current Beneficiary Master tab to a local JSON cache file
    so classify_transactions.py can do fast offline lookups."""
    try:
        ws = spreadsheet.worksheet(MASTER_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        print(f"[WARN] '{MASTER_TAB_NAME}' tab not found — cache not written.")
        return

    rows = ws.get_all_values()
    if len(rows) < 2:
        print("[WARN] Beneficiary Master tab is empty — cache not written.")
        return

    hdr = rows[0]
    try:
        ni = hdr.index("BENEFICIARY NAME")
        hi = hdr.index("HEAD")
    except ValueError:
        print("[WARN] Missing expected columns in Beneficiary Master — cache not written.")
        return

    cache: dict[str, str] = {}
    for row in rows[1:]:
        if len(row) > max(ni, hi):
            name = row[ni].strip().upper()
            head = row[hi].strip()
            if name and head:
                cache[name] = head

    CACHE_FILE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"[OK] Cache written: {len(cache)} entries -> {CACHE_FILE.name}")


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)
    build_master(spreadsheet)
    export_cache(spreadsheet)


if __name__ == "__main__":
    main()
