"""Phase 1: Build the Beneficiary Master tab in Google Sheet.

Scans all classified rows across every account tab, extracts the
beneficiary name from the DESCRIPTION field, and writes a unique
(BENEFICIARY NAME, HEAD) pair per row into a new "Beneficiary Master"
tab — ready for the accounts team to review and edit directly.

New rows land as STATUS="Pending" (or "Conflict" if the same name is
already recorded under a different head) — export_cache() only pulls
"Confirmed" rows into the lookup file classify_transactions.py actually
trusts, so nothing reaches auto-classification until the accounts team
reviews it and flips the status to Confirmed.

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

from __future__ import annotations

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

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
    "STATUS",
]

# Status values (STATUS column). "Confirmed" matches rag_classifier.py's
# MASTER_STATUS_CONFIRMED — the two scripts write to the same tab, so the
# string must stay identical. Rows discovered here (a full historical
# rescan, not reviewed by anyone) start as "Pending" rather than
# "Confirmed" so a wrong extraction can't silently poison future
# classification until a human checks it — see export_cache().
STATUS_CONFIRMED = "Confirmed"
STATUS_PENDING = "Pending"
STATUS_CONFLICT = "Conflict"

TODAY = datetime.date.today().strftime("%d-%b-%Y")

# NOTE: must match classify_transactions.py's _BENEFICIARY_CACHE_PATH.
# This used to point at scripts/debug/beneficiary_master_cache.json — a
# leftover from before the project's config/data folders were
# reorganised — so re-running this script silently wrote to a file
# classify_transactions.py never read, and the real cache went stale.
CACHE_FILE = Path(__file__).resolve().parent.parent.parent / "data" / "beneficiary_master_cache.json"


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
    ws.format("A1:F1", {"textFormat": {"bold": True}})
    print(f"[CREATE] '{MASTER_TAB_NAME}' tab created.")
    return ws


def build_master(spreadsheet: gspread.Spreadsheet) -> None:
    """Rescan every account tab and add any (name, head) pair not already
    in the Beneficiary Master. New rows are written as STATUS="Pending" —
    a full historical rescan has no human review behind it, so it must
    not be trusted the same as a confirmed entry (see export_cache()).

    If a name already exists in the master under a DIFFERENT head, that's
    a conflict (e.g. one occurrence was misclassified) — both the new row
    and every existing row for that name are flagged STATUS="Conflict"
    instead of silently adding a second, contradictory entry.
    """
    beneficiaries = collect_beneficiaries(spreadsheet)
    if not beneficiaries:
        print("[WARN] No classifiable beneficiaries found.")
        return

    ws = get_or_create_master_tab(spreadsheet)

    # Read what's already there to avoid duplicates, and to detect conflicts.
    existing_rows = ws.get_all_values()
    existing_keys: set[tuple[str, str]] = set()
    # name -> list of (sheet_row_number, head, current_status)
    existing_by_name: dict[str, list[tuple[int, str, str]]] = {}
    hi_status = None
    if len(existing_rows) > 1:
        try:
            hi_name = existing_rows[0].index("BENEFICIARY NAME")
            hi_head = existing_rows[0].index("HEAD")
            hi_status = existing_rows[0].index("STATUS") if "STATUS" in existing_rows[0] else None
            for i, row in enumerate(existing_rows[1:], start=2):
                if len(row) > max(hi_name, hi_head):
                    name = row[hi_name].upper()
                    head = row[hi_head]
                    status = row[hi_status] if hi_status is not None and len(row) > hi_status else ""
                    existing_keys.add((name, head))
                    existing_by_name.setdefault(name, []).append((i, head, status))
        except ValueError:
            pass

    new_rows = []
    rows_to_flag_conflict: list[int] = []  # existing sheet row numbers to update
    n_pending = 0
    n_conflict = 0
    for (name, head), count in sorted(beneficiaries.items()):
        if (name, head) in existing_keys:
            continue
        notes = f"Auto-extracted ({count} txn)" if count > 1 else "Auto-extracted"

        conflicting = [
            (row_num, other_head, status)
            for row_num, other_head, status in existing_by_name.get(name, [])
            if other_head != head
        ]
        if conflicting:
            status = STATUS_CONFLICT
            n_conflict += 1
            for row_num, _other_head, other_status in conflicting:
                if other_status != STATUS_CONFLICT:
                    rows_to_flag_conflict.append(row_num)
        else:
            status = STATUS_PENDING
            n_pending += 1

        new_rows.append([name, head, notes, "System", TODAY, status])

    if rows_to_flag_conflict and hi_status is not None:
        status_col_letter = gspread.utils.rowcol_to_a1(1, hi_status + 1).rstrip("0123456789")
        ws.batch_update([
            {"range": f"{status_col_letter}{row_num}", "values": [[STATUS_CONFLICT]]}
            for row_num in rows_to_flag_conflict
        ])
        print(f"[FLAG] Marked {len(rows_to_flag_conflict)} existing row(s) as Conflict "
              f"(same name, different head found).")

    if new_rows:
        ws.append_rows(new_rows, value_input_option="RAW")
        print(f"[OK] Added {len(new_rows)} beneficiaries to '{MASTER_TAB_NAME}' "
              f"({n_pending} Pending, {n_conflict} Conflict).")
    else:
        print("[OK] No new beneficiaries to add — tab is already up to date.")

    print(f"     Total unique (name, head) pairs: {len(beneficiaries)}")
    print(f"     Already existed: {len(existing_keys)}")
    print(f"     Newly added: {len(new_rows)}")


def export_cache(spreadsheet: gspread.Spreadsheet) -> None:
    """Write the current Beneficiary Master tab to a local JSON cache file
    so classify_transactions.py can do fast offline lookups.

    Only STATUS="Confirmed" rows are included. "Pending" (freshly
    discovered, unreviewed) and "Conflict" (name maps to more than one
    head) rows are excluded — Rule 6 in classify_transactions.py trusts
    this cache completely and applies it ahead of every other rule, so an
    unreviewed or contradictory entry must never reach it. A row with no
    STATUS value at all (pre-dates the STATUS column) is treated as
    Confirmed, matching rag_classifier.py's _ensure_status_column()
    migration, which back-filled every pre-existing row that way.
    """
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
    si = hdr.index("STATUS") if "STATUS" in hdr else None

    cache: dict[str, str] = {}
    n_skipped = 0
    for row in rows[1:]:
        if len(row) <= max(ni, hi):
            continue
        status = row[si].strip() if si is not None and len(row) > si else STATUS_CONFIRMED
        if status and status != STATUS_CONFIRMED:
            n_skipped += 1
            continue
        name = row[ni].strip().upper()
        head = row[hi].strip()
        if name and head:
            cache[name] = head

    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    CACHE_FILE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"[OK] Cache written: {len(cache)} entries -> {CACHE_FILE} "
          f"({n_skipped} unconfirmed row(s) excluded)")


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)
    build_master(spreadsheet)
    export_cache(spreadsheet)


if __name__ == "__main__":
    main()
