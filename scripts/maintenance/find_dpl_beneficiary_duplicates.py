"""Read-only report: find likely-duplicate names already sitting in
DPL's live Beneficiary Master tab (e.g. "RAM CHAND ADAV" vs "RAM CHAND
YADAV", "A N FILLING STA TION" vs "A FILLING STATION" vs "A N FILLING
STATION") — a one-time backlog scan using the same fuzzy-match rule
now also applied live, going forward, in classify_transactions.py's
_update_beneficiary_master() (see beneficiary_similarity.py).

This script only reads and prints — it never writes to the sheet.
Merging/fixing a flagged cluster is a manual decision made via the
Beneficiary Master web UI (Admin isn't involved; any user with access
to /beneficiary_master can edit/delete a row there).

Usage:
    py -3 scripts/maintenance/find_dpl_beneficiary_duplicates.py
"""

from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from beneficiary_similarity import is_fuzzy_duplicate
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client

BENEFICIARY_MASTER_TAB_NAME = "Beneficiary Master"


def find_clusters(names_with_rows: list[tuple[str, int]]) -> list[list[tuple[str, int]]]:
    """Group (name, row_num) pairs into clusters of mutual fuzzy
    duplicates. A name only ever belongs to one cluster (first match
    wins) — good enough for a human-facing report, not meant to be a
    perfect transitive-closure clustering."""
    clusters: list[list[tuple[str, int]]] = []
    assigned: set[str] = set()

    for name, row_num in names_with_rows:
        if name in assigned:
            continue
        cluster = [(name, row_num)]
        assigned.add(name)
        for other_name, other_row in names_with_rows:
            if other_name in assigned:
                continue
            if is_fuzzy_duplicate(name, other_name):
                cluster.append((other_name, other_row))
                assigned.add(other_name)
        if len(cluster) > 1:
            clusters.append(cluster)

    return clusters


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)
    ws = spreadsheet.worksheet(BENEFICIARY_MASTER_TAB_NAME)

    rows = ws.get_all_values()
    if not rows:
        print("Beneficiary Master tab is empty.")
        return

    hdr = rows[0]
    if "BENEFICIARY NAME" not in hdr:
        print(f"'BENEFICIARY NAME' column not found in header: {hdr}")
        return
    ni = hdr.index("BENEFICIARY NAME")
    si = hdr.index("STATUS") if "STATUS" in hdr else None

    names_with_rows: list[tuple[str, int]] = []
    status_by_name: dict[str, list[str]] = defaultdict(list)
    for i, row in enumerate(rows[1:], start=2):
        if len(row) <= ni or not row[ni].strip():
            continue
        name = row[ni].strip().upper()
        names_with_rows.append((name, i))
        status_by_name[name].append(row[si].strip() if si is not None and len(row) > si else "")

    # De-duplicate identical (name, row) pairs isn't needed - each row is
    # its own entry - but multiple rows can legitimately share the exact
    # same name (e.g. one Confirmed + one Conflict, per the existing
    # per-name-per-head logic), so this list intentionally keeps every
    # row rather than collapsing by name first.
    clusters = find_clusters(names_with_rows)

    if not clusters:
        print(f"No likely-duplicate name clusters found among {len(names_with_rows)} row(s).")
        return

    print(f"Found {len(clusters)} likely-duplicate cluster(s) among {len(names_with_rows)} row(s):\n")
    for cluster in clusters:
        for name, row_num in cluster:
            statuses = ", ".join(status_by_name[name]) or "(no status)"
            print(f"  row {row_num:>4}  {name!r:40}  status={statuses}")
        print()


if __name__ == "__main__":
    main()
