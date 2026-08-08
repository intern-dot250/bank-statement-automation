"""One-off: build the initial AMB Beneficiary Master (BENEFICIARY NAME,
Head 1, Head 2 only) from the accounts team's own historical AMB ledger.

Source data: a local copy of "INTERN AMB Bank Statements 2026-27.xlsx"
(downloaded once via the Drive API from the accounts team's Drive file,
since that file is a raw .xlsx, not a native Google Sheet the Sheets API
can open directly). This ledger already carries real, accounts-team
-confirmed HEAD values per transaction across 5 account tabs (KVB Master,
KVB IDW, KVB Free, BOM RERA, Axis CA) — far richer ground truth than our
own automation sheet's handful of test-upload rows.

Also reuses config/party_master.json (72 pre-vetted beneficiary names +
aliases, already extracted from this exact same ledger in a prior
session) as the first, highest-confidence name-resolution step, falling
back to pattern-based extraction (stripping bank transaction IDs,
account numbers, and role-word suffixes) for every other transaction.

Usage:
    py -3 scripts/maintenance/build_amb_beneficiary_master.py --xlsx PATH  (dry run, default)
    py -3 scripts/maintenance/build_amb_beneficiary_master.py --xlsx PATH --write   (writes to the AMB sheet)
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import openpyxl

CONFIG_DIR = Path(__file__).resolve().parent.parent.parent / "config"
PARTY_MASTER_PATH = CONFIG_DIR / "party_master.json"

ACCOUNT_TABS = [
    "AMB MASTER KVB-6375",
    "IDW KVB-6535",
    "KVB FREE 1050",
    "AMB RERA BOM 667",
    "AMB CA AXIS-280",
]

# Raw HEAD values that are known data-entry noise, not real classifications.
_JUNK_HEADS = {"", "x", "X"}

# Raw HEAD -> canonical spelling, matching config/heads_config.json's actual
# vocabulary (the one already wired into live AMB classification via
# heads.py) or a previously-established fix (git history shows "DPL" was
# already corrected to "Internal" elsewhere in this project).
_HEAD_NORMALIZE = {
    "Card": "Credit Card",
    "DPL": "Internal",
}

# Role-word suffixes seen stripped off the end of KVB/IMPS-style
# descriptions (case-insensitive) - not an exhaustive whitelist, just used
# to confirm the second-to-last segment is a role word rather than part of
# the name itself, when a description has more than 3 dash-separated parts.
_ROLE_WORDS = {
    "salary", "vendor", "tfr", "imprest", "professional", "prof", "rent",
    "office rent", "hoarding", "commission", "land", "tds", "for tds",
    "for pf", "for esi", "for esi pf", "for esi  pf", "for esi epf",
    "master to rera", "master to free", "master to Free", "master to Rera",
    "internal", "rera", "advance", "broker", "brokrage", "broker comisson",
    "exp", "incentive", "salary advance", "for vendor paymet",
    "for dd dhbvn", "electricity", "maintenance", "contractor",
    "for epf payment", "security refund plot 15",
}


def _load_party_master() -> dict[str, tuple[str, str]]:
    """Return {normalized_alias_or_name: (canonical_name, type)}."""
    data = json.loads(PARTY_MASTER_PATH.read_text(encoding="utf-8"))
    lookup: dict[str, tuple[str, str]] = {}
    for canonical, info in data.get("parties", {}).items():
        ptype = info.get("type", "Unknown")
        for variant in [canonical, *info.get("aliases", [])]:
            lookup[_normalize_key(variant)] = (canonical, ptype)
    return lookup


def _normalize_key(text: str) -> str:
    """Case/space/punctuation-insensitive key for matching name variants."""
    return re.sub(r"[^A-Z0-9]", "", text.upper())


_DIGITS_RE = re.compile(r"^\d{6,}$")
_KVB_TXN_ID_RE = re.compile(r"^KVBL[HRN]\d+$", re.IGNORECASE)


def _looks_like_txn_id_or_account(segment: str) -> bool:
    seg = segment.strip()
    if not seg:
        return True
    if _DIGITS_RE.match(seg):
        return True
    if _KVB_TXN_ID_RE.match(seg):
        return True
    return False


_CARD_SUFFIX_RE = re.compile(r"^(axis\s+)?(credit\s+)?card\s*\d*$", re.IGNORECASE)


def _looks_like_role_word(segment: str) -> bool:
    seg = segment.strip().lower()
    return seg in _ROLE_WORDS or bool(_CARD_SUFFIX_RE.match(seg))


def extract_name(description: str) -> str | None:
    """Best-effort extraction of the beneficiary/party name from a raw
    transaction description. Returns None if no real, distinguishable
    party name survives (pure boilerplate/generic charge text)."""
    desc = " ".join(description.split()).strip()
    if not desc:
        return None

    # "BY CLG:<Name>:<Bank> - <date>"
    m = re.match(r"^BY CLG:(.+?):", desc, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # "DD ISSUED/SAK/<Name...>" (name may be truncated by the source PDF)
    m = re.match(r"^DD ISSUED/SAK/(.+)$", desc, re.IGNORECASE)
    if m:
        name = re.sub(r"/atPar$", "", m.group(1).strip(), flags=re.IGNORECASE)
        return name if len(name) > 2 else None

    # Generic "DD Cancln ####" / "DD CANC & GST" / "By DD Num #### Paid" -
    # no real distinguishable party name.
    if re.match(r"^(DD Cancln|DD CANC|By DD Num)\b", desc, re.IGNORECASE):
        return None

    # "NEFT CR-<IFSC>-<Name1>-<Name2>-<UTR>" / "RTGS CR-<IFSC>-<Name1>-<Name2>-<UTR>"
    m = re.match(r"^(?:NEFT|RTGS) CR-[A-Z0-9]+-(.+)-[A-Z0-9]{10,}$", desc, re.IGNORECASE)
    if m:
        middle = m.group(1)
        parts = [p.strip() for p in middle.split("-") if p.strip()]
        return parts[0] if parts else None

    # BOM's own space-delimited format: "NEFT <BANKCODE+digits glued together>
    # <Name...> <branch code>" (no dash/slash at all - the UTR code and name
    # run together after one space, and a trailing branch code like
    # "KVBL0004114" follows the name), e.g.
    # "NEFT MAHBN12026040647875203 AMbition Colonisers Pr KVBL0004114".
    m = re.match(r"^(?:NEFT|RTGS)\s+[A-Z0-9]{10,}\s+(.+?)\s+[A-Z]{2,6}\d{4,}\s*$", desc, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.match(r"^(?:NEFT|RTGS)\s+[A-Z0-9]{10,}\s+(.+)$", desc, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # "INB/IFT/<Name>/TPARTY TRANSFER" - one observed row has the bank name
    # glued directly onto the name with no space ("SHOBHAAXIS" instead of
    # "SHOBHA .../AXIS"), a source-PDF artifact; strip a trailing "AXIS" in
    # that specific glued form so it merges with the same person's other,
    # correctly-spaced rows ("Shobha Jain") instead of becoming a
    # standalone one-off entry.
    m = re.match(r"^INB/IFT/(.+)/TPARTY TRANSFER$", desc, re.IGNORECASE)
    if m:
        name = m.group(1).strip()
        name = re.sub(r"AXIS$", "", name, flags=re.IGNORECASE).strip()
        return name

    # "NEFT/<code>/<NAME>/<BANK>/..." (Axis slash format)
    if desc.upper().startswith("NEFT/") and "/" in desc[5:]:
        parts = [p.strip() for p in desc.split("/")]
        if len(parts) >= 3 and not _looks_like_txn_id_or_account(parts[2]):
            return parts[2]

    # "RTGS/<code>/<Name>/<Bank>/..." (Axis slash format)
    if desc.upper().startswith("RTGS/"):
        parts = [p.strip() for p in desc.split("/")]
        if len(parts) >= 3 and not _looks_like_txn_id_or_account(parts[2]):
            return parts[2]

    # "IMPS-<digits>-<Name>-<bankcode>-<maskedacct>-<role>"
    m = re.match(r"^IMPS-\d+-(.+?)-[A-Z]{2,6}-", desc, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # "BILLDESK-<code>-<Name>-..."
    m = re.match(r"^BILLDESK-[A-Z0-9]+-(.+?)-", desc, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # "NEFT-RETURN-<txnid>-<Name>-..."
    m = re.match(r"^NEFT-RETURN-[A-Z0-9]+-(.+?)-", desc, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # Generic "INB/<digits>/<Name>/NA" (skip if <Name> is one of the known
    # institutional keywords already handled by heads_config.json's own
    # keyword rules with no real distinguishable party).
    m = re.match(r"^INB/\d+/(.+)/NA$", desc, re.IGNORECASE)
    if m:
        name = m.group(1).strip()
        if not re.search(r"(epfo payment|esic payment|tin 2\.0)", name, re.IGNORECASE):
            return name

    # KVB dash format: "KVBL[HRN]<digits>-<Name>-<account>[-<role>]" or
    # "KVBL[HRN]<digits>-<Name>-<account>" (no trailing role).
    if _KVB_TXN_ID_RE.match(desc.split("-")[0]):
        parts = [p.strip() for p in desc.split("-")]
        parts = parts[1:]  # drop the leading transaction-ID segment
        if parts and _looks_like_txn_id_or_account(parts[-1]):
            parts = parts[:-1]  # drop trailing bare account number, if present
        elif len(parts) >= 2 and _looks_like_role_word(parts[-1]):
            parts = parts[:-1]
            if parts and _looks_like_txn_id_or_account(parts[-1]):
                parts = parts[:-1]
        name = "-".join(parts).strip()
        return name if len(name) > 1 and not _looks_like_txn_id_or_account(name) else None

    # Generic boilerplate with no distinguishable party (charges, taxes,
    # opening balance) - explicitly no beneficiary here.
    if re.search(r"(monthly service chrgs|sms charges|gst @|tin 2\.0|epfo payment|esic payment)", desc, re.IGNORECASE):
        return None

    return None


def _split_and_tally(rows: list[dict[str, Any]], party_lookup: dict[str, tuple[str, str]]) -> dict[str, Counter]:
    """Group rows by canonical beneficiary name, tallying each Head's
    occurrence count."""
    grouped: dict[str, Counter] = defaultdict(Counter)
    unresolved: list[dict[str, Any]] = []

    for row in rows:
        desc = row["description"]
        head_raw = row["head"].strip()
        if not desc or desc.strip().lower() in ("b/f", "b/f...", "b/f…"):
            continue
        if head_raw in _JUNK_HEADS:
            continue
        head = _HEAD_NORMALIZE.get(head_raw, head_raw)

        key = _normalize_key(desc)
        canonical = None
        # Try a party_master match against progressively shorter prefixes
        # of the description's dash/slash-delimited segments, since a full
        # description rarely equals a stored alias exactly.
        for segment in re.split(r"[-/]", desc):
            seg = segment.strip()
            if len(seg) < 3:
                continue
            hit = party_lookup.get(_normalize_key(seg))
            if hit:
                canonical = hit[0]
                break

        if canonical is None:
            extracted = extract_name(desc)
            if extracted:
                # The extracted name itself might exactly be a known
                # party_master alias (e.g. extract_name() pulls "AMbition
                # Colonisers Pr" out of a BOM-format description, which is
                # itself a stored alias for "Ambition Colonisers Private
                # Limited") - canonicalize it the same way, rather than
                # only checking the original description's raw segments.
                hit = party_lookup.get(_normalize_key(extracted))
                if hit:
                    canonical = hit[0]
            if canonical is None and not extracted:
                unresolved.append(row)
                continue
            if canonical is None:
                canonical = " ".join(extracted.split()).strip()

        canonical_upper = canonical.upper()
        grouped[canonical_upper][head] += 1

    grouped, merges = _fuzzy_merge(grouped)
    if merges:
        print("=== Fuzzy-merged near-duplicate names ===")
        for a, b in merges:
            print(f"  {a!r}  ->  {b!r}")
        print()

    return grouped, unresolved


def _fuzzy_merge(grouped: dict[str, Counter]) -> tuple[dict[str, Counter], list[tuple[str, str]]]:
    """Merge names that are near-identical (one is a truncated prefix of
    the other, or >=92% sequence-similar) - covers source-PDF truncation
    (e.g. "...TRIBUN" vs "...TRIBUNAL") and minor spelling variants across
    banks (e.g. "HDFC BANK LTD" vs "HDFC BANK LIMITED"). Conservative on
    purpose: only merges very-high-confidence pairs, everything else is
    left for manual review in the dry-run output."""
    import difflib

    # Corporate-suffix abbreviations collapsed to one token before
    # comparing, so "LIMITED" vs "LTD" / "PRIVATE LIMITED" vs "PVT LTD"
    # don't block an otherwise-identical name from merging.
    _SUFFIX_EQUIV = [
        (r"\bPRIVATE LIMITED\b", "PVTLTD"),
        (r"\bPVT\.?\s*LTD\.?\b", "PVTLTD"),
        (r"\bLIMITED\b", "LTD"),
    ]

    def _fuzzy_key(name: str) -> str:
        text = name.upper()
        for pattern, repl in _SUFFIX_EQUIV:
            text = re.sub(pattern, repl, text)
        return re.sub(r"[^A-Z0-9]", "", text)

    names = sorted(grouped.keys(), key=len, reverse=True)
    canonical_for: dict[str, str] = {n: n for n in names}
    merges: list[tuple[str, str]] = []

    for i, a in enumerate(names):
        if canonical_for[a] != a:
            continue  # already merged into something else
        for b in names[i + 1:]:
            if canonical_for[b] != b:
                continue
            # Check both the raw alnum-only keys AND the suffix-normalized
            # keys - normalizing "Limited"->"Ltd" fixes one class of near-
            # duplicate but can itself break a plain-truncation prefix
            # match (e.g. "AMBITIONCOLONISERSPRIV" is a prefix of the raw
            # canonical name, but not of its suffix-normalized form), so
            # a match on either representation is accepted.
            a_raw, b_raw = re.sub(r"[^A-Z0-9]", "", a.upper()), re.sub(r"[^A-Z0-9]", "", b.upper())
            a_norm, b_norm = _fuzzy_key(a), _fuzzy_key(b)
            if len(b_raw) < 5:
                continue
            is_prefix = (
                a_raw.startswith(b_raw) or b_raw.startswith(a_raw)
                or a_norm.startswith(b_norm) or b_norm.startswith(a_norm)
            )
            ratio = max(
                difflib.SequenceMatcher(None, a_raw, b_raw).ratio(),
                difflib.SequenceMatcher(None, a_norm, b_norm).ratio(),
            )
            if is_prefix or ratio >= 0.92:
                # Keep whichever name is longer (more complete/less truncated).
                keep, drop = (a, b) if len(a) >= len(b) else (b, a)
                canonical_for[drop] = keep
                merges.append((drop, keep))

    merged: dict[str, Counter] = defaultdict(Counter)
    for name, counts in grouped.items():
        target = canonical_for[name]
        while canonical_for[target] != target:
            target = canonical_for[target]
        merged[target].update(counts)

    return dict(merged), merges


def build_master_rows(xlsx_path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True, read_only=True)
    all_rows: list[dict[str, Any]] = []
    for tab in ACCOUNT_TABS:
        ws = wb[tab]
        raw_rows = list(ws.iter_rows(values_only=True))
        header_idx = next(
            (i for i, r in enumerate(raw_rows) if r and "DESCRIPTION" in r and "HEAD" in r),
            None,
        )
        if header_idx is None:
            continue
        header = raw_rows[header_idx]
        desc_i = header.index("DESCRIPTION")
        head_i = header.index("HEAD")
        for r in raw_rows[header_idx + 1:]:
            if r is None:
                continue
            desc = r[desc_i] if desc_i < len(r) else None
            head = r[head_i] if head_i < len(r) else None
            if desc is None and head is None:
                continue
            all_rows.append({
                "tab": tab,
                "description": str(desc).strip() if desc else "",
                "head": str(head).strip() if head else "",
            })

    party_lookup = _load_party_master()
    grouped, unresolved = _split_and_tally(all_rows, party_lookup)

    master_rows = []
    for name, head_counts in sorted(grouped.items()):
        ranked = head_counts.most_common()
        head1 = ranked[0][0]
        head2 = ranked[1][0] if len(ranked) > 1 and ranked[1][1] >= 2 else ""
        master_rows.append({
            "name": name,
            "head1": head1,
            "head2": head2,
            "txn_count": sum(head_counts.values()),
            "head_breakdown": dict(ranked),
        })

    master_rows.sort(key=lambda r: (-r["txn_count"], r["name"]))
    return master_rows, unresolved


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx", required=True, help="Path to the downloaded AMB reference .xlsx")
    parser.add_argument("--write", action="store_true", help="Actually write to the AMB Beneficiary Master tab (default: dry run)")
    args = parser.parse_args()

    master_rows, unresolved = build_master_rows(Path(args.xlsx))

    print(f"=== {len(master_rows)} candidate beneficiaries ===\n")
    for r in master_rows:
        print(f"{r['name']:<45} Head1={r['head1']:<20} Head2={r['head2']:<20} "
              f"(txns={r['txn_count']}, breakdown={r['head_breakdown']})")

    print(f"\n=== {len(unresolved)} rows with no extractable beneficiary (skipped) ===\n")
    for r in unresolved[:60]:
        print(f"  [{r['head']}] {r['description']}")
    if len(unresolved) > 60:
        print(f"  ... and {len(unresolved) - 60} more")

    if not args.write:
        print("\nDry run only - no changes written. Re-run with --write to populate the AMB Beneficiary Master tab.")
        return

    from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS
    import gspread

    AMB_SHEET_ID = "1kVMuah99dU8g3q9zsHxiTtBh7l7E8xIPrFoSmGQpywc"
    BENEFICIARY_MASTER_COLUMNS = [
        "BENEFICIARY NAME", "Head 1", "Head 2", "Head 3", "NOTES", "ADDED BY",
        "DATE ADDED", "STATUS", "ACCOUNT NUMBER", "IFSC CODE", "BANK NAME", "Company",
    ]

    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(AMB_SHEET_ID)
    try:
        ws = ss.worksheet("Beneficiary Master")
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title="Beneficiary Master", rows=len(master_rows) + 10, cols=len(BENEFICIARY_MASTER_COLUMNS))
        ws.update(range_name="A1", values=[BENEFICIARY_MASTER_COLUMNS])
        print("Created 'Beneficiary Master' tab in the AMB sheet.")

    values = []
    for r in master_rows:
        row = [""] * len(BENEFICIARY_MASTER_COLUMNS)
        row[0] = r["name"]
        row[1] = r["head1"]
        row[2] = r["head2"]
        values.append(row)

    ws.append_rows(values, value_input_option="RAW")
    print(f"\nWrote {len(values)} beneficiary rows to the AMB Beneficiary Master tab.")


if __name__ == "__main__":
    main()
