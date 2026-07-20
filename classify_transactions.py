"""Phase 2: Classify transactions in the master Google Sheet.

Reads every row from the existing master worksheet (the same sheet that
upload_to_sheets.py writes to), assigns a business ``Head`` and a
human-readable ``Narration`` to each transaction, and writes both values
back into the SAME row. No rows are appended or duplicated.
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path
from typing import Any, Optional

import gspread
from gspread.utils import rowcol_to_a1

from description_parser import parse_description
from heads import get_head, is_internal_type_head
from narration import generate_narration, OWN_COMPANY_KEYWORDS
from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
)
import credentials_store

SCRIPT_DIR_FOR_RECORDS = Path(__file__).resolve().parent
_RECORDS_FALLBACK_PATH = SCRIPT_DIR_FOR_RECORDS / "data" / "records.json"

# Stage-pair -> Type for RERA IDW label, for internal transfers between two
# of our own tracked accounts. Only pairs confidently confirmed from the
# accounts team's reference sheet are included — any other pair (e.g.
# RERA <-> IDW, which the reference data shows using two DIFFERENT labels
# depending on transaction specifics we can't reliably tell apart from the
# description alone) is intentionally left unmapped — see
# _AMBIGUOUS_STAGE_PAIRS below, which distinguishes "genuinely
# contradictory in the source data" (stays "?") from "just not a
# Casa-Romana-pipeline pair" (resolves to the literal "Internal" label the
# accounts team itself uses for those, e.g. any transfer involving the
# Aravali Heights accounts).
TRANSFER_STAGE_LABELS: dict[frozenset[str], str] = {
    frozenset({"Master", "Free"}): "Master to Free",
    frozenset({"Master", "RERA"}): "Master 2 RERA",
    frozenset({"Free", "IDW"}): "Free & IDW Loan",
}

# Stage pairs where the reference sheet itself uses two DIFFERENT labels
# unpredictably for the same pair (confirmed: YES Rera 0377 and YES IDW
# 0490 both show a RERA<->IDW transfer labeled "RERA IDW New" in some
# rows and "RERA 2 IDW" in others, with no distinguishing signal in the
# description) — genuinely contradictory, so this one stays "?" rather
# than resolving to "Internal" like other unmapped pairs do.
_AMBIGUOUS_STAGE_PAIRS = {frozenset({"RERA", "IDW"})}

# Description prefixes that indicate an incoming payment from an external
# party (as opposed to a transfer between our own tracked accounts).
_INCOMING_PAYMENT_PREFIXES = ("UPI/", "NEFT CR-", "IMPS/", "RTGS CR-", "NET-TPT-", "NET-")

# These bank statement descriptions spell the payment's role out directly as
# one of the hyphen-separated segments (e.g.
# "YIB-NEFT-YESME61850064653-Lalan Yadav-FDRL0002158-contractor-FEDERAL
# BANK") — a much more reliable signal than any keyword-in-free-text
# heuristic, since it's literally printed there by the bank. Keys are
# lowercase for case-insensitive matching; values are the exact Head label
# the accounts team's reference sheet uses. "salary" is handled separately
# (see _resolve_salary_head) since its Head name depends on account stage.
DESCRIPTION_ROLE_TO_HEAD = {
    "vendor": "Vendor",
    "contractor": "Contractor",
    "contract": "Contractor",   # Rule 12: bare "contract" segment = Contractor
    "professional": "Professional",
    "prof": "Professional",     # staff abbreviation, e.g. "-PROF-HDFC BANK"
    "imprest": "Imprest",
}

# Role segments that are followed by a reference code in the SAME
# hyphen-separated segment (e.g. "-Cancellation D6126-", "-Cancellation
# D2025-"), so an exact-match lookup in DESCRIPTION_ROLE_TO_HEAD would
# miss them — matched by prefix instead. Confirmed from the reference
# sheet: Head "Cancellation", Business Unit = the account's own project,
# Type for RERA IDW "Cust Cancellation" (91% of observed cases), TCP Head
# consistently absent/"?" in every observed case (a genuine "unknown",
# not a gap in our data — see resolve_business_fields).
_ROLE_PREFIXES_WITH_TRAILING_CODE = {
    "cancellation": "Cancellation",
}

# Per-account-stage defaults for Type for RERA IDW / TCP Head on
# Vendor/Contractor/Imprest/site-Salary payments, confirmed from the
# reference sheet. "AH-IDW" (the Aravali Heights project's IDW-stage
# account, 2457) uses the same defaults as "IDW" — same structural role
# for a different project. A Free-stage account doesn't need an entry
# here — it's handled by the fixed HO-Admin defaults instead (see
# resolve_business_fields), since the reference sheet confirms Free-stage
# Vendor/Contractor/Imprest all resolve the same way as Professional/
# Salary HO. Any other stage not listed here falls back to "?" rather
# than guessing.
STAGE_VENDOR_DEFAULTS: dict[str, dict[str, str]] = {
    "IDW": {"type_rera_idw": "Dev- Apt", "tcp_head": "IDW Civil Works"},
    "AH-IDW": {"type_rera_idw": "Dev- Apt", "tcp_head": "IDW Civil Works"},
}

# Professional and HO-stage Salary payments are always tagged Business
# Unit "HO" / Type "HO - Admin" / TCP "Other- Administrative Expenses" in
# the reference sheet, regardless of which account they're on — unlike
# Vendor/Contractor/Imprest, which use the account's own project/stage.
_HO_ADMIN_DEFAULTS = {
    "business_unit": "HO",
    "type_rera_idw": "HO - Admin",
    "tcp_head": "Other- Administrative Expenses",
}

# Stages structurally equivalent to "site" for salary purposes — an
# employee paid from one of these accounts gets Head "Salary Site"
# instead of "Salary HO", confirmed on the reference sheet's IDW account.
_SITE_SALARY_STAGES = {"IDW", "AH-IDW"}

# Known recurring professional/CA firms — these transactions' descriptions
# don't spell out a role keyword the way Vendor/Contractor payments do, so
# they're identified by company name instead. Add more firms here as
# they're confirmed; anything not listed stays "?" rather than a guess.
# Matched with spaces removed (see _extract_role_from_description's
# docstring for why — the same PDF mid-word wrap issue applies here too).
KNOWN_PROFESSIONAL_FIRMS = ["NARESH K JAIN"]

# Known individuals who are always Contractors on DPL accounts.  If their
# name appears in the description and no explicit role keyword is present,
# Head defaults to "Contractor" + the account-stage-appropriate Type/TCP.
# Salary still takes priority (April descriptions spell "SALARY" out
# explicitly), so RAM KISHAN's April rows classify correctly as Salary
# without any date check here.
KNOWN_CONTRACTORS = ["RAM KISHAN", "SHER SINGH"]

# Keywords that identify statutory dues payments (PF/ESI/TDS).
# These always resolve to HO-Admin regardless of account type — confirmed
# by accounts team: even site-staff PF/ESI is routed through Free/HO and
# expensed as Admin, never capitalized as IDW.
_STATUTORY_KEYWORDS = [
    "PROVIDENT FUND", "EPF", "ESIC", " PF-", "-PF-", "/PF",
    "E.S.I", " ESI-", "-ESI-", "/ESI",
    " TDS-", "-TDS-", "/TDS", "TAX DEDUCTED",
    "PTAX", "PROFESSIONAL TAX",
]
# Bare 2-3 letter abbreviations (PF/ESI/TDS) need a real word boundary on
# BOTH sides, not just a trailing space like the entries above - "ESI "
# (letters + trailing space only) false-matched inside "ARAVALI HEIGHT
# RESI DENT WALFARE", a PDF word-wrap artifact splitting "RESIDENT" into
# "RESI DENT": the "ESI " landed right inside "R[ESI ]DENT" with nothing
# checking that the character before it wasn't a letter. Matched via
# regex instead, in _mentions_statutory() below.
_STATUTORY_WORD_BOUNDARY_KEYWORDS = ["PF", "ESI", "TDS"]

# Keywords that identify bank service charges (locker fees, POS charges, etc.)
# HEAD = "Bank Charges", TCP = "Other- Others"
_BANK_CHARGE_KEYWORDS = [
    "SBOX", "SAFE BOX", "LOCKER",
    "POS GST", "BANK CHARGE", "SERVICE CHARGE", "ANNUAL FEE",
    "PROCESSING FEE", "CHGS", "CHRGS", "SCREF",
    "AMB CHARGES",  # covers MISC.CR AMB-charge reversal credits too — this
                    # rule has no credit/debit restriction, so no separate
                    # reversal-specific branch is needed.
]

# Keywords that identify marketing / advertising payments.
# Always resolves to Head "HO - Advert/Mkt", TCP "Other-Selling Expenses"
# confirmed from the VJ rulebook analysis.
_MARKETING_KEYWORDS = [
    "MARKETING", "ADVERTISMENT", "ADVERTISEMENT", "ADVERTISING",
    "-MKT-", "/MKT", "PUBLICITY", "BRANDING", "HOARDING",
]

# Account-specific Business Unit fallbacks — applied only when the Supabase
# DB has no business_unit set for that account. Keyed by last-4 digits.
_ACCOUNT_BU_OVERRIDES: dict[str, str] = {
    "0264": "Casa Romana",
    "0377": "Casa Romana",
    "0490": "Casa Romana",
    "2477": "Casa Romana",       # Casa Romana Free account
    "2314": "Aravali Heights",   # AH account (Master/RERA)
    "2457": "Aravali Heights",   # AH-IDW account
}

# Account-specific stage fallbacks, same semantics as above.
_ACCOUNT_STAGE_OVERRIDES: dict[str, str] = {
    "0264": "Free",
    "0377": "RERA",
    "0490": "IDW",
    "2477": "Free",
    "2457": "AH-IDW",   # Aravali Heights IDW account
}

# Last-4-digit account suffixes that always use "Salary Site" even when
# the DB stage is not yet configured.
_SITE_SALARY_ACCOUNT_SUFFIXES: set[str] = {"0490"}

# Any Bank of Maharashtra IFSC starts with "MAHB" — all transfers matching
# this pattern are treated as Internal per accounts team instruction.
# (The old list of two specific IFSCs is kept only as a fallback reference;
#  the live detection now matches the full MAHB pattern — see
#  _find_bom_internal_ifsc.)
KNOWN_INTERNAL_EXTERNAL_IFSC = ["MAHB0001461", "MAHB0001347"]


_SEGMENT_SPLIT_RE = re.compile(r"[-/]")


def _split_role_segments(description: str) -> list[str]:
    """Split a description into candidate role segments on BOTH "-" and
    "/" delimiters. Most descriptions spell the role out as a
    hyphen-separated segment (e.g. "-Vendor-"), but the IMPS/NA-style
    format instead uses slash-separated segments (e.g.
    ".../MUKESH KUMAR/CONTRACTOR") — checking both means the role isn't
    missed just because a particular bank's layout uses a different
    delimiter for the same information."""
    return _SEGMENT_SPLIT_RE.split(description)


def _extract_role_from_description(description: str) -> Optional[str]:
    """Return the Head implied by an explicit role segment in the
    description (e.g. "-Vendor-", "-contractor-", or "/CONTRACTOR" in
    the IMPS/NA slash-delimited format), or a known recurring
    professional firm's name appearing anywhere in it. Returns None if
    neither is present.

    Role-segment matching also tries with internal spaces removed (e.g.
    "VEN DOR"), since PDF extraction sometimes wraps a cell's text across
    two lines mid-word, splitting a single role word like "VENDOR" into
    two fragments separated by a stray space.
    """
    for segment in _split_role_segments(description):
        normalized = segment.strip().lower()
        normalized_nospace = normalized.replace(" ", "")

        head = DESCRIPTION_ROLE_TO_HEAD.get(normalized) or DESCRIPTION_ROLE_TO_HEAD.get(normalized_nospace)
        if head:
            return head

        # Also catch "NAME ROLE" format where role is the last word of a
        # multi-word segment (e.g. "YOGESH SINGH IMPREST", "MUKESH KUMAR VENDOR").
        words = normalized.split()
        if len(words) > 1:
            last_word_head = DESCRIPTION_ROLE_TO_HEAD.get(words[-1])
            if last_word_head:
                return last_word_head

        for prefix, prefix_head in _ROLE_PREFIXES_WITH_TRAILING_CODE.items():
            if normalized_nospace.startswith(prefix):
                return prefix_head

    normalized_description = description.replace(" ", "").upper()
    for firm_name in KNOWN_PROFESSIONAL_FIRMS:
        if firm_name.replace(" ", "") in normalized_description:
            return "Professional"
    for person_name in KNOWN_CONTRACTORS:
        if person_name.replace(" ", "") in normalized_description:
            return "Contractor"

    return None


def _mentions_imprest(description: str) -> bool:
    """Return True if the description explicitly marks this as an imprest (cash advance).

    Imprest is a transaction type, not a person's identity. It must be
    detected before the Beneficiary Master lookup so that a person who
    normally appears as 'Salary Site' is not misclassified when they
    receive an imprest advance.
    """
    for segment in _split_role_segments(description):
        normalized = segment.strip().lower().replace(" ", "")
        if normalized == "imprest":
            return True
        words = segment.strip().lower().split()
        if words and words[-1] == "imprest":
            return True
    return False


# Below this length, a keyword's stripped form is too short to safely
# fallback-match against a whitespace-stripped description — short
# keywords like "PF"/"ESI"/"TDS" deliberately rely on a surrounding space
# ("PF ", "ESI ", " TDS-") as a word-boundary anchor specifically so they
# don't match as a bare substring of an unrelated word. Stripping spaces
# from those destroys the exact protection they exist for. Longer
# keywords ("CHRGS", "AMB CHARGES" -> "AMBCHARGES") are specific enough
# that this risk is negligible.
_MIN_KEYWORD_LEN_FOR_WHITESPACE_FALLBACK = 5


def _keyword_in_description(description: str, keywords: list[str]) -> bool:
    """Substring-match keywords against a description, tolerant of a stray
    space PDF extraction sometimes inserts mid-word (e.g. "CH RGS" instead
    of "CHRGS", first seen in _mentions_bank_charges()). Checks the
    description as-is first (cheap, the common case), then a
    whitespace-stripped copy against whitespace-stripped keywords — this
    only ever adds matches versus a plain substring check, never removes
    one, since a whitespace-tolerant match is a superset of an exact one.

    Short, boundary-anchored keywords (see
    _MIN_KEYWORD_LEN_FOR_WHITESPACE_FALLBACK) are excluded from the
    whitespace-stripped fallback so this doesn't turn "ESI " into a bare
    "ESI" substring match against unrelated text.
    """
    upper = description.upper()
    if any(k in upper for k in keywords):
        return True
    upper_nospace = upper.replace(" ", "")
    return any(
        k.replace(" ", "") in upper_nospace
        for k in keywords
        if len(k.replace(" ", "")) >= _MIN_KEYWORD_LEN_FOR_WHITESPACE_FALLBACK
    )


def _mentions_statutory(description: str) -> bool:
    """Return True if description indicates a statutory dues payment (PF/ESI/TDS)."""
    if _keyword_in_description(description, _STATUTORY_KEYWORDS):
        return True
    upper = description.upper()
    return any(
        re.search(rf"(?<![A-Z]){kw}(?![A-Z])", upper)
        for kw in _STATUTORY_WORD_BOUNDARY_KEYWORDS
    )


def _mentions_bank_charges(description: str) -> bool:
    """Return True if description indicates a bank service charge (locker, POS fee, etc.).

    Also matches a description that is *just* "GST" — banks post GST on
    monthly service/AMB charges as its own line with no other text, so an
    exact-match (not substring) check is used here rather than adding "GST"
    to _BANK_CHARGE_KEYWORDS, which would wrongly catch unrelated
    government GST payments anywhere GST appears in a description.
    """
    if description.strip().upper() == "GST":
        return True
    return _keyword_in_description(description, _BANK_CHARGE_KEYWORDS)


def _mentions_marketing(description: str) -> bool:
    """Return True if description indicates a marketing/advertising payment."""
    return _keyword_in_description(description, _MARKETING_KEYWORDS)


def _mentions_salary(description: str) -> bool:
    """"salary" is handled separately from DESCRIPTION_ROLE_TO_HEAD since
    its resulting Head name ("Salary HO" vs "Salary Site") depends on the
    account's own stage, not just the description."""
    for segment in _split_role_segments(description):
        normalized = segment.strip().lower().replace(" ", "")
        if normalized == "salary":
            return True
        # Also catch "FIRSTNAME LASTNAME SALARY" — salary as the last word
        # of a multi-word segment (e.g. "BHARAT SINGH SALARY").
        words = segment.strip().lower().split()
        if words and words[-1] == "salary":
            return True
    return False


def _is_site_salary_account(own_stage: Optional[str], account_number: str) -> bool:
    """Return True if salary on this account should be 'Salary Site'."""
    if own_stage in _SITE_SALARY_STAGES:
        return True
    return any(account_number.endswith(s) for s in _SITE_SALARY_ACCOUNT_SUFFIXES)


def _resolve_salary_head(own_stage: Optional[str], account_number: str = "") -> dict[str, Any]:
    """"Salary Site" on IDW-stage accounts and account 0490 (Rule 10),
    "Salary HO" everywhere else."""
    if _is_site_salary_account(own_stage, account_number):
        return {"head": "Salary Site", "is_ho": False}
    return {"head": "Salary HO", "is_ho": True}


def _find_bom_internal_ifsc(description: str) -> Optional[str]:
    """Return a BOM identifier only if the description indicates a
    transfer to/from one of DPL's OWN Bank of Maharashtra accounts — i.e.
    both (a) a MAHB IFSC / "BANK OF MAHARASHTRA" text match, AND (b) the
    description also names one of DPL's own related companies
    (OWN_COMPANY_KEYWORDS, e.g. "DWARKADHIS"). IFSC/bank-name text alone
    is NOT enough: Bank of Maharashtra also issues accounts to ordinary
    external customers and individual employees, and matching on IFSC
    alone previously caused two confirmed real-data bugs (cross-checked
    against the accounts team's own reference sheet):

      1. An incoming "NEFT CR-MAHB0001461-DWARKADHIS PROJECTS..." credit
         — a genuine internal transfer from DPL's own related company —
         used to be skipped entirely by an early "incoming-payment
         prefixes are always external" bailout, causing it to fall
         through to the Collection rule instead of Internal (20 rows
         found wrong).
      2. An outgoing Imprest payment to an individual employee
         ("...MUKESH KUMAR-MAHB0001461-IMPREST-...") used to match this
         function purely on the shared MAHB IFSC and get forced to
         Internal, overriding the correct, explicit "-IMPREST-" keyword
         (1 row found wrong) — because that same IFSC is shared by both
         DPL's real internal account and unrelated individuals' personal
         BOM-affiliated accounts.

    Requiring the company-name match fixes both directions: an incoming
    Dwarkadhis credit now matches regardless of its "NEFT CR-" prefix,
    while a BOM-IFSC transaction to/from anyone else (no company name
    present) now correctly does NOT match, falling through to whichever
    rule actually applies (Collection / Imprest / Beneficiary Master /
    Salary / role keyword).
    """
    normalized = description.replace(" ", "").upper()
    has_bom_ifsc = bool(re.search(r'MAHB[A-Z0-9]{7}', normalized)) or (
        "BANKOFMAHARASHTRA" in normalized or "MAHARASHTRABANK" in normalized
    )
    if not has_bom_ifsc:
        return None
    if not any(keyword in normalized for keyword in OWN_COMPANY_KEYWORDS):
        return None
    mahb_match = re.search(r'MAHB[A-Z0-9]{7}', normalized)
    return mahb_match.group() if mahb_match else "BOM"


def _resolve_bom_internal_transfer(own_stage: Optional[str]) -> dict[str, str]:
    """All BOM/MAHB transfers are Internal per accounts team instruction
    (Rule 8) — Type for RERA IDW and TCP Head are both 'Internal'/
    'Internal transfer' regardless of account stage."""
    return {"type_rera_idw": "Internal", "tcp_head": "Internal transfer"}


# ---------------------------------------------------------------------------
# Beneficiary Master lookup (Phase 2)
# ---------------------------------------------------------------------------

_BENEFICIARY_MASTER_ROLE_SUFFIX = re.compile(
    r"\s+(IMPREST|SALARY|CONTRACTOR|PROFESSIONAL|VENDOR|ADVANCE|REFUND)$",
    re.IGNORECASE,
)
_BENEFICIARY_MASTER_TAB_NAME = "Beneficiary Master"
_BENEFICIARY_MASTER_STATUS_CONFIRMED = "Confirmed"
_beneficiary_cache: Optional[dict[str, str]] = None
_beneficiary_conflict_cache: Optional[dict[str, list[str]]] = None
_beneficiary_secondary_heads_cache: Optional[dict[str, list[str]]] = None


def _load_beneficiary_cache(spreadsheet: Optional[gspread.Spreadsheet]) -> dict[str, str]:
    """Load the Beneficiary Master tab live from Google Sheets, once per
    process, keeping only STATUS="Confirmed" rows.

    This used to read a local beneficiary_master_cache.json snapshot that
    had to be manually regenerated by re-running
    scripts/debug/build_beneficiary_master.py — meaning an accounts-team
    member ticking a row to Confirmed in the sheet had no effect on
    classification until someone technical re-ran that script. Reading
    the sheet directly here means a Confirmed edit takes effect on the
    very next transaction processed, no manual step required.

    Pending/Conflict rows (or any status other than Confirmed) are
    excluded — Rule 6 applies this dict ahead of every other
    classification rule, so an unreviewed or contradictory entry must
    never reach it. A row with no STATUS value at all (pre-dates the
    STATUS column) is treated as Confirmed, matching
    rag_classifier.py's _ensure_status_column() migration, which
    back-filled every pre-existing row that way.

    While it's here, this also builds _beneficiary_conflict_cache — a
    name -> [distinct heads] map for every name that has a
    STATUS="Conflict" row, so Rule 6 can surface the conflict explicitly
    (see _lookup_beneficiary_conflict()) instead of just silently not
    matching. Built in the same pass so this stays a single sheet read.

    Also builds _beneficiary_secondary_heads_cache — a name -> [Head 2,
    Head 3, ...] map of any additional non-blank heads recorded on a
    Confirmed row (see _lookup_beneficiary_secondary_heads()). HEAD
    itself remains the only value classification ever uses (Rule 6
    below) - this is purely so the REASON column can flag "this
    beneficiary has more than one head on file, Head 1 was used by
    priority" for the accounts team to double-check, without changing
    which head actually gets applied.

    Returns an empty dict (Rule 6 simply won't match anything, every
    other rule still runs normally) if the tab is missing or Sheets is
    briefly unreachable — a lookup failure must never block
    classification.
    """
    global _beneficiary_cache, _beneficiary_conflict_cache, _beneficiary_secondary_heads_cache
    if _beneficiary_cache is not None:
        return _beneficiary_cache

    _beneficiary_cache = {}
    _beneficiary_conflict_cache = {}
    _beneficiary_secondary_heads_cache = {}
    if spreadsheet is None:
        return _beneficiary_cache

    try:
        ws = spreadsheet.worksheet(_BENEFICIARY_MASTER_TAB_NAME)
        rows = ws.get_all_values()
    except Exception as exc:
        log.warning("Could not load Beneficiary Master tab (%s) — Rule 6 will match nothing this run.", exc)
        return _beneficiary_cache

    if len(rows) < 2:
        return _beneficiary_cache

    hdr = rows[0]
    if "BENEFICIARY NAME" not in hdr or "Head 1" not in hdr:
        return _beneficiary_cache
    ni = hdr.index("BENEFICIARY NAME")
    hi = hdr.index("Head 1")
    si = hdr.index("STATUS") if "STATUS" in hdr else None
    h2i = hdr.index("Head 2") if "Head 2" in hdr else None
    h3i = hdr.index("Head 3") if "Head 3" in hdr else None

    conflict_heads_by_name: dict[str, set[str]] = {}
    for row in rows[1:]:
        if len(row) <= max(ni, hi):
            continue
        status = row[si].strip() if si is not None and len(row) > si else _BENEFICIARY_MASTER_STATUS_CONFIRMED
        name = row[ni].strip().upper()
        head = row[hi].strip()
        if status == _BENEFICIARY_MASTER_STATUS_CONFLICT:
            if name and head:
                conflict_heads_by_name.setdefault(name, set()).add(head)
            continue
        if status and status != _BENEFICIARY_MASTER_STATUS_CONFIRMED:
            continue
        if name and head:
            _beneficiary_cache[name] = head
            secondary = [
                row[i].strip()
                for i in (h2i, h3i)
                if i is not None and len(row) > i and row[i].strip()
            ]
            if secondary:
                _beneficiary_secondary_heads_cache[name] = secondary

    for name, heads in conflict_heads_by_name.items():
        _beneficiary_conflict_cache[name] = sorted(heads)

    return _beneficiary_cache


def _lookup_beneficiary_conflict(
    description: str,
    spreadsheet: Optional[gspread.Spreadsheet],
) -> Optional[list[str]]:
    """Return the list of conflicting Head values recorded for the
    beneficiary named in this description, or None if that name has no
    STATUS="Conflict" rows in the Beneficiary Master (the normal case)."""
    name = _extract_beneficiary_name(description)
    if not name:
        return None
    _load_beneficiary_cache(spreadsheet)  # populates _beneficiary_conflict_cache too
    return _beneficiary_conflict_cache.get(name) if _beneficiary_conflict_cache else None


def _lookup_beneficiary_secondary_heads(
    description: str,
    spreadsheet: Optional[gspread.Spreadsheet],
) -> Optional[list[str]]:
    """Return any additional (Head 2/Head 3) heads recorded for this
    beneficiary's Confirmed Beneficiary Master row, or None if there are
    none / the name can't be extracted. HEAD (Head 1) is still what Rule 6
    applies — this is only used to flag the row's REASON text for review."""
    name = _extract_beneficiary_name(description)
    if not name:
        return None
    _load_beneficiary_cache(spreadsheet)  # populates _beneficiary_secondary_heads_cache too
    return _beneficiary_secondary_heads_cache.get(name) if _beneficiary_secondary_heads_cache else None


def _extract_beneficiary_name(description: str) -> Optional[str]:
    """Extract the beneficiary name from a NEFT or IMPS description.
    Returns None for CHQ DEP, internal transfers, and other formats
    where a name cannot be reliably parsed."""
    upper = description.upper()
    if upper.startswith("YIB-NEFT") or upper.startswith("YIB-TPT"):
        parts = description.split("-")
        if len(parts) >= 4:
            name = parts[3].strip().upper()
            if name and not re.match(r"^[A-Z]{4}0[A-Z0-9]{6}$", name) and not name.isdigit():
                return _BENEFICIARY_MASTER_ROLE_SUFFIX.sub("", name).strip() or None
    elif upper.startswith("IMPS/"):
        parts = description.split("/")
        if len(parts) >= 3:
            name = parts[-2].strip().upper()
            if name and not name.startswith("RRN") and not name.isdigit():
                return _BENEFICIARY_MASTER_ROLE_SUFFIX.sub("", name).strip() or None
    return None


def _lookup_beneficiary_master(
    description: str,
    spreadsheet: Optional[gspread.Spreadsheet],
) -> Optional[str]:
    """Return the HEAD from the Beneficiary Master for the beneficiary named
    in this description, or None if not found / description format not
    supported."""
    name = _extract_beneficiary_name(description)
    if not name:
        return None
    return _load_beneficiary_cache(spreadsheet).get(name)


# Heads with no single named beneficiary — never worth adding to the master
# (matches scripts/debug/build_beneficiary_master.py's SKIP_HEADS).
_BENEFICIARY_MASTER_SKIP_HEADS = {
    "Internal", "Collection", "Cancellation", "?", "", "HO-Admin",
}
_BENEFICIARY_MASTER_STATUS_PENDING = "Pending"
_BENEFICIARY_MASTER_STATUS_CONFLICT = "Conflict"


def _update_beneficiary_master(
    spreadsheet: Optional[gspread.Spreadsheet],
    discovered: dict[tuple[str, str], int],
) -> None:
    """Add any newly-discovered (name, head) pairs to the Beneficiary
    Master tab, as STATUS="Pending" — never "Confirmed" — so a rule-based
    classification (which can itself be wrong, e.g. a keyword typed
    incorrectly by bank staff) can't silently become a trusted lookup
    for that name's future transactions without a human reviewing it
    first. Mirrors scripts/debug/build_beneficiary_master.py's
    build_master(), but scoped to only the pairs this run just resolved
    (cheap — no full-sheet rescan) rather than a periodic manual rebuild.

    If a name is already recorded under a *different* head, both the new
    and the existing row(s) are flagged STATUS="Conflict" instead of
    silently adding a second, contradictory entry (see the earlier
    YOGESH SING H incident this was built to prevent).
    """
    if not discovered or spreadsheet is None:
        return

    try:
        ws = spreadsheet.worksheet(_BENEFICIARY_MASTER_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        return

    rows = ws.get_all_values()
    if not rows:
        return
    hdr = rows[0]
    if "BENEFICIARY NAME" not in hdr or "Head 1" not in hdr:
        return
    ni, hi = hdr.index("BENEFICIARY NAME"), hdr.index("Head 1")
    si = hdr.index("STATUS") if "STATUS" in hdr else None

    existing_keys: set[tuple[str, str]] = set()
    existing_by_name: dict[str, list[tuple[int, str, str]]] = {}
    for i, row in enumerate(rows[1:], start=2):
        if len(row) <= max(ni, hi):
            continue
        name, head = row[ni].strip().upper(), row[hi].strip()
        status = row[si].strip() if si is not None and len(row) > si else _BENEFICIARY_MASTER_STATUS_CONFIRMED
        existing_keys.add((name, head))
        existing_by_name.setdefault(name, []).append((i, head, status))

    import datetime
    today = datetime.date.today().strftime("%d-%b-%Y")

    new_rows = []
    rows_to_flag_conflict: list[int] = []
    for (name, head), count in sorted(discovered.items()):
        if (name, head) in existing_keys:
            continue
        conflicting = [
            (row_num, status)
            for row_num, other_head, status in existing_by_name.get(name, [])
            if other_head != head
        ]
        if conflicting:
            status = _BENEFICIARY_MASTER_STATUS_CONFLICT
            for row_num, other_status in conflicting:
                if other_status != _BENEFICIARY_MASTER_STATUS_CONFLICT:
                    rows_to_flag_conflict.append(row_num)
        else:
            status = _BENEFICIARY_MASTER_STATUS_PENDING
        notes = f"Auto-extracted ({count} txn)" if count > 1 else "Auto-extracted"

        # Build the row by header-name position rather than a fixed column
        # order — the sheet may have columns (e.g. "Head 2"/"Head 3") that
        # don't exist in this function's own field list, and a positional
        # list silently shifts every value left of them into the wrong
        # column. New auto-discovered rows only ever have a single head at
        # creation time, so any column not set below (Head 2/Head 3, etc.)
        # is simply left blank.
        row_values = {"BENEFICIARY NAME": name, "Head 1": head, "NOTES": notes,
                      "ADDED BY": "System (Rules)", "DATE ADDED": today}
        if si is not None:
            row_values["STATUS"] = status
        new_row = [""] * len(hdr)
        for col_name, value in row_values.items():
            if col_name in hdr:
                new_row[hdr.index(col_name)] = value
        new_rows.append(new_row)

    if rows_to_flag_conflict and si is not None:
        status_col_letter = rowcol_to_a1(1, si + 1).rstrip("0123456789")
        ws.batch_update([
            {"range": f"{status_col_letter}{row_num}", "values": [[_BENEFICIARY_MASTER_STATUS_CONFLICT]]}
            for row_num in rows_to_flag_conflict
        ])

    if new_rows:
        ws.append_rows(new_rows, value_input_option="RAW")
        log.info(
            "Beneficiary Master: added %d new Pending/Conflict row(s) from this run.",
            len(new_rows),
        )


_accounts_by_number_cache: Optional[dict[str, dict[str, Any]]] = None


def _get_accounts_by_number() -> dict[str, dict[str, Any]]:
    """Load account_credentials once per process, keyed by account_number."""
    global _accounts_by_number_cache
    if _accounts_by_number_cache is None:
        accounts = credentials_store.list_credentials(_RECORDS_FALLBACK_PATH)
        _accounts_by_number_cache = {
            acc["account_number"]: acc for acc in accounts if acc.get("account_number")
        }
    return _accounts_by_number_cache


def _find_counterparty_account(description: str, own_account_number: str) -> Optional[dict[str, Any]]:
    """If description mentions one of our OTHER tracked account numbers,
    return that account's record — this reliably signals an internal
    transfer between two of our own accounts (the account number is a
    much stronger signal than company-name matching, since our own
    company's name also legitimately appears in ordinary customer-payment
    descriptions as the beneficiary).

    Compares with whitespace stripped from the description, since PDF
    extraction sometimes inserts a stray space in the middle of an
    account number (e.g. "0455632 00000264").
    """
    normalized_description = description.replace(" ", "")
    for account_number, account in _get_accounts_by_number().items():
        if account_number != own_account_number and account_number in normalized_description:
            return account
    return None


def _looks_like_incoming_payment(description: str) -> bool:
    upper = description.strip().upper()
    return upper.startswith(_INCOMING_PAYMENT_PREFIXES)


def _resolve_own_business_unit(account_number: str, own_account: dict[str, Any]) -> str:
    """This account's own Business Unit: DB value if set, otherwise the
    account-specific override table (0264 and 0490 are always Casa Romana)."""
    raw_bu = own_account.get("business_unit")
    if raw_bu:
        return raw_bu
    return next(
        (bu for sfx, bu in _ACCOUNT_BU_OVERRIDES.items() if account_number.endswith(sfx)),
        UNKNOWN_MAPPING_VALUE,
    )


def resolve_business_fields(
    account_number: str,
    description: str,
    deposits: float,
    withdrawals: float,
    spreadsheet: Optional[gspread.Spreadsheet] = None,
) -> dict[str, Any]:
    """Thin wrapper around _resolve_business_fields that tags the result
    with this account's own Business Unit/stage, so _explain_resolved_field()
    can recognize "this value equals the account's own BU/stage default"
    without every one of the 12 rule branches below needing to repeat that
    bookkeeping in its return statement.
    """
    result = _resolve_business_fields(account_number, description, deposits, withdrawals, spreadsheet)
    accounts = _get_accounts_by_number()
    own_account = accounts.get(account_number, {})
    result["_own_business_unit"] = _resolve_own_business_unit(account_number, own_account)
    result["_own_stage"] = own_account.get("account_stage") or next(
        (s for sfx, s in _ACCOUNT_STAGE_OVERRIDES.items() if account_number.endswith(sfx)),
        None,
    )
    return result


def _resolve_business_fields(
    account_number: str,
    description: str,
    deposits: float,
    withdrawals: float,
    spreadsheet: Optional[gspread.Spreadsheet] = None,
) -> dict[str, Any]:
    """Determine Head/Business Unit/Type for RERA IDW/TCP Head using the
    most reliable, generalizable rules confirmed from 2 years of the
    accounts team's own reference sheet:

      1. Internal transfer between two of our own tracked accounts
         (detected via a counterparty account number appearing in the
         description) -> Head "Internal", Business Unit = this account's
         own project, and a Type for RERA IDW label looked up by (this
         account's stage, counterparty's stage): a confirmed stage-pair
         label when known, the literal "Internal" label when the pair
         isn't part of the Master/Free/RERA/IDW pipeline (e.g. any
         transfer involving the Aravali Heights accounts — confirmed
         consistently labeled this way), or "?" only for the one pair
         (RERA<->IDW) confirmed to be genuinely ambiguous in the source
         data itself.
      2. A transfer to/from one of DPL's own external (non-YES-BANK)
         accounts, identified by a known IFSC + "Dwarkadhis" beneficiary
         name -> Head "Internal", with Type for RERA IDW/TCP Head
         resolved per this account's own stage (see
         _resolve_bom_internal_transfer) — confidently known for
         Master/Free stages, "?" where the source data itself is
         contradictory (IDW stage) or unobserved.
      3. The description spells the role out directly (e.g.
         "-Vendor-"/"-contractor-"/"-imprest-"), or is a salary payment
         -> Head = that role (Salary further splits into "Salary HO"/
         "Salary Site" by account stage), with Business Unit/Type for
         RERA IDW/TCP Head from a per-account-stage default table (or
         the fixed HO/HO-Admin defaults for Professional and HO-stage
         Salary) when known for that stage.
      4. An incoming payment (UPI/NEFT/IMPS/RTGS/NET-TPT) that ISN'T an
         internal transfer -> Head "Collection", TCP Head "Credit- no
         effect", Type for RERA IDW "Customer Collection".

    Anything else returns head=None (caller falls back to the existing
    get_head() heuristic) with business_unit/type_rera_idw/tcp_head all
    "?", per the explicit instruction to leave fields blank/unknown
    rather than guess.

    Returns:
        Dict with keys "head" (str or None), "business_unit",
        "type_rera_idw", "tcp_head".
    """
    accounts = _get_accounts_by_number()
    own_account = accounts.get(account_number, {})
    own_business_unit = _resolve_own_business_unit(account_number, own_account)

    # Stage: use DB value if set; otherwise fall back to account-specific override
    # (Rule 1/7 — 0377 = RERA, 0490 = IDW).
    own_stage = own_account.get("account_stage") or next(
        (s for sfx, s in _ACCOUNT_STAGE_OVERRIDES.items() if account_number.endswith(sfx)),
        None,
    )

    reasons: dict[str, str] = {}
    if own_business_unit == UNKNOWN_MAPPING_VALUE:
        reasons["business_unit"] = "this account has no Business Unit configured"

    # ── Rule 1: internal transfer between two of our own tracked accounts ──
    counterparty = _find_counterparty_account(description, account_number)
    if counterparty is not None:
        # Apply same BU/stage overrides to counterparty so matching works
        # even when the counterparty's DB config is incomplete.
        raw_cpty_bu = counterparty.get("business_unit")
        cpty_account_number = counterparty.get("account_number", "")
        counterparty_business_unit = raw_cpty_bu or next(
            (bu for sfx, bu in _ACCOUNT_BU_OVERRIDES.items() if cpty_account_number.endswith(sfx)),
            None,
        )
        counterparty_stage = counterparty.get("account_stage") or next(
            (s for sfx, s in _ACCOUNT_STAGE_OVERRIDES.items() if cpty_account_number.endswith(sfx)),
            None,
        )

        type_rera_idw = "Internal"
        tcp_head = "Internal transfer"

        if (
            own_stage
            and counterparty_stage
            and own_business_unit != UNKNOWN_MAPPING_VALUE
            and own_business_unit == counterparty_business_unit
        ):
            stage_pair = frozenset({own_stage, counterparty_stage})
            if stage_pair in TRANSFER_STAGE_LABELS:
                type_rera_idw = TRANSFER_STAGE_LABELS[stage_pair]
            elif stage_pair in _AMBIGUOUS_STAGE_PAIRS:
                # RERA↔IDW: accounts team uses "RERA IDW New" (confirmed from
                # 0377 sheet Jul 2026). TCP always "Internal transfer".
                type_rera_idw = "RERA IDW New"
                tcp_head = "Internal transfer"

        return {
            "head": "Internal",
            "business_unit": own_business_unit,
            "type_rera_idw": type_rera_idw,
            "tcp_head": tcp_head,
            "confidence": "High",
            "classified_by": "Rule 1: Internal transfer (counterparty account number found in description)",
            "reasons": reasons,
        }

    # ── Rule 2 (Rule 8): BOM / MAHB account — always Internal ───────────────
    bom_ifsc = _find_bom_internal_ifsc(description)
    if bom_ifsc is not None:
        resolved = _resolve_bom_internal_transfer(own_stage)
        return {
            "head": "Internal",
            "business_unit": own_business_unit,
            "type_rera_idw": resolved["type_rera_idw"],
            "tcp_head": resolved["tcp_head"],
            "confidence": "High",
            "classified_by": "Rule 2: Internal transfer (Bank of Maharashtra IFSC detected)",
            "reasons": reasons,
        }

    # ── Rule 3: CHQ DEP / cheque deposit — Collection (incoming) ────────────
    # Moved before master list so incoming cheques are never misclassified
    # by a beneficiary name that happens to appear in the description.
    if deposits > 0:
        desc_nospace = description.upper().replace(" ", "")
        if "CHQDEP" in desc_nospace or "CHEQDEP" in desc_nospace or "BYCLG" in desc_nospace:
            return {
                "head": "Collection",
                "business_unit": own_business_unit,
                "type_rera_idw": "Customer Collection",
                "tcp_head": "Credit- no effect",
                "confidence": "High",
                "classified_by": "Rule 3: Collection (cheque deposit detected)",
                "reasons": reasons,
            }

    # ── Rule 4: incoming payment (UPI / NEFT CR / IMPS / RTGS) — Collection ─
    if deposits > 0 and _looks_like_incoming_payment(description):
        return {
            "head": "Collection",
            "business_unit": own_business_unit,
            "type_rera_idw": "Customer Collection",
            "tcp_head": "Credit- no effect",
            "confidence": "High",
            "classified_by": "Rule 4: Collection (incoming payment prefix — UPI/NEFT/IMPS/RTGS)",
            "reasons": reasons,
        }

    # ── Rule 5: Imprest — must run BEFORE master list ────────────────────────
    # Imprest is a transaction TYPE (petty-cash advance), not a person's
    # identity. A person can appear in the master as "Salary Site" but receive
    # an imprest advance in the same month. The master would wrongly return
    # "Salary Site" in that case, so we check the IMPREST keyword first.
    if _mentions_imprest(description):
        if own_stage == "Free":
            return {
                "head": "Imprest",
                "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
                "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
                "tcp_head": _HO_ADMIN_DEFAULTS["tcp_head"],
                "confidence": "Low",
                "classified_by": "Rule 5: Imprest (keyword in description — verify staff typed correct remark)",
                "reasons": {},
            }
        defaults = STAGE_VENDOR_DEFAULTS.get(own_stage, {})
        return {
            "head": "Imprest",
            "business_unit": own_business_unit,
            "type_rera_idw": defaults.get("type_rera_idw", UNKNOWN_MAPPING_VALUE),
            "tcp_head": defaults.get("tcp_head", UNKNOWN_MAPPING_VALUE),
            "confidence": "Low",
            "classified_by": "Rule 5: Imprest (keyword in description — verify staff typed correct remark)",
            "reasons": reasons,
        }

    # ── Rule 6a: Beneficiary Master conflict — surface it, don't guess ──────
    # If this name has two (or more) different heads flagged STATUS="Conflict"
    # in the Master (see _update_beneficiary_master()), the identity is
    # genuinely unresolved — falling through to a keyword rule or the AI
    # classifier would silently paper over that with a guess. Show "?" and
    # name both candidates so the accounts team resolves it in the Master
    # itself, rather than the pipeline picking one.
    conflict_heads = _lookup_beneficiary_conflict(description, spreadsheet)
    if conflict_heads:
        _conflict_name = _extract_beneficiary_name(description) or "beneficiary"
        _conflict_reason = (
            f"'{_conflict_name}' has {len(conflict_heads)} conflicting heads recorded "
            f"in the Beneficiary Master ({', '.join(conflict_heads)}) — resolve which "
            "one is correct in the Beneficiary Master tab"
        )
        return {
            "head": UNKNOWN_MAPPING_VALUE,
            "business_unit": UNKNOWN_MAPPING_VALUE,
            "type_rera_idw": UNKNOWN_MAPPING_VALUE,
            "tcp_head": UNKNOWN_MAPPING_VALUE,
            "confidence": "Low",
            "classified_by": f"Rule 6: Beneficiary Master conflict — {_conflict_reason}",
            "reasons": {
                "business_unit": _conflict_reason,
                "type_rera_idw": _conflict_reason,
                "tcp_head": _conflict_reason,
            },
        }

    # ── Rule 6: Beneficiary Master lookup — runs FIRST among outgoing rules ──
    # Name-based identity check overrides role keywords below.
    # Rationale: keywords in bank descriptions are typed by DPL staff when
    # initiating the payment. A wrong remark (e.g. "CONTRACTOR" typed for a
    # Vendor payment) must not corrupt the classification. The master list
    # records the confirmed identity of each payee, so it takes priority.
    # Incoming payments (Collection) and Imprest are already caught above.
    master_head = _lookup_beneficiary_master(description, spreadsheet)
    if master_head:
        _master_name = _extract_beneficiary_name(description) or "beneficiary"
        # Beneficiary Master stores one fixed Head per name, but "Salary
        # HO" vs "Salary Site" is account-dependent, not person-dependent -
        # a person recorded as "Salary HO" (paid from a Free/Master-stage
        # account normally) can still receive a payment through a site
        # account (e.g. 0490/IDW, or any AH-IDW account), which must show
        # as "Salary Site" there regardless of what's recorded in the
        # Master. Without this, Rule 6 would apply "Salary HO" verbatim on
        # a site account, which the accounts team's reference sheet never
        # does.
        if master_head == "Salary HO" and _is_site_salary_account(own_stage, account_number):
            master_head = "Salary Site"
        _master_reason = f"Rule 6: Beneficiary Master — '{_master_name}' confirmed as {master_head}"
        _secondary_heads = _lookup_beneficiary_secondary_heads(description, spreadsheet)
        if _secondary_heads:
            _master_reason += (
                f" (Beneficiary Master also lists {', '.join(_secondary_heads)} as additional "
                f"head(s) for this beneficiary — Head 1 ('{master_head}') was used by priority; "
                "kindly recheck in the Beneficiary Master tab)"
            )
        if master_head in ("Salary HO", "Professional") or own_stage == "Free":
            return {
                "head": master_head,
                "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
                "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
                "tcp_head": _HO_ADMIN_DEFAULTS["tcp_head"],
                "confidence": "High",
                "classified_by": _master_reason,
                "reasons": {},
            }
        if master_head == "Salary Site":
            defaults = STAGE_VENDOR_DEFAULTS.get(own_stage, {})
            return {
                "head": master_head,
                "business_unit": own_business_unit,
                "type_rera_idw": defaults.get("type_rera_idw", UNKNOWN_MAPPING_VALUE),
                "tcp_head": defaults.get("tcp_head", UNKNOWN_MAPPING_VALUE),
                "confidence": "High",
                "classified_by": _master_reason,
                "reasons": reasons,
            }
        defaults = STAGE_VENDOR_DEFAULTS.get(own_stage, {})
        type_rera_idw = defaults.get("type_rera_idw", UNKNOWN_MAPPING_VALUE)
        tcp_head = defaults.get("tcp_head", UNKNOWN_MAPPING_VALUE)
        return {
            "head": master_head,
            "business_unit": own_business_unit,
            "type_rera_idw": type_rera_idw,
            "tcp_head": tcp_head,
            "confidence": "High",
            "classified_by": _master_reason,
            "reasons": reasons,
        }

    # ── Rule 7: Salary keyword ───────────────────────────────────────────────
    # Runs after master list — if a person is in the master, their identity
    # overrides even a "SALARY" remark. Falls back here only for payees not
    # yet in the master list.
    if _mentions_salary(description):
        salary = _resolve_salary_head(own_stage, account_number)
        if salary["is_ho"]:
            return {
                "head": salary["head"],
                "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
                "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
                "tcp_head": _HO_ADMIN_DEFAULTS["tcp_head"],
                "confidence": "Low",
                "classified_by": "Rule 7: Salary HO (SALARY keyword in description — verify staff typed correct remark)",
                "reasons": {},
            }
        defaults = STAGE_VENDOR_DEFAULTS.get(own_stage, {})
        type_rera_idw = defaults.get("type_rera_idw", UNKNOWN_MAPPING_VALUE)
        tcp_head = defaults.get("tcp_head", UNKNOWN_MAPPING_VALUE)
        if type_rera_idw == UNKNOWN_MAPPING_VALUE or tcp_head == UNKNOWN_MAPPING_VALUE:
            reasons["type_rera_idw"] = reasons["tcp_head"] = (
                "no historical data for Salary Site payments from this account's stage"
            )
        return {
            "head": salary["head"],
            "business_unit": own_business_unit,
            "type_rera_idw": type_rera_idw,
            "tcp_head": tcp_head,
            "confidence": "Low",
            "classified_by": "Rule 7: Salary Site (SALARY keyword in description — verify staff typed correct remark)",
            "reasons": reasons,
        }

    # ── Rule 8: Statutory Dues (PF / ESI / TDS) ─────────────────────────────
    if _mentions_statutory(description):
        return {
            "head": "Statutory Dues",
            "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
            "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
            "tcp_head": _HO_ADMIN_DEFAULTS["tcp_head"],
            "confidence": "Low",
            "classified_by": "Rule 8: Statutory Dues (PF/ESI/TDS keyword in description — verify staff typed correct remark)",
            "reasons": {},
        }

    # ── Rule 9: Marketing / Advertising ─────────────────────────────────────
    if _mentions_marketing(description):
        return {
            "head": "HO - Advert/Mkt",
            "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
            "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
            "tcp_head": "Other-Selling Expenses",
            "confidence": "Low",
            "classified_by": "Rule 9: Marketing (MARKETING/ADVERTISEMENT keyword in description — verify staff typed correct remark)",
            "reasons": {},
        }

    # ── Rule 10: Bank Charges (locker fees, POS charges, service fees) ───────
    if _mentions_bank_charges(description):
        return {
            "head": "Bank Charges",
            "business_unit": own_business_unit,
            "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
            "tcp_head": "Other- Others",
            "confidence": "Low",
            "classified_by": "Rule 10: Bank Charges (LOCKER/CHGS/SERVICE CHARGE keyword in description — verify staff typed correct remark)",
            "reasons": {},
        }

    # ── Rule 11: explicit role keyword in description ─────────────────────────
    # Last resort for outgoing payments — only fires if the payee is NOT in
    # the master list. Keywords here (VENDOR / CONTRACTOR / IMPREST /
    # PROFESSIONAL) are what DPL staff typed, which can be wrong. Any payee
    # that repeatedly appears should be added to the Beneficiary Master so
    # future transactions bypass this step entirely.
    role_head = _extract_role_from_description(description)
    if role_head == "Cancellation":
        reasons["tcp_head"] = (
            "not recorded in 2 years of historical data for Cancellation transactions"
        )
        return {
            "head": role_head,
            "business_unit": own_business_unit,
            "type_rera_idw": "Cust Cancellation",
            "tcp_head": UNKNOWN_MAPPING_VALUE,
            "confidence": "Low",
            "classified_by": "Rule 11: Cancellation (role keyword in description — verify staff typed correct remark)",
            "reasons": reasons,
        }

    if role_head:
        if role_head == "Professional" or own_stage == "Free":
            return {
                "head": role_head,
                "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
                "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
                "tcp_head": _HO_ADMIN_DEFAULTS["tcp_head"],
                "confidence": "Low",
                "classified_by": f"Rule 11: {role_head} (role keyword in description — verify staff typed correct remark)",
                "reasons": {},
            }
        defaults = STAGE_VENDOR_DEFAULTS.get(own_stage, {})
        type_rera_idw = defaults.get("type_rera_idw", UNKNOWN_MAPPING_VALUE)
        tcp_head = defaults.get("tcp_head", UNKNOWN_MAPPING_VALUE)
        if type_rera_idw == UNKNOWN_MAPPING_VALUE or tcp_head == UNKNOWN_MAPPING_VALUE:
            reasons["type_rera_idw"] = reasons["tcp_head"] = (
                f"no historical data for {role_head} payments from this account's stage"
            )
        return {
            "head": role_head,
            "business_unit": own_business_unit,
            "type_rera_idw": type_rera_idw,
            "tcp_head": tcp_head,
            "confidence": "Low",
            "classified_by": f"Rule 11: {role_head} (role keyword in description — verify staff typed correct remark)",
            "reasons": reasons,
        }

    # ── Rule 12: get_head() confidently resolves to an internal-type Head ──
    # get_head() can recognize an internal transfer via its own party_master
    # lookup, a keyword like "dwarkadhis"/"for esi", or a description
    # pattern — none of resolve_business_fields()'s rules above needed to
    # match for it to be confident. That used to only surface as a bare
    # Head string from get_head()'s Head-only fallback path in
    # classify_rows(), with Business Unit/Type/TCP left "?" forever since
    # only this function's own rules ever filled those in. Checking
    # is_internal_type_head() — driven by heads_config.json's party_types,
    # not a hardcoded list of specific Head names — means this closes the
    # gap for any current or future Head tagged that way, not just the
    # ones already discovered.
    party_head = get_head(description, deposits, withdrawals)
    if is_internal_type_head(party_head):
        return {
            "head": party_head,
            "business_unit": own_business_unit,
            "type_rera_idw": "Internal",
            "tcp_head": "Internal transfer",
            "confidence": "Low",
            "classified_by": f"Rule 12: Internal transfer (party recognized as {party_head} by name)",
            "reasons": {},
        }

    # ── Fallback ─────────────────────────────────────────────────────────────
    reasons["business_unit"] = reasons["type_rera_idw"] = reasons["tcp_head"] = (
        "description format not recognized by any existing rule"
    )
    return {
        "head": None,
        "business_unit": UNKNOWN_MAPPING_VALUE,
        "type_rera_idw": UNKNOWN_MAPPING_VALUE,
        "tcp_head": UNKNOWN_MAPPING_VALUE,
        "confidence": "Low",
        "classified_by": "No rule matched — sent to RAG AI for classification",
        "reasons": reasons,
    }

LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("classify_transactions")

# Columns this script is responsible for adding/populating.
BUSINESS_UNIT_COLUMN = "BUSINESS UNIT"
HEAD_COLUMN = "HEAD"
TYPE_RERA_IDW_COLUMN = "TYPE FOR RERA IDW"
TCP_HEAD_COLUMN = "TCP Head"
NARRATION_COLUMN = "NARRATION"
CONFIDENCE_COLUMN = "CONFIDENCE"
REASON_COLUMN = "REASON"
APPROVAL_1_COLUMN = "APPROVAL 1"
APPROVAL_2_COLUMN = "APPROVAL 2"
APPROVAL_3_COLUMN = "APPROVAL 3"

# All columns that must be present, in the order they're appended if missing.
CLASSIFICATION_COLUMNS = [
    BUSINESS_UNIT_COLUMN,
    HEAD_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
    NARRATION_COLUMN,
    CONFIDENCE_COLUMN,
    REASON_COLUMN,
    APPROVAL_1_COLUMN,
    APPROVAL_2_COLUMN,
    APPROVAL_3_COLUMN,
]

# Used only to decide whether a row is "already classified" (and can be
# skipped). Excludes REASON, CONFIDENCE, and APPROVAL columns — these are
# supplementary and their absence never means "row not yet classified".
_REQUIRED_NON_BLANK_COLUMNS = [
    BUSINESS_UNIT_COLUMN,
    HEAD_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
    NARRATION_COLUMN,
]

SCRIPT_DIR = Path(__file__).resolve().parent

# Value written for Business Unit/Type for RERA IDW/TCP Head whenever we
# aren't confident enough to fill them in from a known rule — never
# invented/guessed.
UNKNOWN_MAPPING_VALUE = "?"


# ---------------------------------------------------------------------------
# Worksheet access
# ---------------------------------------------------------------------------

def open_account_worksheet(
    client: gspread.Client,
    sheet_id: str,
    worksheet_name: str,
) -> gspread.Worksheet:
    """Open the given account's worksheet/tab.

    Raises:
        gspread.exceptions.WorksheetNotFound: If the worksheet does not exist.
            This script only updates an existing sheet; it never creates one
            (upload_to_sheets.py is responsible for creating account tabs).
    """
    spreadsheet = client.open_by_key(sheet_id)
    return spreadsheet.worksheet(worksheet_name)


# ---------------------------------------------------------------------------
# Header management
# ---------------------------------------------------------------------------

def ensure_classification_columns(worksheet: gspread.Worksheet) -> tuple[list[str], dict[str, int]]:
    """Ensure all classification columns exist in the header row.

    Columns ensured: Head, Narration, Project, Head - Income Tax,
    Type for RERA IDW, TCP Head. Any that are missing are appended to the
    end of the header row (existing columns are never reordered or
    removed).

    Returns:
        A tuple of (updated_header_row, column_indices), where
        column_indices maps each column name in CLASSIFICATION_COLUMNS to
        its 1-based column index.
    """
    header_row = worksheet.row_values(1)

    if not header_row:
        raise ValueError("Worksheet has no header row; cannot classify an empty sheet.")

    for column_name in CLASSIFICATION_COLUMNS:
        if column_name not in header_row:
            header_row.append(column_name)
            log.info("%s column not found — adding it.", column_name)

    worksheet.update(range_name="A1", values=[header_row])

    column_indices = {name: header_row.index(name) + 1 for name in CLASSIFICATION_COLUMNS}

    return header_row, column_indices


# ---------------------------------------------------------------------------
# Row helpers
# ---------------------------------------------------------------------------

def _get_cell(row: list[str], header_row: list[str], column_name: str) -> str:
    """Safely read a cell value by column name, tolerating short/ragged rows."""
    if column_name not in header_row:
        return ""
    index = header_row.index(column_name)
    if index >= len(row):
        return ""
    return row[index].strip()


def _to_float(value: str) -> float:
    """Safely convert a raw sheet cell value to a float, defaulting to 0.0."""
    cleaned = value.replace(",", "").strip()
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _parse_amount(deposits_raw: str, withdrawals_raw: str) -> float:
    """Return Deposits if positive, otherwise Withdrawals, as a float."""
    deposits = _to_float(deposits_raw)
    if deposits > 0:
        return deposits
    return _to_float(withdrawals_raw)


def _is_row_empty(row: list[str]) -> bool:
    """A row is empty if every cell is blank."""
    return all(cell.strip() == "" for cell in row)


def _safe_parse_description(description: str, sheet_row_number: int) -> dict | None:
    """Run parse_description() defensively so a parsing issue never
    blocks classification of a row.

    Returns:
        The parsed-fields dict on success (even if every field inside it
        is None — that just means no known pattern matched), or None if
        parse_description() itself raised. Callers must fall back to the
        raw description in either case; get_head()/generate_narration()
        already operate on the raw description string, so this fallback
        is automatic and no transaction is ever skipped.
    """
    try:
        parsed = parse_description(description)
    except Exception as exc:
        log.warning(
            "Row %d: description_parser raised %s — falling back to raw description.",
            sheet_row_number, exc,
        )
        return None

    if any(value is not None for value in parsed.values()):
        log.debug("Row %d: description parsed successfully: %s", sheet_row_number, parsed)
    else:
        log.debug(
            "Row %d: no known pattern matched description — falling back to raw description.",
            sheet_row_number,
        )

    return parsed


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------

def _explain_resolved_field(
    key: str, value: str, resolved: dict[str, Any],
) -> Optional[str]:
    """Return a plain-language reason a resolved (non-'?') field got its
    specific value, based on which fixed default/table it matches. Returns
    None for a value this generic logic doesn't recognize (e.g. one that
    came from a per-row lookup like the Beneficiary Master or a
    counterparty's own stage) — those cases already explain themselves via
    classified_by, so no generic fallback text is forced onto them.
    """
    own_business_unit = resolved.get("_own_business_unit")

    if key == "business_unit":
        if value == own_business_unit:
            return "this account's own assigned Business Unit/project (a transaction is booked against the Business Unit of the account it moved through, unless the Head itself is a fixed HO-level expense)"
        if value == _HO_ADMIN_DEFAULTS["business_unit"]:
            return "fixed 'HO' default used for HO-level Heads (Salary HO, Professional, Statutory Dues, Marketing, Bank Charges' Type, Imprest at Free-stage accounts) — these are always booked to HO regardless of which account the payment physically came from, per the reference sheet"
        return None

    if key == "type_rera_idw":
        if value == _HO_ADMIN_DEFAULTS["type_rera_idw"]:
            return "fixed 'HO - Admin' default paired with an HO-level Head — the reference sheet always books these as administrative overhead, not project-specific spend"
        if value == "Internal":
            return "fixed label for a transfer between two of DPL's own tracked accounts — no external party is involved, so it can't be Dev/Collection/Admin spend"
        if value == "Customer Collection":
            return "fixed label for money coming IN from an external party (customer/UPI/NEFT credit) — Collections are always tagged this way, never a spend category"
        if value == "Cust Cancellation":
            return "fixed label used specifically for the 'Cancellation' role keyword — a refund/reversal to a customer, distinct from a normal Collection"
        defaults = STAGE_VENDOR_DEFAULTS.get(resolved.get("_own_stage"), {})
        if value == defaults.get("type_rera_idw"):
            return f"this account's stage ({resolved.get('_own_stage')}) default Type for outgoing Vendor/Contractor/Imprest/Salary-Site style payments, per the reference sheet's stage-specific mapping"
        return None

    if key == "tcp_head":
        if value == _HO_ADMIN_DEFAULTS["tcp_head"]:
            return "fixed 'Other- Administrative Expenses' default paired with an HO-level Head"
        if value == "Internal transfer":
            return "fixed label — internal transfers never hit a P&L expense/income TCP head"
        if value == "Credit- no effect":
            return "fixed label for incoming Collections — a customer receipt has no TCP expense effect"
        if value == "Other- Others":
            return "fixed catch-all TCP for Bank Charges — small recurring bank-levied fees (AMB/locker/POS/GST-on-charges) aren't split into a more specific expense line in the reference sheet"
        if value == "Other-Selling Expenses":
            return "fixed TCP for Marketing/Advertising spend, per the reference sheet"
        defaults = STAGE_VENDOR_DEFAULTS.get(resolved.get("_own_stage"), {})
        if value == defaults.get("tcp_head"):
            return f"this account's stage ({resolved.get('_own_stage')}) default TCP for outgoing Vendor/Contractor/Imprest/Salary-Site style payments"
        return None

    return None


def _build_reason_text(display_head: str, resolved: dict[str, Any]) -> str:
    """Build a detailed, field-by-field human-readable reason string for the
    REASON column: which rule fired and why (from classified_by), then a
    specific explanation for Business Unit / Type for RERA IDW / TCP Head —
    either why that exact value was chosen (resolved fields) or why it
    couldn't be determined ('?' fields).
    """
    parts: list[str] = []
    classified_by = resolved.get("classified_by", "")
    reasons = resolved.get("reasons", {})

    if classified_by:
        parts.append(f"HEAD = '{display_head}' — {classified_by}")
    if display_head == UNKNOWN_MAPPING_VALUE and "conflict" not in classified_by.lower():
        parts.append("HEAD could not be determined — check description format or add payee to Beneficiary Master")

    for key, label in (
        ("business_unit", "Business Unit"),
        ("type_rera_idw", "Type for RERA IDW"),
        ("tcp_head", "TCP Head"),
    ):
        value = resolved.get(key)
        if value == UNKNOWN_MAPPING_VALUE:
            parts.append(f"{label} = ? — {reasons.get(key, 'not resolved by any existing rule')}")
        else:
            explanation = reasons.get(key) or _explain_resolved_field(key, value, resolved)
            if explanation:
                parts.append(f"{label} = '{value}' — {explanation}")

    return " | ".join(parts)


def classify_rows(
    worksheet: gspread.Worksheet,
    header_row: list[str],
    column_indices: dict[str, int],
) -> int:
    """Classify each data row and write all classification columns back to
    the sheet: Head, Narration, Project, Head - Income Tax,
    Type for RERA IDW, TCP Head.

    Skips:
      * Fully empty rows
      * Rows without a Description
      * Rows that already have every classification column filled in
        (idempotent — this also lets previously Head/Narration-only rows
        get backfilled with the 4 new columns on the next run, since they
        won't yet have Project/Head - Income Tax/Type for RERA IDW/TCP Head)

    Returns:
        The number of rows updated.
    """
    all_values = worksheet.get_all_values()
    data_rows = all_values[1:]  # exclude header

    updates: list[gspread.cell.Cell] = []
    updated_count = 0
    updated_rows: list[int] = []
    discovered_beneficiaries: dict[tuple[str, str], int] = {}

    for offset, row in enumerate(data_rows):
        sheet_row_number = offset + 2  # +1 for header, +1 for 1-based index

        if _is_row_empty(row):
            continue

        description = _get_cell(row, header_row, "DESCRIPTION")
        if not description:
            log.debug("Skipping row %d: no Description.", sheet_row_number)
            continue

        already_classified = all(
            _get_cell(row, header_row, column_name)
            for column_name in _REQUIRED_NON_BLANK_COLUMNS
        )
        if already_classified:
            log.debug("Skipping row %d: already fully classified.", sheet_row_number)
            continue

        # Structured parsing step (description_parser.py). Parsing is
        # advisory only: get_head()/generate_narration() below still take
        # the raw description string, so a parse failure or an
        # unrecognized description format never skips the row.
        _safe_parse_description(description, sheet_row_number)

        deposits_raw = _get_cell(row, header_row, "CREDITS")
        withdrawals_raw = _get_cell(row, header_row, "DEBITS")
        deposits = _to_float(deposits_raw)
        withdrawals = _to_float(withdrawals_raw)
        amount = _parse_amount(deposits_raw, withdrawals_raw)
        account_number = _get_cell(row, header_row, "Account Number")

        # Try the confident, generalizable business rules first (internal
        # transfer between our own tracked accounts, an incoming customer
        # payment, or a party_master-recognized internal-type company —
        # resolve_business_fields() now covers all of these with
        # Business Unit/Type/TCP filled in together). Falls back to the
        # existing get_head() heuristic — with business_unit/type_rera_idw/
        # tcp_head left as "?" — only for whatever none of those confidently
        # cover.
        resolved = resolve_business_fields(
            account_number, description, deposits, withdrawals,
            spreadsheet=worksheet.spreadsheet,
        )
        head = resolved["head"] or get_head(description, deposits, withdrawals)

        # heads.py's own emergency catch-all ("Others") means it genuinely
        # doesn't know either — show "?" instead, consistent with how
        # Business Unit/Type for RERA IDW/TCP Head already show "?" when
        # unknown, rather than a label that looks like a confirmed answer.
        display_head = UNKNOWN_MAPPING_VALUE if head == "Others" else head

        if display_head not in _BENEFICIARY_MASTER_SKIP_HEADS:
            beneficiary_name = _extract_beneficiary_name(description)
            if beneficiary_name:
                key = (beneficiary_name, display_head)
                discovered_beneficiaries[key] = discovered_beneficiaries.get(key, 0) + 1

        narration = generate_narration(
            description,
            display_head,
            amount,
            business_unit=resolved["business_unit"],
            type_rera_idw=resolved["type_rera_idw"],
            deposits=deposits,
            withdrawals=withdrawals,
            own_account_number=account_number,
        )

        reason_text = _build_reason_text(display_head, resolved)

        row_values = {
            BUSINESS_UNIT_COLUMN: resolved["business_unit"],
            HEAD_COLUMN: display_head,
            TYPE_RERA_IDW_COLUMN: resolved["type_rera_idw"],
            TCP_HEAD_COLUMN: resolved["tcp_head"],
            NARRATION_COLUMN: narration,
            CONFIDENCE_COLUMN: resolved.get("confidence", "Low"),
            REASON_COLUMN: reason_text,
            # Approval columns: only write if currently blank — never overwrite
            # a value the accounts team has already entered.
            **{
                col: ""
                for col in (APPROVAL_1_COLUMN, APPROVAL_2_COLUMN, APPROVAL_3_COLUMN)
                if not _get_cell(row, header_row, col)
            },
        }

        for column_name, value in row_values.items():
            updates.append(
                gspread.cell.Cell(
                    row=sheet_row_number,
                    col=column_indices[column_name],
                    value=value,
                )
            )
        updated_rows.append(sheet_row_number)
        updated_count += 1

    if updates:
        worksheet.update_cells(updates, value_input_option="RAW")
        _mark_rows_unverified(worksheet, updated_rows, column_indices)
        log.info("Updated %d row(s) with full classification.", updated_count)
    else:
        log.info("No rows required classification.")

    _update_beneficiary_master(worksheet.spreadsheet, discovered_beneficiaries)

    return updated_count


# Red text signals "auto-classified, not yet reviewed" to the accounts team —
# they change it to black once they've verified a row.
UNVERIFIED_TEXT_COLOR = {"red": 0.8, "green": 0.0, "blue": 0.0}


def _mark_rows_unverified(
    worksheet: gspread.Worksheet,
    sheet_row_numbers: list[int],
    column_indices: dict[str, int],
) -> None:
    """Color the classification columns (Business Unit..Narration) red on
    every newly-classified row, so the accounts team can see at a glance
    which rows are auto-generated and still need manual verification —
    they simply change the text color to black once checked. Failures are
    logged but never raised; correct data with default formatting is still
    useful even if the color-coding doesn't apply."""
    if not sheet_row_numbers:
        return

    # Color only these 5 specific columns — NOT a min..max span, since the
    # blank columns interspersed between them in the sheet layout (SUB HEAD,
    # RECO, CONCERN, CUST ID, APT#, ACC REMARKS, CRM REMARKS) must never be
    # colored, and REFERENCE/DEBITS/CREDITS/BALANCE (which sit before this
    # block) must never be touched either.
    target_columns = [
        BUSINESS_UNIT_COLUMN, HEAD_COLUMN, TYPE_RERA_IDW_COLUMN,
        TCP_HEAD_COLUMN, NARRATION_COLUMN,
    ]

    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": row - 1,  # 0-based, inclusive
                    "endRowIndex": row,  # 0-based, exclusive
                    "startColumnIndex": column_indices[column_name] - 1,  # 0-based, inclusive
                    "endColumnIndex": column_indices[column_name],  # 0-based, exclusive
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"foregroundColor": UNVERIFIED_TEXT_COLOR}
                    }
                },
                "fields": "userEnteredFormat.textFormat.foregroundColor",
            }
        }
        for row in sheet_row_numbers
        for column_name in target_columns
    ]

    try:
        worksheet.spreadsheet.batch_update({"requests": requests})
    except Exception as exc:
        log.warning("Could not apply unverified-row text color: %s", exc)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def classify_transactions(
    credentials_path: Path,
    worksheet_name: str,
    sheet_id: str = MASTER_SHEET_ID,
    spreadsheet: Optional[Any] = None,
) -> int:
    """Classify all unclassified transactions in one account's worksheet/tab.

    Args:
        credentials_path: Path to the Google service-account credentials JSON.
        worksheet_name: The account's worksheet/tab name (e.g. "YES BANK - 2477").
        sheet_id: Spreadsheet ID containing the account tabs.
        spreadsheet: Optional pre-opened gspread.Spreadsheet. If provided,
            skips re-authentication (faster when called from the pipeline).

    Returns:
        Number of rows updated.
    """
    if spreadsheet is not None:
        worksheet = spreadsheet.worksheet(worksheet_name)
    else:
        client = get_gspread_client(credentials_path)
        worksheet = open_account_worksheet(client, sheet_id, worksheet_name)

    header_row, column_indices = ensure_classification_columns(worksheet)

    return classify_rows(worksheet, header_row, column_indices)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify transactions in one account's Google Sheet tab (Head + Narration)."
    )

    parser.add_argument(
        "-c",
        "--credentials",
        type=Path,
        default=DEFAULT_CREDENTIALS,
    )

    parser.add_argument(
        "--sheet-id",
        default=MASTER_SHEET_ID,
        help="Override the spreadsheet ID.",
    )

    parser.add_argument(
        "--worksheet-name",
        required=True,
        help="The account's worksheet/tab name (e.g. 'YES BANK - 2477').",
    )

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        updated = classify_transactions(
            credentials_path=args.credentials,
            sheet_id=args.sheet_id,
            worksheet_name=args.worksheet_name,
        )
        log.info("Classification complete. Rows updated: %d", updated)
    except Exception as exc:
        log.exception(exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
