"""Sync missing transactions from an account's automated Google Sheet tab
into the Accounts Team's own Main Sheet tab, append-only.

This module owns "get missing data into the Main Sheet correctly" as a
separate concern from upload_to_sheets.py, which owns "get data into the
automated sheet correctly" in the first place. The two sheets are
compared and reconciled independently on every call — there is no saved
"last synced" cursor, so re-running is always safe (see
find_missing_rows()'s docstring for why).

Currently DPL-only, tested against a COPY of the live Main Sheet before
ever pointing at the real one.
"""

from __future__ import annotations

import json
import re
import time
from collections import Counter, defaultdict
from datetime import date as _date
from pathlib import Path
from typing import Any, Callable, TypeVar

from dateutil import parser as _date_parser

_MAIN_SHEETS_CONFIG_PATH = Path(__file__).resolve().parent / "config" / "main_sheets.json"

# Same red used by classify_transactions.py's UNVERIFIED_TEXT_COLOR to mark
# auto-classified, not-yet-human-verified cells on the automated sheet — the
# Main Sheet's freshly-synced rows are equally unverified, so they get the
# same visual signal (the accounts team changes it to black once checked).
UNVERIFIED_TEXT_COLOR = {"red": 0.8, "green": 0.0, "blue": 0.0}
UNVERIFIED_COLOR_COLUMNS = ["BUSINESS UNIT", "HEAD", "TYPE FOR RERA IDW", "TCP Head", "NARRATION"]

_T = TypeVar("_T")


def _is_quota_error(exc: Exception) -> bool:
    """True if exc is a gspread APIError caused by a 429 quota
    response — same check web_app.py's own _is_quota_error() already
    uses, duplicated here rather than imported to avoid a circular
    import (web_app.py doesn't import this module)."""
    import gspread.exceptions

    if not isinstance(exc, gspread.exceptions.APIError):
        return False
    try:
        return exc.response.status_code == 429
    except Exception:
        return False


def call_with_retry(func: Callable[[], _T], *, max_attempts: int = 4, base_delay: float = 2.0) -> _T:
    """Call func() with short exponential backoff on a 429 quota error
    (2s, 4s, 8s, ...) — confirmed necessary against real data: this
    session hit explicit Google Sheets 429 errors repeatedly under
    heavy use, and the sync step (several reads/writes per account, on
    top of everything else the pipeline already does in one request)
    is exactly the kind of call sequence that trips it. Any other
    exception is raised immediately, unretried — a 429 is the only
    error where "the same call would probably succeed a moment later"
    actually holds."""
    last_exc: Exception | None = None
    for attempt in range(max_attempts):
        try:
            return func()
        except Exception as exc:
            if not _is_quota_error(exc) or attempt == max_attempts - 1:
                raise
            last_exc = exc
            time.sleep(base_delay * (2 ** attempt))
    raise last_exc  # pragma: no cover — loop always returns or raises above


def get_main_sheet_id_for_company(company: str | None) -> str | None:
    """Resolve a company name to its Accounts Team Main Sheet's
    spreadsheet ID.

    Looks up main_sheet_store (Admin -> Main Sheet Sync Settings) first —
    the web-configurable destination, verified against the live sheet at
    save time. Falls back to config/main_sheets.json if no row exists for
    this company yet, so nothing breaks for a company that hasn't been
    reconfigured through the UI. Returns None if neither source has a
    Main Sheet for this company (e.g. AMB before it's set up) — callers
    should skip the sync entirely in that case rather than guess a
    fallback spreadsheet."""
    from upload_to_sheets import DEFAULT_COMPANY, _extract_sheet_id_from_url

    company = (company or DEFAULT_COMPANY).strip() or DEFAULT_COMPANY

    import main_sheet_store

    row = main_sheet_store.get_main_sheet_link(company)
    if row and row.get("sheet_url"):
        extracted = _extract_sheet_id_from_url(row["sheet_url"])
        if extracted:
            return extracted

    try:
        with open(_MAIN_SHEETS_CONFIG_PATH, encoding="utf-8") as f:
            config = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None
    url = config.get(company)
    if not url:
        return None
    return _extract_sheet_id_from_url(url)

# How close two dates must be to treat a same-amount/same-description
# pair as "probably the same transaction with a corrected date" rather
# than two separate occurrences. Confirmed necessary against real data:
# some transaction descriptions (e.g. a recurring internal transfer
# between two of DPL's own accounts) are near-identical boilerplate text
# that recurs every 1-2 weeks for months — without this limit, a brand
# new occurrence would be wrongly matched against an old leftover entry
# from weeks earlier just because the amount and description happen to
# coincide, and silently treated as "already present" instead of a
# genuinely new transaction.
_DATE_PROXIMITY_DAYS = 3


# ---------------------------------------------------------------------------
# Normalization — building a stable "fingerprint" for one transaction
# ---------------------------------------------------------------------------
# Deliberately does NOT use the bank's own REFERENCE number as (part of)
# the key: config/bank_profiles.json confirms Bank of Maharashtra never
# populates that column at all ("excluded_fields": ["reference"]), so it
# can't be trusted as a universal identifier across every bank this
# project handles.
#
# DESCRIPTION *is* included here, unlike upload_to_sheets.py's own
# UNIQUE_KEY_COLUMNS (which deliberately excludes it). That exclusion
# exists because upload_to_sheets.py compares two INDEPENDENT PDF
# extraction passes of the same statement, where OCR/parsing noise can
# shift description text slightly run-to-run. This module compares
# against the Main Sheet, whose rows are (today) literally copy-pasted
# from the automated sheet by the Accounts Team — the same source text,
# not a second extraction — so DESCRIPTION is a reliable anchor here, and
# including it is what correctly distinguishes two different transactions
# that happen to share the same date and amount.

_WHITESPACE_RE = re.compile(r"\s+")


def normalize_date(value: Any) -> str:
    """Parse and reformat to a canonical ISO date (YYYY-MM-DD) — unlike
    upload_to_sheets.py's own load_existing_data(), which only strips
    whitespace (safe there, since it compares rows written by this same
    pipeline in one consistent format). Confirmed against the real Main
    Sheet copy that this can't be assumed here: our automated sheet
    writes 4-digit years ("06-Jul-2026") while the Accounts Team's Main
    Sheet uses 2-digit years and no leading zero on the day
    ("6-Jul-26") — a plain string comparison would treat the same real
    date as two different transactions and silently duplicate rows.
    dayfirst=True since every format observed is day-month-year; a
    genuinely blank/unparseable date returns "" rather than raising, so
    a bad date cell can't crash the whole sync (it just won't match
    anything, which is the safe failure mode for a sync tool)."""
    text = str(value if value is not None else "").strip()
    if not text:
        return ""
    try:
        return _date_parser.parse(text, dayfirst=True).strftime("%Y-%m-%d")
    except (ValueError, OverflowError):
        return text


def normalize_amount(value: Any) -> float:
    """Same convention as upload_to_sheets.py's load_existing_data():
    strip thousands-separator commas, parse to float, round to paise,
    blank/unparseable -> 0.0 (so a blank DEBITS cell and a blank CREDITS
    cell always compare equal to each other, not as distinct values)."""
    if value is None:
        return 0.0
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return round(float(text), 2)
    except ValueError:
        return 0.0


def normalize_description(value: Any) -> str:
    """Uppercase and strip ALL whitespace (not just collapse repeats)
    — confirmed necessary against real data: a stray single space
    sometimes lands in the middle of a reference number between the
    automated sheet and the Main Sheet (e.g. "...HDFCH01173383398" vs
    "...H DFCH01173383398", almost certainly a PDF-extraction word-wrap
    artifact), which a same-real-transaction should still match
    exactly on. These are fixed-format bank reference/narration
    strings, not free-form text, so removing whitespace entirely
    doesn't risk conflating two genuinely different transactions."""
    text = str(value if value is not None else "").upper()
    return _WHITESPACE_RE.sub("", text)


def _parse_iso_date(value: str) -> _date | None:
    """Parse a normalize_date() output (YYYY-MM-DD) back into a date
    object for day-difference math. Returns None for blank/unparseable
    input rather than raising."""
    if not value:
        return None
    try:
        year, month, day = (int(part) for part in value.split("-"))
        return _date(year, month, day)
    except (ValueError, TypeError):
        return None


def build_transaction_key(row: dict[str, Any]) -> tuple[str, float, float, str]:
    """The composite fingerprint: (date, debit, credit, description),
    each normalized. Two rows with the same key are considered the same
    real-world transaction."""
    return (
        normalize_date(row.get("TXN DATE")),
        normalize_amount(row.get("DEBITS")),
        normalize_amount(row.get("CREDITS")),
        normalize_description(row.get("DESCRIPTION")),
    )


# ---------------------------------------------------------------------------
# Comparison — 3-way classification against the Main Sheet
# ---------------------------------------------------------------------------
# An automated-sheet row lands in exactly one bucket:
#   "present" — an exact key match already exists in the Main Sheet;
#               nothing to do.
#   "new"     — no similar transaction exists at all; safe to append.
#   "review"  — a CLOSE but not exact match exists: same date+amount
#               with a different description, or same amount+description
#               with a different date. Confirmed against the real copy
#               sheet this genuinely happens (one BOM 675 transaction is
#               dated a day apart between the two sheets; another has
#               extra digits appended to its description on the Main
#               Sheet side) — these are most likely the same real
#               transaction with a manual edit/correction already made
#               directly in the Main Sheet, not a new transaction. Never
#               auto-appended: appending would create a near-duplicate
#               that looks like a second, separate transaction. Surfaced
#               for a human to resolve instead, same "don't guess"
#               principle this whole project already follows elsewhere.

def classify_for_sync(
    automated_rows: list[dict[str, Any]],
    main_rows: list[dict[str, Any]],
) -> dict[str, Any]:
    """Classify every automated-sheet row against the Main Sheet's
    current contents. Returns:
        {
            "new": [row, ...],              # append these, in order
            "review": [{"automated": row, "main": row, "reason": str}, ...],
            "present_count": int,           # already-matched count, for logging
        }

    Full-set comparison, not a "last date" cursor: every call recomputes
    the answer fresh from both sheets' current contents, so it is
    correct regardless of upload order, reprocessed statements, or
    out-of-order transaction dates, and naturally idempotent — running
    this twice with no new data in between always returns "new": [] and
    "review": [], because the second call's main_rows already contains
    everything the first call appended.

    Exact matching is multiset-based, not set-based: if the exact same
    transaction (same key) genuinely occurs twice in the automated sheet
    (e.g. two identical fee charges on the same day), and only one copy
    exists in the Main Sheet, exactly one copy is reported as new —
    never zero (which would silently drop a real transaction) and never
    two (which would create an unwanted duplicate).
    """
    main_keys = [build_transaction_key(row) for row in main_rows]
    main_key_counts: Counter[tuple[str, float, float, str]] = Counter(main_keys)

    unmatched_automated: list[dict[str, Any]] = []
    present_count = 0
    for row in automated_rows:
        key = build_transaction_key(row)
        if main_key_counts[key] > 0:
            main_key_counts[key] -= 1
            present_count += 1
        else:
            unmatched_automated.append(row)

    # Reconstruct exactly which physical Main Sheet rows were never
    # claimed by an exact match — main_key_counts now holds, per key,
    # how many instances are still unconsumed.
    remaining = dict(main_key_counts)
    leftover_main_rows: list[dict[str, Any]] = []
    for row, key in zip(main_rows, main_keys):
        if remaining.get(key, 0) > 0:
            leftover_main_rows.append(row)
            remaining[key] -= 1

    # Two weak-match indexes over the leftover Main Sheet rows only —
    # a row already claimed by an exact match is never offered as a
    # "close match" candidate for a different automated row.
    weak_by_date_amount: dict[tuple[str, float, float], list[dict[str, Any]]] = defaultdict(list)
    weak_by_amount_desc: dict[tuple[float, float, str], list[dict[str, Any]]] = defaultdict(list)
    for row in leftover_main_rows:
        date, debit, credit, desc = build_transaction_key(row)
        weak_by_date_amount[(date, debit, credit)].append(row)
        weak_by_amount_desc[(debit, credit, desc)].append(row)

    def _claim(row: dict[str, Any]) -> None:
        # Remove from BOTH indexes so this same Main Sheet row can't be
        # matched a second time against a different automated row.
        date, debit, credit, desc = build_transaction_key(row)
        bucket = weak_by_date_amount[(date, debit, credit)]
        if row in bucket:
            bucket.remove(row)
        bucket = weak_by_amount_desc[(debit, credit, desc)]
        if row in bucket:
            bucket.remove(row)

    new_rows: list[dict[str, Any]] = []
    review_items: list[dict[str, Any]] = []
    for row in unmatched_automated:
        date, debit, credit, desc = build_transaction_key(row)

        candidates = weak_by_date_amount.get((date, debit, credit))
        if candidates:
            matched = candidates[0]
            _claim(matched)
            review_items.append({
                "automated": row,
                "main": matched,
                "reason": "same date and amount, different description",
            })
            continue

        candidates = weak_by_amount_desc.get((debit, credit, desc))
        if candidates:
            # Only accept the closest candidate, and only if it's within
            # _DATE_PROXIMITY_DAYS — a recurring transaction with
            # boilerplate description (e.g. a routine internal transfer)
            # can otherwise false-match against an old, unrelated
            # occurrence purely because the amount/description repeat.
            this_date = _parse_iso_date(date)
            closest, closest_diff = None, None
            for candidate in candidates:
                candidate_date = _parse_iso_date(build_transaction_key(candidate)[0])
                if this_date is None or candidate_date is None:
                    continue
                diff = abs((this_date - candidate_date).days)
                if diff <= _DATE_PROXIMITY_DAYS and (closest_diff is None or diff < closest_diff):
                    closest, closest_diff = candidate, diff
            if closest is not None:
                _claim(closest)
                review_items.append({
                    "automated": row,
                    "main": closest,
                    "reason": "same amount and description, different date",
                })
                continue

        new_rows.append(row)

    return {"new": new_rows, "review": review_items, "present_count": present_count}


def find_missing_rows(
    automated_rows: list[dict[str, Any]],
    main_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Convenience wrapper: just the rows safe to append (the "new"
    bucket of classify_for_sync()). "review" rows are deliberately
    excluded — call classify_for_sync() directly to also see those."""
    return classify_for_sync(automated_rows, main_rows)["new"]


# ---------------------------------------------------------------------------
# Sheet I/O — reading both tabs, in the shape classify_for_sync() expects
# ---------------------------------------------------------------------------
# The automated sheet's header is row 1; the Main Sheet's header is row 2
# (row 1 holds a "LAST UPDATE"/opening-totals summary row) — confirmed by
# directly inspecting the real DPL copy sheet. Both are handled by the
# same reader, parameterized by header_row (1-indexed).

def read_sheet_rows(worksheet: Any, header_row: int) -> tuple[list[str], list[dict[str, Any]]]:
    """Read a worksheet's header and data rows as dicts, skipping any row
    whose TXN DATE cell doesn't actually parse as a date (spacer/summary
    rows, or a boundary-marker row — see read_sheet_snapshot())."""
    all_values = call_with_retry(worksheet.get_all_values)
    header = all_values[header_row - 1]
    date_idx = header.index("TXN DATE")
    rows = []
    for raw_row in all_values[header_row:]:
        if len(raw_row) <= date_idx or not _parse_iso_date(normalize_date(raw_row[date_idx])):
            continue
        rows.append({header[i]: (raw_row[i] if i < len(raw_row) else "") for i in range(len(header))})
    return header, rows


# ---------------------------------------------------------------------------
# Column mapping onto the Main Sheet's own header
# ---------------------------------------------------------------------------
# Confirmed by inspecting the real DPL copy sheet: every column name up
# through NARRATION matches our automated sheet's EXPECTED_COLUMNS
# exactly, with exactly one rename.
MAIN_SHEET_HEADER_ALIASES: dict[str, str] = {
    "Bal as per AI": "Balance (AI)",
}

# Columns that exist on the Main Sheet but must NEVER be populated by the
# sync: SL#/QTR/MONTH/BALANCE/Check are formula-driven (handled
# separately, see write_formula_columns() below) — a plain value here
# would break their live chain. Auth MB/MK/NM and Tax Rate are the
# Accounts Team's own manual sign-off/working columns with no
# automated-sheet equivalent at all; a newly appended/inserted row
# leaves them blank for a human to fill in, never guessed.
MAIN_SHEET_FORMULA_COLUMNS = {"SL#", "QTR", "MONTH", "BALANCE", "Check"}
MAIN_SHEET_MANUAL_COLUMNS = {"Auth MB", "Auth MK", "Auth NM", "Tax Rate"}


# Columns that must be written as real numbers, not comma-formatted text
# — confirmed necessary the hard way: writing "1,98,629" as a RAW string
# stores it as literal TEXT in Sheets, which then breaks any formula
# elsewhere that does arithmetic on that cell (BALANCE's chained
# addition/subtraction raised "#VALUE!" on the first live write tested).
# Same fields upload_to_sheets.py's own append_unique_rows() already
# treats as numeric for the same reason.
MAIN_SHEET_NUMERIC_COLUMNS = {"DEBITS", "CREDITS", "Bal as per AI"}


def map_row_to_main_header(row: dict[str, Any], main_header: list[str]) -> list[Any]:
    """Build one row of plain values, in the Main Sheet's own column
    order. Matches by column name (case-sensitive exact match first,
    then the known alias table); MAIN_SHEET_FORMULA_COLUMNS are left
    blank here (filled in separately as live formulas) and
    MAIN_SHEET_MANUAL_COLUMNS are always left blank. Any Main Sheet
    column with no automated-sheet equivalent at all is also left
    blank, never guessed. DEBITS/CREDITS/Bal-as-per-AI are written as
    actual numbers, not text — see MAIN_SHEET_NUMERIC_COLUMNS."""
    values: list[Any] = []
    for col_name in main_header:
        if col_name in MAIN_SHEET_FORMULA_COLUMNS or col_name in MAIN_SHEET_MANUAL_COLUMNS:
            values.append("")
            continue
        source_name = MAIN_SHEET_HEADER_ALIASES.get(col_name, col_name)
        raw_value = row.get(source_name, "")
        if col_name in MAIN_SHEET_NUMERIC_COLUMNS:
            values.append(normalize_amount(raw_value))
        else:
            values.append(str(raw_value))
    return values


# ---------------------------------------------------------------------------
# Chronological direction — which end of the existing data new rows join
# ---------------------------------------------------------------------------

def detect_sort_direction(main_rows: list[dict[str, Any]]) -> str:
    """"descending" (newest first, confirmed on the YES Bank tabs) or
    "ascending" (oldest first, confirmed on the BOM tabs) — by comparing
    the first and last data row's own dates, the same approach
    upload_to_sheets.py already uses for its own sheet. Defaults to
    "ascending" (append at the bottom, the lower-risk operation) if the
    direction can't be determined from too little data."""
    dates = [normalize_date(row.get("TXN DATE")) for row in main_rows]
    dates = [d for d in dates if d]
    if len(dates) < 2:
        return "ascending"
    return "descending" if dates[0] > dates[-1] else "ascending"


def find_data_row_bounds(worksheet: Any, header_row: int) -> tuple[int | None, int | None]:
    """The actual 1-indexed sheet row numbers of the first and last data
    row (first non-blank-TXN-DATE row after the header, and the last
    one). Confirmed necessary against the real Main Sheet: several tabs
    have a few blank rows between the header and where transactions
    actually start (an Accounts Team formatting convention), so "row
    right after the header" is NOT always where data begins — inserting
    there would push those blank spacer rows down rather than sit
    directly above the real data. Returns (None, None) if the tab has no
    data rows at all.

    Standalone convenience wrapper around read_sheet_snapshot() — costs
    its own get_all_values() call. sync_account_to_main_sheet() does
    NOT call this directly; it uses one shared snapshot instead (see
    read_sheet_snapshot()) to avoid re-reading the same sheet 3 times."""
    return read_sheet_snapshot(worksheet, header_row)["bounds"]


def read_sheet_rows_with_positions(worksheet: Any, header_row: int) -> list[tuple[int, dict[str, Any]]]:
    """Like read_sheet_rows(), but keeps each row's actual 1-indexed
    sheet row number attached — needed to pinpoint an exact insertion
    gap by physical position, not just by list order.

    Standalone convenience wrapper around read_sheet_snapshot() — see
    that function's note on why sync_account_to_main_sheet() doesn't
    call this directly."""
    return read_sheet_snapshot(worksheet, header_row)["rows_with_positions"]


def read_sheet_snapshot(worksheet: Any, header_row: int) -> dict[str, Any]:
    """One get_all_values() call, with everything read_sheet_rows() /
    find_data_row_bounds() / read_sheet_rows_with_positions() each used
    to fetch separately now derived from it in pure Python. Confirmed
    necessary: those 3 functions each doing their own full-sheet read
    meant syncing a single account made 3 redundant API calls against
    the same (sometimes 200+ row) Main Sheet tab, on top of everything
    else the pipeline already does in the same request — a real
    contributor to hitting Google Sheets' rate limit / running slow
    enough to risk a request timeout, confirmed by direct reproduction.

    Returns {"header": [...], "rows": [...], "rows_with_positions":
    [(row_number, row_dict), ...], "bounds": (first_data_row,
    last_data_row)}."""
    all_values = call_with_retry(worksheet.get_all_values)
    header = all_values[header_row - 1]
    date_idx = header.index("TXN DATE")

    rows: list[dict[str, Any]] = []
    rows_with_positions: list[tuple[int, dict[str, Any]]] = []
    first_row, last_row = None, None
    for i, raw_row in enumerate(all_values[header_row:], start=header_row + 1):
        # Requires the TXN DATE cell to actually PARSE as a date, not just
        # be non-blank — confirmed necessary against real data: some Main
        # Sheets have pre-existing "boundary marker" rows (every column
        # filled with a literal placeholder like "x") separating historical
        # bulk-imported data from newer tracked data. A blank-only check
        # wrongly treated such a row as the last real transaction, causing
        # new rows to insert in the wrong place, copy formulas from a
        # non-formula row, and overwrite the marker row's own BALANCE cell.
        if len(raw_row) <= date_idx or not _parse_iso_date(normalize_date(raw_row[date_idx])):
            continue
        row_dict = {header[j]: (raw_row[j] if j < len(raw_row) else "") for j in range(len(header))}
        rows.append(row_dict)
        rows_with_positions.append((i, row_dict))
        if first_row is None:
            first_row = i
        last_row = i

    return {
        "header": header,
        "rows": rows,
        "rows_with_positions": rows_with_positions,
        "bounds": (first_row, last_row),
    }


def _extract_balance(row: dict[str, Any]) -> float:
    """The PDF's own true reported balance for this transaction — key
    name differs between the two sheets ("Balance (AI)" on the
    automated sheet, "Bal as per AI" on the Main Sheet).

    Falls back to the Main Sheet's own live "BALANCE" formula column
    when "Bal as per AI" is blank — confirmed necessary against real
    data: older historical rows on the real Main Sheet copy have an
    empty "Bal as per AI" cell (that column looks like it was only
    backfilled from a certain point onward), which would otherwise
    silently resolve to 0.0 and break the balance-chain match right at
    that row. BALANCE is a formula, never blank for a real data row,
    and is derived from the same underlying credit/debit chain, so
    it's a reliable substitute for matching purposes even where the
    PDF-verified figure isn't available."""
    if "Balance (AI)" in row:
        return normalize_amount(row.get("Balance (AI)"))
    ai_value = str(row.get("Bal as per AI", "")).strip()
    if ai_value:
        return normalize_amount(ai_value)
    return normalize_amount(row.get("BALANCE"))


# How close two balance values must be to count as "the same point in
# the running total" — not an exact-equality requirement, to absorb the
# same rounding noise already observed elsewhere in the Main Sheet's
# own historical data (confirmed real drift of ~1.78 and ~0.3-0.4 on
# real rows in this session's testing — the original 1.0 tolerance was
# too tight to cover the 1.78 case and silently failed to find a real
# match). 5.0 comfortably covers observed drift while staying far
# below what any two genuinely different transactions would need to
# coincidentally differ by to false-match.
_BALANCE_MATCH_TOLERANCE = 5.0


def resolve_insertion_plan(
    main_rows_with_positions: list[tuple[int, dict[str, Any]]],
    new_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Figure out exactly where each new row belongs in the Main
    Sheet's existing chronological chain, using the running-balance
    identity rather than assuming everything new is either newer or
    older than everything that already exists. For descending
    (newest-first) sheets: the row immediately ABOVE (chronologically
    newer, physically higher) any given row X satisfies
    `above.BALANCE == X.BALANCE + above.CREDIT - above.DEBIT`. A new
    row's correct position is found by matching this identity against
    both the existing Main Sheet rows AND the other new rows in this
    same batch (so a contiguous run of several new rows attaches to the
    existing data as one block, in the right internal order).

    Returns a list of insertion blocks:
        [{"upper_row": int | None, "lower_row": int | None,
          "rows": [new row dicts, newest-first]}, ...]
    "upper_row" is the existing sheet row immediately above the gap
    (None if this block is newer than everything currently in the
    sheet — i.e. belongs at the very top). "lower_row" is the existing
    row immediately below the gap (None if older than everything —
    belongs at the very bottom). Any new row that can't be matched to
    any position at all (no existing row's balance connects to it) is
    treated as belonging at the very top, consistent with "no evidence
    it's anything but the newest thing that's happened" — same
    don't-guess-past-the-evidence approach used throughout this module,
    just erring toward the lower-risk edge (top-insert, not a random
    middle position).
    """
    existing = [{"row": r, "balance": _extract_balance(row), "credit": normalize_amount(row.get("CREDITS")),
                 "debit": normalize_amount(row.get("DEBITS"))} for r, row in main_rows_with_positions]
    pending = [{"row": None, "balance": _extract_balance(row), "credit": normalize_amount(row.get("CREDITS")),
                "debit": normalize_amount(row.get("DEBITS")), "source": row} for row in new_rows]

    pool = existing + pending

    def find_lower_neighbor(item: dict[str, Any]) -> dict[str, Any] | None:
        # The item directly below/older than `item`: its balance equals
        # item.balance - item.credit + item.debit (undoing item's own
        # transaction gets you the balance right before it happened,
        # which is exactly the next-older row's own resulting balance).
        target = item["balance"] - item["credit"] + item["debit"]
        for candidate in pool:
            if candidate is item:
                continue
            if abs(candidate["balance"] - target) <= _BALANCE_MATCH_TOLERANCE:
                return candidate
        return None

    # Link every pending (new) row to its lower (older) neighbor, if any
    # — this can be an existing row OR another pending row, so several
    # new rows that belong right next to each other form one run.
    lower_of: dict[int, dict[str, Any]] = {}
    for item in pending:
        neighbor = find_lower_neighbor(item)
        if neighbor is not None:
            lower_of[id(item)] = neighbor

    # A run must start from its "head" — the pending row nothing else in
    # this batch points down into — not an arbitrary pending row, or the
    # walk below could start mid-chain and build the run backwards.
    pending_targets = {id(v) for v in lower_of.values() if v in pending}
    heads = [item for item in pending if id(item) not in pending_targets]

    placed: set[int] = set()
    blocks: list[dict[str, Any]] = []
    for item in heads:
        if id(item) in placed:
            continue
        run = [item]
        placed.add(id(item))
        cursor = item
        lower_anchor = None
        while True:
            neighbor = lower_of.get(id(cursor))
            if neighbor is None:
                break
            if neighbor in existing:
                lower_anchor = neighbor["row"]
                break
            if id(neighbor) in placed:
                # Already visited via a different head's walk (e.g. two
                # pending rows both resolving to the same lower
                # neighbor) — confirmed necessary against real data:
                # without this check, a convergent (non-cyclic but
                # overlapping) chain structure caused the same nodes to
                # be re-walked repeatedly, hanging for 27 same-day rows
                # against real Main Sheet data. Stop here rather than
                # duplicate this row into a second block.
                break
            # neighbor is another pending row, continuing this same run
            run.append(neighbor)
            placed.add(id(neighbor))
            cursor = neighbor

        # This run's upper anchor: the existing row (if any) whose own
        # lower-neighbor target matches this run's topmost balance.
        upper_anchor = None
        for e in existing:
            target = e["balance"] - e["credit"] + e["debit"]
            if abs(target - run[0]["balance"]) <= _BALANCE_MATCH_TOLERANCE:
                upper_anchor = e["row"]
                break

        blocks.append({
            "upper_row": upper_anchor,
            "lower_row": lower_anchor,
            "rows": [r["source"] for r in run],
        })

    # Defensive fallback: any pending row that never got placed (e.g. a
    # genuine cycle, which shouldn't happen with real financial data) is
    # still reported rather than silently dropped — treated as its own
    # unattached, "newest" block, the same safe default used when no
    # match is found at all.
    for item in pending:
        if id(item) not in placed:
            blocks.append({"upper_row": None, "lower_row": None, "rows": [item["source"]]})

    return blocks


# ---------------------------------------------------------------------------
# Formula replication — QTR/MONTH/Check (per-row) and BALANCE (chained)
# ---------------------------------------------------------------------------
# Confirmed against real data: these columns hold live formulas on the
# Main Sheet, not static values (e.g. QTR = '=IFERROR("Q"&(INT(MOD(C6-4,
# 12)/3)+1),"")', BALANCE = '=K7+J6-I6' — a running total chained from
# the chronologically-adjacent row). Writing plain text into these would
# permanently break that row's live calculation. Instead, an existing
# row's own formula is read and re-targeted at the new row number(s) —
# QTR/MONTH/Check only ever reference cells in their OWN row (confirmed
# by inspecting real formulas on both the YES Bank and BOM tabs), so
# renumbering is a safe find-and-replace of that one row number. BALANCE
# is handled separately since it references an ADJACENT row, not itself.

_ROW_REF_RE_TEMPLATE = r"([A-Za-z]+){row}(?!\d)"


def renumber_same_row_formula(formula: str, old_row: int, new_row: int) -> str:
    """Re-target a formula that only references its own row (QTR, MONTH,
    Check) from old_row to new_row."""
    pattern = re.compile(_ROW_REF_RE_TEMPLATE.format(row=old_row))
    return pattern.sub(lambda m: f"{m.group(1)}{new_row}", formula)


def build_chained_balance_formula(
    row: int, adjacent_row: int, credit_col: str, debit_col: str, balance_col: str
) -> str:
    """A running-balance formula in the same shape confirmed on every
    real sample checked: this row's balance = the chronologically
    PREVIOUS row's balance (wherever that physically sits — one row
    below in descending/newest-first tabs, one row above in
    ascending/oldest-first tabs) + this row's credit - this row's debit."""
    return f"={balance_col}{adjacent_row}+{credit_col}{row}-{debit_col}{row}"


def col_letter(header: list[str], column_name: str) -> str:
    import gspread.utils as _gs_utils
    return _gs_utils.rowcol_to_a1(1, header.index(column_name) + 1).rstrip("0123456789")


def mark_rows_unverified(main_ws: Any, main_header: list[str], first_row: int, last_row: int) -> None:
    """Color UNVERIFIED_COLOR_COLUMNS red on rows first_row..last_row
    (inclusive), matching classify_transactions.py's own
    _mark_rows_unverified() exactly — same color, same signal, same
    "accounts team changes it to black once verified" convention,
    just applied to the Main Sheet's freshly-synced rows instead of
    the automated sheet's freshly-classified ones. Columns not present
    in this particular tab's header are silently skipped rather than
    raising (matches the rest of this module's "don't assume every
    tab has every column" handling)."""
    target_columns = [c for c in UNVERIFIED_COLOR_COLUMNS if c in main_header]
    if not target_columns or first_row > last_row:
        return
    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": main_ws.id,
                    "startRowIndex": first_row - 1,
                    "endRowIndex": last_row,
                    "startColumnIndex": main_header.index(column_name),
                    "endColumnIndex": main_header.index(column_name) + 1,
                },
                "cell": {"userEnteredFormat": {"textFormat": {"foregroundColor": UNVERIFIED_TEXT_COLOR}}},
                "fields": "userEnteredFormat.textFormat.foregroundColor",
            }
        }
        for column_name in target_columns
    ]
    call_with_retry(lambda: main_ws.spreadsheet.batch_update({"requests": requests}))


# ---------------------------------------------------------------------------
# Main orchestrator
# ---------------------------------------------------------------------------

def sync_account_to_main_sheet(
    automated_ws: Any,
    main_ws: Any,
    *,
    automated_header_row: int = 1,
    main_header_row: int = 2,
    dry_run: bool = True,
) -> dict[str, Any]:
    """Compare one account's automated-sheet tab against its Main Sheet
    tab and append/insert only the transactions that are missing.

    dry_run=True (default): computes and returns everything that WOULD
    happen, writes nothing. dry_run=False: actually writes, then
    re-reads the Check column for every newly written row (and the
    row immediately adjacent to the insertion point) to confirm the
    running balance is self-consistent — a wrong chronological position
    or a credit/debit sign error would show up there as a non-zero
    Check value, since Check = Balance (AI) - BALANCE and Balance (AI)
    is the PDF's own independently-reported true balance.

    Returns a report dict — see the "report" variable below for exact
    keys — safe to log or show to a human either way.
    """
    import gspread

    import gspread.utils as gs_utils

    automated_header, automated_rows = read_sheet_rows(automated_ws, automated_header_row)
    # One shared snapshot of the Main Sheet — used for classification,
    # direction detection, data-row bounds, AND (for descending tabs)
    # the positioned rows resolve_insertion_plan() needs. Previously
    # each of those was its own get_all_values() call against the same
    # sheet; now it's one, regardless of how many of them are needed.
    main_snapshot = read_sheet_snapshot(main_ws, main_header_row)
    main_header, main_rows = main_snapshot["header"], main_snapshot["rows"]

    classification = classify_for_sync(automated_rows, main_rows)
    new_rows = classification["new"]

    report: dict[str, Any] = {
        "present_count": classification["present_count"],
        "review": classification["review"],
        "new_count": len(new_rows),
        "written": False,
        "direction": None,
        "insert_row": None,
        "validation": [],
    }

    if not new_rows:
        return report

    direction = detect_sort_direction(main_rows)
    report["direction"] = direction

    first_data_row, last_data_row = main_snapshot["bounds"]
    # Tracked separately from the first_data_row/last_data_row fallback
    # below — confirmed necessary against real data: a genuinely empty
    # tab (no transaction rows at all yet) has no real row to anchor
    # formula-copying or the BALANCE chain to. Without this flag, the
    # code treated the HEADER row itself as if it were a real data row
    # once first_data_row/last_data_row got defaulted to main_header_row,
    # copying whatever formulas happened to live in the header row
    # (in one real case, cross-sheet reference formulas unrelated to a
    # per-row QTR/MONTH pattern) verbatim into every new row, and
    # chaining BALANCE off nothing — producing wrong, not just blank,
    # numbers.
    tab_was_empty = first_data_row is None
    if first_data_row is None:
        # Empty tab (no data rows at all yet) — insert right after the
        # header, nothing to chain a running balance from.
        first_data_row = last_data_row = main_header_row

    credit_col = col_letter(main_header, "CREDITS")
    debit_col = col_letter(main_header, "DEBITS")
    balance_col = col_letter(main_header, "BALANCE") if "BALANCE" in main_header else None

    if direction == "descending":
        # Fast path: if EVERY new row is chronologically newer than (or
        # equal to) the current top-of-data row's date, this is simply
        # a top-of-everything insert — no interior gap to locate at
        # all. Confirmed necessary against real data: skipping straight
        # to resolve_insertion_plan()'s balance-chain matching for this
        # case produced fragmented, spurious blocks (multiple new rows
        # falsely cross-matching each other via balance-tolerance
        # coincidence, purely because the Main Sheet's own historical
        # drift isn't uniform across the tab — tight enough tolerance
        # to avoid one false match was too tight for a real match
        # elsewhere, and vice versa). Same top-insert logic already
        # validated safe on the first 4 accounts tested.
        top_row = main_snapshot["rows_with_positions"][0][1] if main_snapshot["rows_with_positions"] else None
        top_date = normalize_date(top_row.get("TXN DATE")) if top_row else ""
        all_newer_than_top = all(normalize_date(r.get("TXN DATE")) >= top_date for r in new_rows) if top_date else True

        if all_newer_than_top:
            blocks = [{
                "upper_row": None,
                "lower_row": first_data_row,
                "insert_row": first_data_row,
                "rows": sorted(new_rows, key=lambda r: normalize_date(r.get("TXN DATE")), reverse=True),
            }]
        else:
            # A missing transaction can belong ANYWHERE in the existing
            # history (e.g. a deleted row from the middle), so the
            # exact gap is found via the running-balance identity
            # (resolve_insertion_plan), not assumed. Each resulting
            # block gets its own insert position; a block with no
            # matching existing row above it (upper_row is None)
            # genuinely belongs at the very top — that's the same "no
            # evidence otherwise" case validated in earlier testing.
            blocks = resolve_insertion_plan(main_snapshot["rows_with_positions"], new_rows)
            for block in blocks:
                block["insert_row"] = (block["upper_row"] + 1) if block["upper_row"] is not None else first_data_row
                # block["rows"] is already in the correct newest-to-oldest
                # chain order from resolve_insertion_plan's chain-walk — do
                # NOT re-sort by balance value, which is not monotonic with
                # chronology (depends on each row's own credit/debit sign).
    else:
        # Ascending (oldest-first, e.g. BOM tabs): mid-sheet insertion
        # isn't implemented yet — only rows chronologically newer than
        # (or equal to) the tab's current bottom-most row are safe to
        # append there. Confirmed necessary against real data: a
        # transaction dated weeks before the current last row got
        # force-appended after it anyway, producing a nonsensical
        # negative BALANCE (chained against a much smaller existing
        # balance than the transaction's own true prior history).
        # Anything older is left for manual placement instead of
        # guessed — same don't-guess-past-the-evidence principle
        # already applied to descending sheets' top-insert fast path.
        last_date = normalize_date(main_rows[-1].get("TXN DATE")) if (not tab_was_empty and main_rows) else ""
        appendable = [r for r in new_rows if not last_date or normalize_date(r.get("TXN DATE")) >= last_date]
        unplaceable = [r for r in new_rows if r not in appendable]
        for row in unplaceable:
            report["review"].append({
                "automated": row,
                "main": None,
                "reason": "dated before the tab's current last row on an ascending "
                          "(oldest-first) sheet — mid-sheet insertion isn't supported "
                          "yet, needs manual placement",
            })
        blocks = [{
            "upper_row": None if tab_was_empty else last_data_row,
            "lower_row": None,
            "insert_row": last_data_row + 1,
            "rows": sorted(appendable, key=lambda r: normalize_date(r.get("TXN DATE"))),
        }] if appendable else []

    # Recompute now that ascending's chronological-order check may have
    # moved some rows out of "new" and into "review" above — new_count
    # must reflect what's actually about to be appended, not the
    # original pre-filter total.
    report["new_count"] = sum(len(b["rows"]) for b in blocks)

    report["insert_row"] = [b["insert_row"] for b in blocks]
    report["blocks"] = [
        {"upper_row": b["upper_row"], "lower_row": b["lower_row"], "insert_row": b["insert_row"],
         "mapped": list(zip(b["rows"], [map_row_to_main_header(r, main_header) for r in b["rows"]]))}
        for b in blocks
    ]
    # True if any block has no real existing row to connect to on either
    # side — BALANCE is deliberately left blank for those rows (see
    # block_has_real_anchor below) rather than guessing a zero opening
    # balance, so callers/humans know to fill it in manually.
    report["balance_needs_manual_entry"] = any(
        b["upper_row"] is None and b["lower_row"] is None for b in blocks
    )

    if dry_run:
        return report

    # Process blocks bottom-to-top (highest insert_row first) so that
    # inserting one block never shifts the row numbers already computed
    # for a block above it.
    blocks_sorted = sorted(blocks, key=lambda b: b["insert_row"], reverse=True)

    formula_cols = [c for c in ("QTR", "MONTH", "Check") if c in main_header]
    sl_letter = col_letter(main_header, "SL#") if "SL#" in main_header else None

    for block in blocks_sorted:
        insert_row = block["insert_row"]
        upper_row = block["upper_row"]
        rows_sorted = block["rows"]
        n = len(rows_sorted)
        # anchor_row: the existing row whose formulas we copy the
        # per-row pattern from. Prefer upper_row (right above the new
        # block) — falls back to first_data_row only for a genuine
        # top-of-everything insert where a real row sits there (NOT for
        # a tab that started genuinely empty — tab_was_empty means
        # first_data_row is just the header row's own number, not a
        # real transaction row, and has no formula pattern worth
        # copying).
        anchor_row = upper_row if upper_row is not None else (None if tab_was_empty else first_data_row)

        # --- Read the anchor row's existing formulas BEFORE inserting,
        # since an insert at/above it would shift its row number. All
        # of QTR/MONTH/Check/SL# fetched in ONE batch_get call rather
        # than one .acell() round-trip each — confirmed necessary: the
        # per-cell version was slow/quota-heavy enough to contribute to
        # Google Sheets rate-limiting on real data. Skipped entirely
        # when there's no real anchor_row (cold-start empty tab) — new
        # rows are written with QTR/MONTH/Check/BALANCE left blank
        # rather than guessing a pattern from a nonexistent row. ---
        anchor_cols = list(formula_cols) + (["SL#"] if sl_letter else [])
        anchor_ranges = [f"{col_letter(main_header, col)}{anchor_row}" for col in anchor_cols] if anchor_row is not None else []
        anchor_values = call_with_retry(
            lambda: main_ws.batch_get(anchor_ranges, value_render_option=gs_utils.ValueRenderOption.formula)
        ) if anchor_ranges else []

        def _cell_value(matrix: list[list[Any]]) -> Any:
            return matrix[0][0] if matrix and matrix[0] else None

        anchor_formulas: dict[str, str | None] = {}
        for col, matrix in zip(formula_cols, anchor_values):
            value = _cell_value(matrix)
            anchor_formulas[col] = value if isinstance(value, str) and value.startswith("=") else None

        # SL# convention varies by tab (see MAIN_SHEET module docstring
        # notes elsewhere) — static number on some tabs, a self-
        # adjusting "=ROW()-N" formula on others.
        anchor_sl_static: int | None = None
        if sl_letter and anchor_row is not None:
            raw_value = _cell_value(anchor_values[len(formula_cols)])
            raw = str(raw_value).strip() if raw_value is not None else ""
            if re.fullmatch(r"=ROW\(\)-\d+", raw, re.IGNORECASE):
                anchor_formulas["SL#"] = raw
            else:
                try:
                    anchor_sl_static = int(raw)
                except (TypeError, ValueError):
                    anchor_sl_static = None

        # If this block is being inserted BELOW an existing row in a
        # DESCENDING (newest-first) tab, that existing row's own BALANCE
        # formula currently points to whatever used to sit directly
        # below it. Google Sheets' native row-shift will auto-adjust
        # that reference to the OLD neighbor's new (shifted) position —
        # which is wrong, since it would skip straight over our newly
        # inserted block. It must be explicitly redirected to the new
        # block's topmost row instead.
        #
        # This does NOT apply to ASCENDING (oldest-first) tabs: there,
        # each row's own BALANCE formula points to the row ABOVE it
        # (chronologically previous), confirmed against every real
        # untouched sample checked — build_chained_balance_formula()'s
        # own docstring already documented this asymmetry, but the
        # write loop below previously ignored it. An existing row that
        # sits directly above a newly-appended block never needed
        # fixing at all: nothing points FROM it TO the rows below, so
        # inserting after it doesn't disturb its own formula. Applying
        # this "fix" anyway (the bug, now corrected) overwrote that
        # existing row's correct upward-pointing formula with a wrong
        # downward-pointing one — confirmed on real AMB data.
        upper_row_balance_fix = None
        if upper_row is not None and balance_col and direction == "descending":
            upper_row_balance_fix = build_chained_balance_formula(upper_row, insert_row, credit_col, debit_col, balance_col)

        # A block only has something real to chain BALANCE against if it
        # connects to an actual existing row on at least one side (above
        # or below). A block with neither (upper_row and lower_row both
        # None) is a cold-start case — the tab had zero real transaction
        # rows before this sync — and chaining anyway would silently
        # assume a zero opening balance, which is a guess, not a fact:
        # confirmed against real data that a "genuinely empty" tab can
        # still have real prior history the accounts team just hasn't
        # entered yet, and writing a false zero-anchored chain produced
        # confidently wrong (internally consistent, but wrong) negative
        # BALANCE values. Left blank instead for a human to fill in the
        # true opening balance.
        block_has_real_anchor = upper_row is not None or block.get("lower_row") is not None

        mapped_values = [map_row_to_main_header(r, main_header) for r in rows_sorted]
        call_with_retry(lambda: main_ws.insert_rows(mapped_values, row=insert_row, value_input_option="RAW"))
        mark_rows_unverified(main_ws, main_header, insert_row, insert_row + n - 1)

        updates: list[dict[str, Any]] = []
        if upper_row_balance_fix is not None:
            updates.append({"range": f"{balance_col}{upper_row}", "values": [[upper_row_balance_fix]]})

        for i in range(n):
            row_number = insert_row + i
            for col, anchor_formula in anchor_formulas.items():
                if anchor_formula is None:
                    continue
                letter = col_letter(main_header, col)
                new_formula = renumber_same_row_formula(anchor_formula, anchor_row, row_number)
                updates.append({"range": f"{letter}{row_number}", "values": [[new_formula]]})

            if balance_col and block_has_real_anchor:
                # Direction-aware, matching build_chained_balance_formula's
                # own documented convention: descending tabs chain to the
                # row BELOW (chronologically previous, physically lower);
                # ascending tabs chain to the row ABOVE (chronologically
                # previous, physically higher) — confirmed against real
                # untouched data on both DPL's BOM tabs and AMB's own
                # older rows. For ascending, rows_sorted is oldest-first,
                # so this row's own "previous" (row_number - 1) is either
                # the real anchor (upper_row, for the block's first/
                # oldest row) or the prior new row in this same batch —
                # both already correct without any special-casing.
                adjacent_row = row_number + 1 if direction == "descending" else row_number - 1
                formula = build_chained_balance_formula(row_number, adjacent_row, credit_col, debit_col, balance_col)
                updates.append({"range": f"{balance_col}{row_number}", "values": [[formula]]})

            if anchor_sl_static is not None and sl_letter and upper_row is None:
                # Only safe to extend a STATIC SL# sequence for a
                # genuine top-of-everything insert — extending it from
                # a mid-sheet anchor can collide with a value that
                # already exists further up (confirmed the hard way:
                # inserting between two existing rows produced a
                # duplicate SL# elsewhere in the tab). A full-sheet
                # renumber would avoid that, but disturbing hundreds of
                # untouched existing rows is a bigger risk than leaving
                # this one column blank here — left blank instead.
                sl_value = anchor_sl_static + (n - i)  # topmost (newest) new row gets the highest number
                updates.append({"range": f"{sl_letter}{row_number}", "values": [[sl_value]]})

        if updates:
            call_with_retry(lambda: main_ws.batch_update(updates, value_input_option=gspread.utils.ValueInputOption.user_entered))

        # Record this block's FINAL row numbers for validation. Blocks
        # are processed largest-insert_row-first, so every block
        # processed from here on inserts ABOVE this one and pushes it
        # (and every block already recorded) down by its own row count
        # — apply that shift retroactively to what's already recorded.
        block["final_rows"] = list(range(insert_row, insert_row + n))
        if upper_row is not None:
            block["final_rows"].append(upper_row)
        for earlier_block in blocks_sorted:
            if earlier_block is block:
                break
            earlier_block["final_rows"] = [r + n for r in earlier_block.get("final_rows", [])]

    # False (not just default) when every "new" row turned out to be
    # unplaceable (ascending chronological-order check moved them all
    # to review) — nothing was actually written in that case.
    report["written"] = len(blocks) > 0

    # --- Validation (per explicit request): re-read the Check column
    # for every newly written row, plus each block's upper anchor row
    # (to confirm the reconnect itself, not just the new rows in
    # isolation) — Check = Balance (AI) - BALANCE, so a non-zero value
    # here means either the row landed in the wrong chronological
    # position or a credit/debit sign is backwards. ---
    if "Check" in main_header:
        check_letter = col_letter(main_header, "Check")
        rows_to_verify: set[int] = set()
        for block in blocks:
            rows_to_verify.update(block.get("final_rows", []))
        sorted_rows = sorted(rows_to_verify)
        if sorted_rows:
            # One batch_get for every row being verified, instead of one
            # .acell() round-trip each.
            ranges = [f"{check_letter}{r}" for r in sorted_rows]
            matrices = call_with_retry(
                lambda: main_ws.batch_get(ranges, value_render_option=gs_utils.ValueRenderOption.unformatted)
            )
            for row_number, matrix in zip(sorted_rows, matrices):
                value = matrix[0][0] if matrix and matrix[0] else None
                ok = (value == "" or value is None) or (isinstance(value, (int, float)) and abs(value) < 0.01)
                report["validation"].append({"row": row_number, "check_value": value, "ok": ok})

    return report
