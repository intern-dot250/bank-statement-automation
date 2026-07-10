"""Phase 2D: Validate the complete reporting pipeline (Master -> Summary ->
Final Report).

Reads all three worksheets and cross-checks that:
  - Master's raw data is internally consistent (required columns present,
    no blank Head/Narration on classified rows, Deposits/Withdrawals
    numeric).
  - Summary's totals match what Master's data actually sums to.
  - Final Report's totals match Summary's totals.

Prints a clean PASS/FAIL report per check plus one overall status line.
Never modifies any worksheet — read-only.

Not integrated into run_pipeline.py or any other module — this is a
standalone script for Phase 2D, run independently via:

    py validate_report.py
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import gspread

from generate_final_report import FINAL_REPORT_WORKSHEET_NAME, TABLE_HEADER
from generate_final_report import SummaryData, parse_summary_data
from generate_summary import SUMMARY_WORKSHEET_NAME
from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
    load_combined_account_values,
)

LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("validate_report")

REQUIRED_MASTER_COLUMNS = [
    "TXN DATE",
    "VALUE DATE",
    "DESCRIPTION",
    "CREDITS",
    "DEBITS",
    "BALANCE",
    "HEAD",
    "NARRATION",
]

# Currency comparisons tolerate small floating-point/rounding drift.
FLOAT_TOLERANCE = 0.01


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class ValidationCheck:
    """A single validation result."""

    section: str
    name: str
    passed: bool
    expected: Any = None
    actual: Any = None
    detail: str = ""


@dataclass
class MasterStats:
    """Data extracted from the Master worksheet for validation."""

    header_ok: bool = True
    missing_columns: list[str] = field(default_factory=list)
    total_rows: int = 0
    classified_count: int = 0  # rows with non-blank Head (mirrors generate_summary.py's rule)
    blank_head_rows: list[int] = field(default_factory=list)
    blank_narration_rows: list[int] = field(default_factory=list)
    non_numeric_deposits_rows: list[int] = field(default_factory=list)
    non_numeric_withdrawals_rows: list[int] = field(default_factory=list)
    total_deposits: float = 0.0
    total_withdrawals: float = 0.0


@dataclass
class FinalReportData:
    """Parsed contents of the Final Report worksheet."""

    head_rows: list[tuple[str, float, float, int]] = field(default_factory=list)
    total_credits: float = 0.0
    total_debits: float = 0.0
    net_collection: float = 0.0
    found_totals: bool = False


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_row_blank(row: list[str]) -> bool:
    return not any(cell.strip() for cell in row)


def _is_numeric(value: str) -> bool:
    """True if value is blank (allowed — a row has either a deposit OR a
    withdrawal, not both) or parses cleanly as a float (including
    currency-formatted values — see _to_float())."""
    raw = value.strip()
    if not raw:
        return True
    stripped = raw[1:-1] if raw.startswith("(") and raw.endswith(")") else raw
    cleaned = _NON_NUMERIC_RE.sub("", stripped).replace("-", "")
    if not cleaned:
        return False
    try:
        float(cleaned)
        return True
    except ValueError:
        return False


# Matches anything that is NOT a digit, a decimal point, or a minus sign —
# i.e. currency symbols (₹, $, €, £, ...), thousands separators (commas,
# regardless of Western "753,777" or Indian "7,53,777" grouping), and
# stray whitespace. Kept as a module-level constant so it's compiled once.
_NON_NUMERIC_RE = re.compile(r"[^\d.\-]")


def _to_float(value: str) -> float:
    """Parse a raw or currency-formatted sheet cell into a float.

    Handles plain numbers ("753777.00"), thousands separators in either
    grouping style ("753,777.00" / "7,53,777.00"), a leading currency
    symbol ("₹753,777.00"), a minus sign before OR after the currency
    symbol ("-₹35,329.00", "₹-35,329.00"), and accounting-style negative
    parentheses ("(35,329.00)"). Returns 0.0 for blank or unparseable
    input — never raises.
    """
    raw = value.strip()
    if not raw:
        return 0.0

    negative = False
    if raw.startswith("(") and raw.endswith(")"):
        negative = True
        raw = raw[1:-1]

    if "-" in raw:
        negative = True

    cleaned = _NON_NUMERIC_RE.sub("", raw).replace("-", "")
    if not cleaned:
        return 0.0

    try:
        result = float(cleaned)
    except ValueError:
        return 0.0

    return -result if negative else result


def _floats_equal(a: float, b: float, tolerance: float = FLOAT_TOLERANCE) -> bool:
    return abs(a - b) <= tolerance


def _check(
    checks: list[ValidationCheck],
    section: str,
    name: str,
    passed: bool,
    expected: Any = None,
    actual: Any = None,
    detail: str = "",
) -> None:
    checks.append(ValidationCheck(section, name, passed, expected, actual, detail))


# ---------------------------------------------------------------------------
# Worksheet access (reuses upload_to_sheets.py's auth — no duplication)
# ---------------------------------------------------------------------------

def open_worksheet_safely(
    spreadsheet: gspread.Spreadsheet, worksheet_name: str,
) -> gspread.Worksheet | None:
    """Open a worksheet by name, returning None (and logging) if missing
    rather than raising — so one missing sheet doesn't abort the whole
    validation run."""
    try:
        return spreadsheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        log.error("Worksheet %r not found.", worksheet_name)
        return None


# ---------------------------------------------------------------------------
# MASTER: extraction + validation
# ---------------------------------------------------------------------------

def load_master_stats(master_values: list[list[str]]) -> MasterStats:
    """Extract validation-relevant statistics from the Master worksheet.

    "Classified" here means a row has a non-blank Head — this deliberately
    mirrors generate_summary.py's own counting rule (which does not
    require Narration to be non-blank), so the cross-check against
    Summary's transaction count compares like with like. Blank Narration
    on an otherwise-classified row is still flagged separately as its own
    data-quality check below.
    """
    stats = MasterStats()

    if not master_values:
        stats.header_ok = False
        stats.missing_columns = list(REQUIRED_MASTER_COLUMNS)
        return stats

    header_row = master_values[0]
    header_index = {name.strip(): i for i, name in enumerate(header_row)}

    stats.missing_columns = [c for c in REQUIRED_MASTER_COLUMNS if c not in header_index]
    stats.header_ok = not stats.missing_columns

    if not stats.header_ok:
        return stats

    def cell(row: list[str], column: str) -> str:
        index = header_index[column]
        return row[index].strip() if index < len(row) else ""

    log.debug("Master worksheet: %d physical row(s) read (including header).", len(master_values))

    for row_offset, row in enumerate(master_values[1:]):
        sheet_row_number = row_offset + 2  # +1 header, +1 for 1-based indexing

        if _is_row_blank(row):
            log.debug("Master row %d: raw=%r -> BLANK ROW, skipped.", sheet_row_number, row)
            continue

        stats.total_rows += 1

        head = cell(row, "HEAD")
        narration = cell(row, "NARRATION")
        deposits_raw = cell(row, "CREDITS")
        withdrawals_raw = cell(row, "DEBITS")

        if not head:
            stats.blank_head_rows.append(sheet_row_number)
            log.debug(
                "Master row %d: raw=%r -> BLANK HEAD (not classified). "
                "Deposits=%r Withdrawals=%r",
                sheet_row_number, row, deposits_raw, withdrawals_raw,
            )
            continue  # unclassified row — excluded from totals, matching generate_summary.py

        stats.classified_count += 1
        log.debug(
            "Master row %d: raw=%r -> CLASSIFIED. Head=%r Narration=%r Deposits=%r Withdrawals=%r",
            sheet_row_number, row, head, narration, deposits_raw, withdrawals_raw,
        )

        if not narration:
            stats.blank_narration_rows.append(sheet_row_number)

        if not _is_numeric(deposits_raw):
            stats.non_numeric_deposits_rows.append(sheet_row_number)
        if not _is_numeric(withdrawals_raw):
            stats.non_numeric_withdrawals_rows.append(sheet_row_number)

        stats.total_deposits += _to_float(deposits_raw)
        stats.total_withdrawals += _to_float(withdrawals_raw)

    log.info(
        "Master parsing summary: %d physical row(s) read, %d non-blank, %d classified "
        "(non-blank Head), %d blank-Head.",
        len(master_values), stats.total_rows, stats.classified_count, len(stats.blank_head_rows),
    )

    return stats


def validate_master(stats: MasterStats) -> list[ValidationCheck]:
    """Run all MASTER-section validation checks."""
    checks: list[ValidationCheck] = []

    _check(
        checks, "MASTER", "Required columns exist",
        passed=stats.header_ok,
        expected=REQUIRED_MASTER_COLUMNS,
        actual=f"missing: {stats.missing_columns}" if stats.missing_columns else "all present",
    )

    if not stats.header_ok:
        # Every other Master check depends on the header being valid.
        _check(checks, "MASTER", "No blank Head values", False, "0 blank", "header invalid — skipped")
        _check(checks, "MASTER", "No blank Narration values", False, "0 blank", "header invalid — skipped")
        _check(checks, "MASTER", "Credits are numeric", False, "all numeric", "header invalid — skipped")
        _check(checks, "MASTER", "Debits are numeric", False, "all numeric", "header invalid — skipped")
        _check(checks, "MASTER", "Classified row count", False, expected=">= 0", actual="header invalid — skipped")
        return checks

    _check(
        checks, "MASTER", "No blank Head values",
        passed=len(stats.blank_head_rows) == 0,
        expected="0 blank Head rows",
        actual=f"{len(stats.blank_head_rows)} blank (rows: {stats.blank_head_rows[:10]})",
    )

    _check(
        checks, "MASTER", "No blank Narration values",
        passed=len(stats.blank_narration_rows) == 0,
        expected="0 blank Narration rows",
        actual=f"{len(stats.blank_narration_rows)} blank (rows: {stats.blank_narration_rows[:10]})",
    )

    _check(
        checks, "MASTER", "Credits are numeric",
        passed=len(stats.non_numeric_deposits_rows) == 0,
        expected="0 non-numeric Credits",
        actual=f"{len(stats.non_numeric_deposits_rows)} non-numeric (rows: {stats.non_numeric_deposits_rows[:10]})",
    )

    _check(
        checks, "MASTER", "Debits are numeric",
        passed=len(stats.non_numeric_withdrawals_rows) == 0,
        expected="0 non-numeric Debits",
        actual=f"{len(stats.non_numeric_withdrawals_rows)} non-numeric (rows: {stats.non_numeric_withdrawals_rows[:10]})",
    )

    _check(
        checks, "MASTER", "Classified row count",
        passed=True,
        expected=">= 0",
        actual=f"{stats.classified_count} classified row(s) out of {stats.total_rows} total",
    )

    return checks


# ---------------------------------------------------------------------------
# SUMMARY: validation against Master
# ---------------------------------------------------------------------------

def validate_summary(master_stats: MasterStats, summary_data: SummaryData | None) -> list[ValidationCheck]:
    """Run all SUMMARY-section validation checks, cross-referenced against Master."""
    checks: list[ValidationCheck] = []

    if summary_data is None:
        for name in (
            "Total Credits equals Master Credits",
            "Total Debits equals Master Debits",
            "Transaction count equals classified Master rows",
            "Net Collection = Credits - Debits",
        ):
            _check(checks, "SUMMARY", name, False, detail="Summary worksheet not found")
        return checks

    _check(
        checks, "SUMMARY", "Total Credits equals Master Credits",
        passed=_floats_equal(summary_data.total_credits, master_stats.total_deposits),
        expected=round(master_stats.total_deposits, 2),
        actual=round(summary_data.total_credits, 2),
    )

    _check(
        checks, "SUMMARY", "Total Debits equals Master Debits",
        passed=_floats_equal(summary_data.total_debits, master_stats.total_withdrawals),
        expected=round(master_stats.total_withdrawals, 2),
        actual=round(summary_data.total_debits, 2),
    )

    _check(
        checks, "SUMMARY", "Transaction count equals classified Master rows",
        passed=summary_data.total_transactions == master_stats.classified_count,
        expected=master_stats.classified_count,
        actual=summary_data.total_transactions,
    )

    expected_net = summary_data.total_credits - summary_data.total_debits
    _check(
        checks, "SUMMARY", "Net Collection = Credits - Debits",
        passed=_floats_equal(summary_data.net_collection, expected_net),
        expected=round(expected_net, 2),
        actual=round(summary_data.net_collection, 2),
    )

    return checks


# ---------------------------------------------------------------------------
# FINAL REPORT: extraction + validation against Summary
# ---------------------------------------------------------------------------

def load_final_report_data(final_report_values: list[list[str]]) -> FinalReportData:
    """Parse the Final Report worksheet's title/table/totals layout.

    Locates the table header row by exact match against TABLE_HEADER
    (imported from generate_final_report.py) rather than assuming a
    fixed row number, and locates the totals by their row labels.
    """
    data = FinalReportData()

    header_row_index = None
    for i, row in enumerate(final_report_values):
        if [c.strip() for c in row[: len(TABLE_HEADER)]] == TABLE_HEADER:
            header_row_index = i
            break

    if header_row_index is None:
        log.error("Could not locate the table header row in the Final Report worksheet.")
        return data

    for row in final_report_values[header_row_index + 1:]:
        if _is_row_blank(row):
            continue

        label = row[0].strip()
        if not label:
            continue

        if label == "Total Credits":
            data.total_credits = _to_float(row[1]) if len(row) > 1 else 0.0
            data.found_totals = True
        elif label == "Total Debits":
            data.total_debits = _to_float(row[1]) if len(row) > 1 else 0.0
        elif label == "Net Collection":
            data.net_collection = _to_float(row[1]) if len(row) > 1 else 0.0
        else:
            credits_val = _to_float(row[1]) if len(row) > 1 else 0.0
            debits_val = _to_float(row[2]) if len(row) > 2 else 0.0
            try:
                transactions_val = int(float(row[3])) if len(row) > 3 and row[3].strip() else 0
            except ValueError:
                transactions_val = 0
            data.head_rows.append((label, credits_val, debits_val, transactions_val))

    return data


def validate_final_report(
    summary_data: SummaryData | None,
    final_report_data: FinalReportData | None,
) -> list[ValidationCheck]:
    """Run all FINAL REPORT-section validation checks, cross-referenced against Summary."""
    checks: list[ValidationCheck] = []

    if final_report_data is None:
        for name in (
            "Total Credits equals Summary",
            "Total Debits equals Summary",
            "Net Collection equals Summary",
            "Transaction counts equal Summary",
        ):
            _check(checks, "FINAL REPORT", name, False, detail="Final Report worksheet not found")
        return checks

    if summary_data is None:
        for name in (
            "Total Credits equals Summary",
            "Total Debits equals Summary",
            "Net Collection equals Summary",
            "Transaction counts equal Summary",
        ):
            _check(checks, "FINAL REPORT", name, False, detail="Summary worksheet not found — cannot cross-check")
        return checks

    if not final_report_data.found_totals:
        for name in ("Total Credits equals Summary", "Total Debits equals Summary", "Net Collection equals Summary"):
            _check(checks, "FINAL REPORT", name, False, detail="Totals section not found in Final Report")
    else:
        _check(
            checks, "FINAL REPORT", "Total Credits equals Summary",
            passed=_floats_equal(final_report_data.total_credits, summary_data.total_credits),
            expected=round(summary_data.total_credits, 2),
            actual=round(final_report_data.total_credits, 2),
        )
        _check(
            checks, "FINAL REPORT", "Total Debits equals Summary",
            passed=_floats_equal(final_report_data.total_debits, summary_data.total_debits),
            expected=round(summary_data.total_debits, 2),
            actual=round(final_report_data.total_debits, 2),
        )
        _check(
            checks, "FINAL REPORT", "Net Collection equals Summary",
            passed=_floats_equal(final_report_data.net_collection, summary_data.net_collection),
            expected=round(summary_data.net_collection, 2),
            actual=round(final_report_data.net_collection, 2),
        )

    summary_by_head = {h: (c, d, t) for h, c, d, t in summary_data.head_rows}
    report_by_head = {h: (c, d, t) for h, c, d, t in final_report_data.head_rows}

    mismatches = []
    for head, (s_credits, s_debits, s_txn) in summary_by_head.items():
        r = report_by_head.get(head)
        if r is None:
            mismatches.append(f"{head}: missing from Final Report")
            continue

        r_credits, r_debits, r_txn = r
        field_diffs = []
        if r_txn != s_txn:
            field_diffs.append(f"Transactions: Summary={s_txn} FinalReport={r_txn}")
        if not _floats_equal(r_credits, s_credits):
            field_diffs.append(f"Credits: Summary={s_credits:.2f} FinalReport={r_credits:.2f}")
        if not _floats_equal(r_debits, s_debits):
            field_diffs.append(f"Debits: Summary={s_debits:.2f} FinalReport={r_debits:.2f}")

        if field_diffs:
            mismatches.append(f"{head}: " + "; ".join(field_diffs))

    extra_heads = set(report_by_head) - set(summary_by_head)
    for head in extra_heads:
        mismatches.append(f"{head}: present in Final Report but not in Summary")

    _check(
        checks, "FINAL REPORT", "Transaction counts equal Summary",
        passed=len(mismatches) == 0,
        expected=f"{len(summary_by_head)} Head(s) matching Summary",
        actual="all match" if not mismatches else f"{len(mismatches)} mismatch(es): {mismatches[:10]}",
    )

    return checks


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

def print_report(checks: list[ValidationCheck]) -> bool:
    """Print the full validation report to the terminal.

    Returns:
        True if every check passed, False otherwise.
    """
    print("=" * 78)
    print("VALIDATION REPORT — Master -> Summary -> Final Report")
    print("=" * 78)

    current_section = None
    for check in checks:
        if check.section != current_section:
            current_section = check.section
            print(f"\n[{current_section}]")

        status = "PASS" if check.passed else "FAIL"
        print(f"  [{status}] {check.name}")
        if not check.passed:
            print(f"         Expected : {check.expected}")
            print(f"         Actual   : {check.actual}")
            if check.detail:
                print(f"         Detail   : {check.detail}")

    overall_passed = all(check.passed for check in checks)

    print("\n" + "=" * 78)
    print(f"OVERALL STATUS : {'PASS' if overall_passed else 'FAIL'}")
    print("=" * 78)

    return overall_passed


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def validate_report(
    credentials_path: Path,
    sheet_id: str = MASTER_SHEET_ID,
    summary_worksheet_name: str = SUMMARY_WORKSHEET_NAME,
    final_report_worksheet_name: str = FINAL_REPORT_WORKSHEET_NAME,
) -> bool:
    """Run the full (combined account data) -> Summary -> Final Report
    validation pipeline.

    Args:
        credentials_path: Path to the Google service-account credentials JSON.
        sheet_id: Spreadsheet ID containing all the account tabs plus
            Summary/Final Report.
        summary_worksheet_name: Worksheet/tab name of the Summary report.
        final_report_worksheet_name: Worksheet/tab name of the Final Report.

    Returns:
        True if every validation check passed, False otherwise. Never
        raises for missing worksheets — those are reported as FAIL
        checks instead, so the report always prints in full.
    """
    # Credential resolution (file vs GOOGLE_CREDENTIALS_JSON env var fallback)
    # is handled entirely inside get_gspread_client() — no upfront existence
    # check here, since that would bypass the env var fallback it supports.
    client = get_gspread_client(credentials_path)
    spreadsheet = client.open_by_key(sheet_id)

    master_values = load_combined_account_values(spreadsheet)
    summary_ws = open_worksheet_safely(spreadsheet, summary_worksheet_name)
    final_report_ws = open_worksheet_safely(spreadsheet, final_report_worksheet_name)

    all_checks: list[ValidationCheck] = []

    if not master_values:
        _check(all_checks, "MASTER", "Required columns exist", False, detail="No account worksheets with data found")
        master_stats = MasterStats(header_ok=False)
    else:
        master_stats = load_master_stats(master_values)
        all_checks.extend(validate_master(master_stats))

    summary_data = None
    if summary_ws is not None:
        try:
            summary_data = parse_summary_data(summary_ws.get_all_values())
        except Exception as exc:
            log.error("Could not parse Summary worksheet: %s", exc)
    all_checks.extend(validate_summary(master_stats, summary_data))

    final_report_data = None
    if final_report_ws is not None:
        try:
            # UNFORMATTED_VALUE: generate_final_report.py applies CURRENCY
            # cell formatting (e.g. "₹#,##0.00"), and get_all_values()
            # defaults to returning the DISPLAYED text ("₹7,53,777.00"),
            # not the underlying number. Reading unformatted values here
            # avoids having to strip currency symbols/locale-specific
            # digit grouping ourselves. UNFORMATTED_VALUE returns native
            # JSON types (numbers as int/float, not str), so normalize
            # every cell back to a string before handing it to
            # load_final_report_data(), which expects str.strip()-able cells.
            raw_final_report_values = final_report_ws.get_all_values(
                value_render_option="UNFORMATTED_VALUE"
            )
            log.debug("Final Report raw values (UNFORMATTED_VALUE), %d row(s):", len(raw_final_report_values))
            for i, row in enumerate(raw_final_report_values, start=1):
                log.debug("  Final Report row %d (raw): %s", i, row)

            normalized_final_report_values = [
                ["" if cell is None else str(cell) for cell in row]
                for row in raw_final_report_values
            ]
            for i, row in enumerate(normalized_final_report_values, start=1):
                log.debug("  Final Report row %d (normalized to str): %s", i, row)

            final_report_data = load_final_report_data(normalized_final_report_values)
            log.debug(
                "Final Report parsed: total_credits=%r total_debits=%r net_collection=%r head_rows=%s",
                final_report_data.total_credits, final_report_data.total_debits,
                final_report_data.net_collection, final_report_data.head_rows,
            )
        except Exception as exc:
            log.error("Could not parse Final Report worksheet: %s", exc)
    all_checks.extend(validate_final_report(summary_data, final_report_data))

    return print_report(all_checks)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate the Master -> Summary -> Final Report reporting pipeline."
    )

    parser.add_argument("-c", "--credentials", type=Path, default=DEFAULT_CREDENTIALS)
    parser.add_argument("--sheet-id", default=MASTER_SHEET_ID, help="Override the spreadsheet ID.")
    parser.add_argument("--summary-worksheet-name", default=SUMMARY_WORKSHEET_NAME)
    parser.add_argument("--final-report-worksheet-name", default=FINAL_REPORT_WORKSHEET_NAME)
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable DEBUG logging (row-by-row Master trace, raw Final Report values, parsed totals).",
    )

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    if args.verbose:
        logging.getLogger("validate_report").setLevel(logging.DEBUG)

    try:
        passed = validate_report(
            credentials_path=args.credentials,
            sheet_id=args.sheet_id,
            summary_worksheet_name=args.summary_worksheet_name,
            final_report_worksheet_name=args.final_report_worksheet_name,
        )
    except Exception as exc:
        log.error("Validation could not be completed: %s", exc)
        return 1

    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
