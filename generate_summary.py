"""Phase 2B: Generate a per-Head financial Summary combined across every
account's Google Sheet tab.

Reads every account worksheet (the tabs upload_to_sheets.py writes to and
classify_transactions.py classifies — one per bank account, e.g.
"YES BANK - 2477"), aggregates Credits (Deposits), Debits (Withdrawals),
and transaction counts per Head across ALL of them, and writes the
combined result into a separate "Summary" worksheet in the same
spreadsheet.

Not integrated into run_pipeline.py or any other module — this is a
standalone script for Phase 2B, run independently via:

    py generate_summary.py
"""

from __future__ import annotations

import argparse
import logging
import sys
from dataclasses import dataclass, field
from pathlib import Path

import gspread

from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
    load_combined_account_values,
)

LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("generate_summary")

SUMMARY_WORKSHEET_NAME = "Summary"

SUMMARY_HEADER = ["Head", "Credits", "Debits", "Transactions"]


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class HeadTotals:
    """Running totals for one Head."""

    credits: float = 0.0
    debits: float = 0.0
    transactions: int = 0


@dataclass
class SummaryResult:
    """Full summary computation result."""

    per_head: dict[str, HeadTotals] = field(default_factory=dict)
    total_credits: float = 0.0
    total_debits: float = 0.0
    net_collection: float = 0.0
    rows_considered: int = 0
    rows_skipped: int = 0


# ---------------------------------------------------------------------------
# Worksheet access (reuses upload_to_sheets.py's auth — no duplication)
# ---------------------------------------------------------------------------

def get_or_create_summary_worksheet(
    spreadsheet: gspread.Spreadsheet,
    worksheet_name: str = SUMMARY_WORKSHEET_NAME,
) -> gspread.Worksheet:
    """Return the Summary worksheet, creating it if it doesn't exist.

    If it already exists, its contents are cleared (the worksheet itself
    is never deleted) so the summary can be safely regenerated.
    """
    existing_titles = [ws.title for ws in spreadsheet.worksheets()]

    if worksheet_name in existing_titles:
        log.info("Summary worksheet exists. Clearing contents: %s", worksheet_name)
        worksheet = spreadsheet.worksheet(worksheet_name)
        worksheet.clear()
    else:
        log.info("Creating Summary worksheet: %s", worksheet_name)
        worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="200", cols="10")

    return worksheet


# ---------------------------------------------------------------------------
# Row helpers
# ---------------------------------------------------------------------------

def _get_cell(row: list[str], header_index: dict[str, int], column_name: str) -> str:
    """Safely read a cell value by column name, tolerating short/ragged rows."""
    index = header_index.get(column_name)
    if index is None or index >= len(row):
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


def _is_row_empty(row: list[str]) -> bool:
    """A row is empty if every cell is blank."""
    return all(cell.strip() == "" for cell in row)


# ---------------------------------------------------------------------------
# Summary computation
# ---------------------------------------------------------------------------

def compute_summary(all_values: list[list[str]]) -> SummaryResult:
    """Compute per-Head Credits/Debits/Transaction totals from raw sheet values.

    Args:
        all_values: Full worksheet contents as returned by
            worksheet.get_all_values(), including the header row.

    Rules:
        - Blank rows are ignored.
        - Rows without a Head value are ignored (unclassified rows are
          not counted in the summary).
        - Missing/malformed Deposits or Withdrawals values are treated
          as 0.0, never raising an error.

    Returns:
        A SummaryResult with per-Head totals, overall totals, and Net
        Collection (Total Credits - Total Debits).
    """
    if not all_values:
        log.warning("Worksheet is empty — nothing to summarize.")
        return SummaryResult()

    header_row = all_values[0]
    header_index = {name.strip(): i for i, name in enumerate(header_row)}

    required_columns = ["HEAD", "CREDITS", "DEBITS"]
    missing = [c for c in required_columns if c not in header_index]
    if missing:
        raise ValueError(f"Master sheet is missing required column(s): {missing}")

    result = SummaryResult()

    for row in all_values[1:]:
        if _is_row_empty(row):
            result.rows_skipped += 1
            continue

        head = _get_cell(row, header_index, "HEAD")
        if not head:
            result.rows_skipped += 1
            continue

        deposits = _to_float(_get_cell(row, header_index, "CREDITS"))
        withdrawals = _to_float(_get_cell(row, header_index, "DEBITS"))

        totals = result.per_head.setdefault(head, HeadTotals())
        totals.credits += deposits
        totals.debits += withdrawals
        totals.transactions += 1

        result.total_credits += deposits
        result.total_debits += withdrawals
        result.rows_considered += 1

    result.net_collection = result.total_credits - result.total_debits

    log.info(
        "Summary computed: %d Head(s), %d row(s) considered, %d row(s) skipped.",
        len(result.per_head), result.rows_considered, result.rows_skipped,
    )

    return result


# ---------------------------------------------------------------------------
# Sheet output
# ---------------------------------------------------------------------------

def build_summary_rows(result: SummaryResult) -> list[list[str]]:
    """Build the 2D grid of values to write into the Summary worksheet.

    Layout:
        Head | Credits | Debits | Transactions
        <one row per unique Head, sorted alphabetically>
        (blank separator row)
        TOTAL | total_credits | total_debits | total_transactions
        Net Collection | net_collection
    """
    rows: list[list[str]] = [SUMMARY_HEADER]

    for head in sorted(result.per_head.keys()):
        totals = result.per_head[head]
        rows.append([
            head,
            f"{totals.credits:.2f}",
            f"{totals.debits:.2f}",
            str(totals.transactions),
        ])

    rows.append(["", "", "", ""])
    rows.append([
        "TOTAL",
        f"{result.total_credits:.2f}",
        f"{result.total_debits:.2f}",
        str(result.rows_considered),
    ])
    rows.append(["Net Collection", f"{result.net_collection:.2f}", "", ""])

    return rows


def write_summary(worksheet: gspread.Worksheet, result: SummaryResult) -> None:
    """Write the computed summary into the (already cleared) Summary worksheet."""
    rows = build_summary_rows(result)
    worksheet.update(range_name="A1", values=rows, value_input_option="RAW")
    log.info("Wrote %d row(s) to the Summary worksheet.", len(rows))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def generate_summary(
    credentials_path: Path,
    sheet_id: str = MASTER_SHEET_ID,
    summary_worksheet_name: str = SUMMARY_WORKSHEET_NAME,
) -> SummaryResult:
    """Generate the per-Head Summary, combined across every account's
    worksheet tab, and write it to the Summary worksheet.

    Args:
        credentials_path: Path to the Google service-account credentials JSON.
        sheet_id: Spreadsheet ID containing all the account tabs.
        summary_worksheet_name: Worksheet/tab name to write the summary into.

    Returns:
        The computed SummaryResult.
    """
    # Credential resolution (file vs GOOGLE_CREDENTIALS_JSON env var fallback)
    # is handled entirely inside get_gspread_client() — no upfront existence
    # check here, since that would bypass the env var fallback it supports.
    client = get_gspread_client(credentials_path)
    spreadsheet = client.open_by_key(sheet_id)

    all_values = load_combined_account_values(spreadsheet)
    result = compute_summary(all_values)

    summary_worksheet = get_or_create_summary_worksheet(spreadsheet, summary_worksheet_name)
    write_summary(summary_worksheet, result)

    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a per-Head Summary from the classified master Google Sheet."
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
        "--summary-worksheet-name",
        default=SUMMARY_WORKSHEET_NAME,
        help="Override the Summary worksheet/tab name.",
    )

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        result = generate_summary(
            credentials_path=args.credentials,
            sheet_id=args.sheet_id,
            summary_worksheet_name=args.summary_worksheet_name,
        )
        log.info(
            "Summary generation complete. Heads=%d TotalCredits=%.2f TotalDebits=%.2f NetCollection=%.2f",
            len(result.per_head), result.total_credits, result.total_debits, result.net_collection,
        )
    except Exception as exc:
        log.exception(exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
