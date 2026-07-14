"""Phase 2C: Generate a formatted Final Report from the Summary worksheet.

Reads the "Summary" worksheet produced by generate_summary.py (Head,
Credits, Debits, Transactions, plus TOTAL and Net Collection rows), and
writes a clean, formatted business report into a separate "Final Report"
worksheet in the same spreadsheet — with a title, generation timestamp,
bold headers, currency-formatted numbers, and auto-sized columns.

Not integrated into run_pipeline.py or any other module — this is a
standalone script for Phase 2C, run independently via:

    py generate_final_report.py
"""

from __future__ import annotations

import argparse
import logging
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import gspread

from generate_summary import SUMMARY_WORKSHEET_NAME
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client

LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("generate_final_report")

FINAL_REPORT_WORKSHEET_NAME = "Final Report"
REPORT_TITLE = "Bank Statement Automation — Final Report"

TABLE_HEADER = ["Head", "Credits", "Debits", "Transactions"]

CURRENCY_NUMBER_FORMAT = {"type": "CURRENCY", "pattern": "₹#,##0.00"}


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class SummaryData:
    """Parsed contents of the Summary worksheet."""

    head_rows: list[tuple[str, float, float, int]] = field(default_factory=list)
    total_credits: float = 0.0
    total_debits: float = 0.0
    total_transactions: int = 0
    net_collection: float = 0.0


# ---------------------------------------------------------------------------
# Worksheet access (reuses upload_to_sheets.py's auth — no duplication)
# ---------------------------------------------------------------------------

def open_summary_worksheet(
    spreadsheet: gspread.Spreadsheet,
    summary_worksheet_name: str,
) -> gspread.Worksheet:
    """Open the existing Summary worksheet.

    Raises:
        gspread.exceptions.WorksheetNotFound: If the Summary worksheet
            does not exist. This script only reads it; it never creates
            it — that is generate_summary.py's responsibility.
    """
    return spreadsheet.worksheet(summary_worksheet_name)


def get_or_create_final_report_worksheet(
    spreadsheet: gspread.Spreadsheet,
    worksheet_name: str = FINAL_REPORT_WORKSHEET_NAME,
) -> gspread.Worksheet:
    """Return the Final Report worksheet, creating it if it doesn't exist.

    If it already exists, its contents (and any prior formatting) are
    cleared — the worksheet itself is never deleted — so the report can
    be safely regenerated.
    """
    existing_titles = [ws.title for ws in spreadsheet.worksheets()]

    if worksheet_name in existing_titles:
        log.info("Final Report worksheet exists. Clearing contents: %s", worksheet_name)
        worksheet = spreadsheet.worksheet(worksheet_name)
        worksheet.clear()
        _clear_formatting(spreadsheet, worksheet)
    else:
        log.info("Creating Final Report worksheet: %s", worksheet_name)
        worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="200", cols="10")

    return worksheet


def _clear_formatting(spreadsheet: gspread.Spreadsheet, worksheet: gspread.Worksheet) -> None:
    """Reset any previously-applied cell formatting before rewriting the report."""
    try:
        spreadsheet.batch_update({
            "requests": [{
                "updateCells": {
                    "range": {"sheetId": worksheet.id},
                    "fields": "userEnteredFormat",
                }
            }]
        })
    except Exception as exc:
        log.warning("Could not clear prior formatting on %s: %s", worksheet.title, exc)


# ---------------------------------------------------------------------------
# Parsing the Summary worksheet
# ---------------------------------------------------------------------------

def _to_float(value: str) -> float:
    """Safely convert a raw sheet cell value to a float, defaulting to 0.0."""
    cleaned = value.replace(",", "").strip()
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _to_int(value: str) -> int:
    """Safely convert a raw sheet cell value to an int, defaulting to 0."""
    try:
        return int(float(value.strip())) if value.strip() else 0
    except ValueError:
        return 0


def parse_summary_data(summary_values: list[list[str]]) -> SummaryData:
    """Parse the raw Summary worksheet grid into structured data.

    Locates the "TOTAL" and "Net Collection" rows by their label in the
    first column rather than assuming fixed row positions, so this stays
    correct even if generate_summary.py's exact row count changes.

    Args:
        summary_values: Full Summary worksheet contents, as returned by
            worksheet.get_all_values(), including the header row.

    Returns:
        A SummaryData with per-Head rows and overall totals. Missing or
        malformed numeric values are treated as 0, never raising.
    """
    data = SummaryData()

    if not summary_values or len(summary_values) < 2:
        log.warning("Summary worksheet is empty or has no data rows.")
        return data

    for row in summary_values[1:]:
        if not row or not any(cell.strip() for cell in row):
            continue

        label = row[0].strip()
        if not label:
            continue

        if label == "TOTAL":
            data.total_credits = _to_float(row[1]) if len(row) > 1 else 0.0
            data.total_debits = _to_float(row[2]) if len(row) > 2 else 0.0
            data.total_transactions = _to_int(row[3]) if len(row) > 3 else 0
        elif label == "Net Collection":
            data.net_collection = _to_float(row[1]) if len(row) > 1 else 0.0
        else:
            credits_val = _to_float(row[1]) if len(row) > 1 else 0.0
            debits_val = _to_float(row[2]) if len(row) > 2 else 0.0
            transactions_val = _to_int(row[3]) if len(row) > 3 else 0
            data.head_rows.append((label, credits_val, debits_val, transactions_val))

    log.info(
        "Parsed Summary: %d Head row(s), TotalCredits=%.2f TotalDebits=%.2f NetCollection=%.2f",
        len(data.head_rows), data.total_credits, data.total_debits, data.net_collection,
    )

    return data


# ---------------------------------------------------------------------------
# Report layout
# ---------------------------------------------------------------------------

def build_report_rows(data: SummaryData) -> list[list[str | float | int]]:
    """Build the 2D grid of values to write into the Final Report worksheet.

    Layout:
        Row 1: Report title
        Row 2: Generated date & time
        Row 3: (blank)
        Row 4: Table header (Head | Credits | Debits | Transactions)
        Row 5..: one row per Head
        Row N:   (blank)
        Row N+1: Total Credits
        Row N+2: Total Debits
        Row N+3: Net Collection
    """
    generated_at = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

    rows: list[list[str | float | int]] = [
        [REPORT_TITLE],
        [f"Generated: {generated_at}"],
        [],
        TABLE_HEADER,
    ]

    for head, credits_val, debits_val, transactions_val in data.head_rows:
        rows.append([head, credits_val, debits_val, transactions_val])

    rows.append([])
    rows.append(["Total Credits", data.total_credits])
    rows.append(["Total Debits", data.total_debits])
    rows.append(["Net Collection", data.net_collection])

    return rows


def write_report(worksheet: gspread.Worksheet, rows: list[list]) -> None:
    """Write the report grid into the (already cleared) Final Report worksheet."""
    worksheet.update(range_name="A1", values=rows, value_input_option="USER_ENTERED")
    log.info("Wrote %d row(s) to the Final Report worksheet.", len(rows))


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------

def apply_formatting(
    spreadsheet: gspread.Spreadsheet,
    worksheet: gspread.Worksheet,
    data: SummaryData,
) -> None:
    """Apply bold headers, currency number formatting, and column auto-sizing.

    Formatting failures are logged but never raised — a report with
    correct values and imperfect formatting is still useful; a crash
    here should not discard already-written data.
    """
    header_row_index = 4  # 1-based: Row 4 is the table header (see build_report_rows)
    last_head_row_index = header_row_index + len(data.head_rows)
    totals_start_row_index = last_head_row_index + 2  # skip the blank separator row

    try:
        worksheet.format("A1", {"textFormat": {"bold": True, "fontSize": 14}})
        worksheet.format("A2", {"textFormat": {"italic": True}})
        worksheet.format(
            f"A{header_row_index}:D{header_row_index}",
            {"textFormat": {"bold": True}},
        )

        if data.head_rows:
            credit_debit_range = f"B{header_row_index + 1}:C{last_head_row_index}"
            worksheet.format(credit_debit_range, {"numberFormat": CURRENCY_NUMBER_FORMAT})

        totals_end_row_index = totals_start_row_index + 2  # Total Credits/Debits/Net Collection
        worksheet.format(
            f"A{totals_start_row_index}:A{totals_end_row_index}",
            {"textFormat": {"bold": True}},
        )
        worksheet.format(
            f"B{totals_start_row_index}:B{totals_end_row_index}",
            {"numberFormat": CURRENCY_NUMBER_FORMAT},
        )

        spreadsheet.batch_update({
            "requests": [{
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": len(TABLE_HEADER),
                    }
                }
            }]
        })

        log.info("Formatting applied to Final Report worksheet.")
    except Exception as exc:
        log.warning("Could not fully apply formatting to Final Report: %s", exc)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def generate_final_report(
    credentials_path: Path,
    sheet_id: str = MASTER_SHEET_ID,
    summary_worksheet_name: str = SUMMARY_WORKSHEET_NAME,
    final_report_worksheet_name: str = FINAL_REPORT_WORKSHEET_NAME,
    spreadsheet=None,
) -> SummaryData:
    """Generate the formatted Final Report from the Summary worksheet.

    Args:
        credentials_path: Path to the Google service-account credentials JSON.
        sheet_id: Spreadsheet ID containing both the Summary and Final Report
            worksheets.
        summary_worksheet_name: Worksheet/tab name to read the summary from.
        final_report_worksheet_name: Worksheet/tab name to write the report into.
        spreadsheet: Optional pre-opened gspread.Spreadsheet (skip re-auth).

    Returns:
        The parsed SummaryData used to build the report.

    Raises:
        FileNotFoundError: If the credentials file does not exist.
        RuntimeError: If the Summary worksheet does not exist yet — the
            Summary must be generated first via generate_summary.py.
    """
    if spreadsheet is None:
        client = get_gspread_client(credentials_path)
        spreadsheet = client.open_by_key(sheet_id)

    try:
        summary_worksheet = open_summary_worksheet(spreadsheet, summary_worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        log.error(
            "Summary worksheet %r not found. Run generate_summary.py first.",
            summary_worksheet_name,
        )
        raise RuntimeError(
            f"Summary worksheet {summary_worksheet_name!r} does not exist. "
            "Run generate_summary.py before generate_final_report.py."
        )

    summary_values = summary_worksheet.get_all_values()
    data = parse_summary_data(summary_values)

    final_report_worksheet = get_or_create_final_report_worksheet(
        spreadsheet, final_report_worksheet_name,
    )
    rows = build_report_rows(data)
    write_report(final_report_worksheet, rows)
    apply_formatting(spreadsheet, final_report_worksheet, data)

    return data


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a formatted Final Report from the Summary worksheet."
    )

    parser.add_argument("-c", "--credentials", type=Path, default=DEFAULT_CREDENTIALS)
    parser.add_argument("--sheet-id", default=MASTER_SHEET_ID, help="Override the spreadsheet ID.")
    parser.add_argument(
        "--summary-worksheet-name",
        default=SUMMARY_WORKSHEET_NAME,
        help="Override the Summary worksheet/tab name.",
    )
    parser.add_argument(
        "--final-report-worksheet-name",
        default=FINAL_REPORT_WORKSHEET_NAME,
        help="Override the Final Report worksheet/tab name.",
    )

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        data = generate_final_report(
            credentials_path=args.credentials,
            sheet_id=args.sheet_id,
            summary_worksheet_name=args.summary_worksheet_name,
            final_report_worksheet_name=args.final_report_worksheet_name,
        )
        log.info(
            "Final Report generation complete. Heads=%d TotalCredits=%.2f TotalDebits=%.2f NetCollection=%.2f",
            len(data.head_rows), data.total_credits, data.total_debits, data.net_collection,
        )
    except Exception as exc:
        log.error("Final Report generation failed: %s", exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
