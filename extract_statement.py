from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber

DEFAULT_INPUT = Path("output/unlocked_statement.pdf")
DEFAULT_OUTPUT = Path("output/bank_statement.xlsx")

LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

EXPECTED_COLUMNS = [
    "Transaction Date",
    "Value Date",
    "Description",
    "Cheque No/Reference No",
    "Credits",
    "Debits",
    "Balance",
]

EXCLUDE_PATTERNS = [
    "opening balance",
    "closing balance",
    "total deposits",
    "total withdrawals",
    "summary",
    "page",
]

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("extract_statement")


_DATE_PATTERNS = (
    re.compile(r"^\d{2}-[A-Za-z]{3}-\d{4}$"),  # 22-Jun-2026
    re.compile(r"^\d{4}-\d{2}-\d{2}$"),         # 2026-06-22 (ISO format, also seen in live statements)
)


def is_valid_date(text):
    if not text:
        return False
    stripped = str(text).strip()
    return any(pattern.match(stripped) for pattern in _DATE_PATTERNS)


def extract_tables_from_pdf(pdf_path: Path):
    """Returns a list of tables (each table itself a list of raw rows),
    preserving per-table structure — needed so each table's own header row
    can be used to figure out its column layout (see build_dataframe)."""
    log.info("Opening PDF: %s", pdf_path)

    all_tables = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        log.info("Processing %d pages...", len(pdf.pages))

        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()

            if not tables:
                continue

            for table in tables:
                cleaned_table = [row for row in table if row]
                if cleaned_table:
                    all_tables.append(cleaned_table)

    return all_tables


def clean_row(row):
    cleaned = []

    for cell in row:
        if cell is None:
            cleaned.append("")
        else:
            cleaned.append(" ".join(str(cell).split()))

    return cleaned


def should_skip_row(row):
    row_text = " ".join(row).lower()

    for pattern in EXCLUDE_PATTERNS:
        if pattern in row_text:
            return True

    return False


# Different banks label the same columns differently (e.g. "Withdrawals"/
# "Deposits" vs "Debits"/"Credits"), and — critically — don't always put
# them in the same left-to-right order. Detecting each table's own header
# row and matching columns by keyword (rather than assuming a fixed
# position) is what keeps a debit from ever being silently swapped with a
# credit just because one bank's layout lists them in the opposite order.
_HEADER_FIELD_KEYWORDS = {
    "txn_date": ["transaction date"],
    "value_date": ["value date"],
    "description": ["description"],
    "reference": ["reference number", "cheque no", "reference"],
    "debit": ["withdrawal", "debit"],
    "credit": ["deposit", "credit"],
    "balance": ["balance"],
}

# These two fields must be present for a table to be usable at all — a
# table with no discernible date or description column is not a
# transaction table we can safely map, so it's skipped entirely rather
# than guessed at.
_REQUIRED_HEADER_FIELDS = ("txn_date", "description")


def _is_header_row(cleaned_row):
    row_text = " ".join(cleaned_row).lower()
    return "transaction date" in row_text and "description" in row_text


def _detect_header_map(cleaned_header_row):
    """Map each EXPECTED_COLUMNS field to the column index that matches
    its keyword(s) in this specific table's header row. A field whose
    keyword isn't found gets index None (rendered as an empty cell for
    every row in that table, never guessed)."""
    lowered = [cell.lower() for cell in cleaned_header_row]
    field_index = {}
    for field, keywords in _HEADER_FIELD_KEYWORDS.items():
        found_index = None
        for i, cell in enumerate(lowered):
            if any(keyword in cell for keyword in keywords):
                found_index = i
                break
        field_index[field] = found_index
    return field_index


def _cell_at(row, index):
    if index is None or index >= len(row):
        return ""
    return row[index]


def build_dataframe(tables):
    transactions = []

    for table in tables:
        header_map = None

        for raw_row in table:
            row = clean_row(raw_row)

            if _is_header_row(row):
                header_map = _detect_header_map(row)
                continue

            if header_map is None:
                # No header seen yet in this table — nothing to map
                # against, so this row can't be safely interpreted.
                continue

            if any(header_map[field] is None for field in _REQUIRED_HEADER_FIELDS):
                continue

            if should_skip_row(row):
                continue

            txn_date = _cell_at(row, header_map["txn_date"])

            # Only accept rows that start with a valid date — continuation
            # rows (wrapped description text with no date/amounts) are
            # intentionally skipped, not treated as separate transactions.
            if not is_valid_date(txn_date):
                continue

            transactions.append([
                txn_date,
                _cell_at(row, header_map["value_date"]),
                _cell_at(row, header_map["description"]),
                _cell_at(row, header_map["reference"]),
                _cell_at(row, header_map["credit"]),
                _cell_at(row, header_map["debit"]),
                _cell_at(row, header_map["balance"]),
            ])

    log.info("Valid transactions found: %d", len(transactions))
    print(f"Retained {len(transactions)} data row(s) after filtering.")

    df = pd.DataFrame(transactions, columns=EXPECTED_COLUMNS)

    return df


def save_to_excel(df, output_path):
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.exists():
        output_path.unlink()

    df.to_excel(output_path, index=False)

    log.info("Saved Excel: %s", output_path)


def extract_statement(input_path, output_path):
    if not input_path.exists():
        raise FileNotFoundError(f"Input PDF not found: {input_path}")

    log.info("=" * 50)
    log.info("Starting extraction")
    log.info("=" * 50)

    tables = extract_tables_from_pdf(input_path)

    if not tables:
        raise ValueError("No rows found in PDF")

    df = build_dataframe(tables)

    if df.empty:
        raise ValueError("No valid transaction rows extracted")

    save_to_excel(df, output_path)

    log.info("Extracted rows count: %d", len(df))
    log.info("=" * 50)
    log.info("Extraction complete")
    log.info("=" * 50)


def parse_args(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        default=DEFAULT_INPUT
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    try:
        extract_statement(args.input, args.output)
    except Exception as exc:
        log.exception("Error: %s", exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())