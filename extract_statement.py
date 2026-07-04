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
    "Deposits",
    "Withdrawals",
    "Running Balance",
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


def is_valid_date(text):
    if not text:
        return False
    return bool(re.match(r"\d{2}-[A-Za-z]{3}-\d{4}", str(text).strip()))


def extract_tables_from_pdf(pdf_path: Path):
    log.info("Opening PDF: %s", pdf_path)

    all_rows = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        log.info("Processing %d pages...", len(pdf.pages))

        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()

            if not tables:
                continue

            for table in tables:
                for row in table:
                    if row:
                        all_rows.append(row)

    return all_rows


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

    if "transaction date" in row_text:
        return True

    for pattern in EXCLUDE_PATTERNS:
        if pattern in row_text:
            return True

    return False


def build_dataframe(rows):
    transactions = []

    for row in rows:
        row = clean_row(row)

        if len(row) < 7:
            continue

        if should_skip_row(row):
            continue

        # Only accept rows that start with valid date
        if not is_valid_date(row[0]):
            continue

        transactions.append(row[:7])

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

    rows = extract_tables_from_pdf(input_path)

    if not rows:
        raise ValueError("No rows found in PDF")

    df = build_dataframe(rows)

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