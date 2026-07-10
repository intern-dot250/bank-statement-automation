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


# pdfplumber's grid-based table detection (grouping rows by the PDF's
# drawn ruling lines) turned out to be unreliable on real statements: many
# bank statement PDFs don't draw a horizontal line under every single
# transaction row, only between visual sections. That caused the grid
# detector to silently merge a genuine transaction's date/amount cells
# into the row above (or drop them as None) — recovering zero amounts for
# 3 out of 4 real transactions on one confirmed page — with no error, so
# real money was quietly missing from the sheet.
#
# Row reconstruction below instead uses each word's own (x, y) position on
# the page: header words are matched by keyword to figure out which x
# range belongs to which field (Transaction Date / Description / etc,
# tolerant of a bank's own header wording and column order), then every
# word below the header is bucketed into a field by its x position and
# into a transaction by its y position (a new transaction starts at each
# line whose Transaction Date bucket holds a valid date; every line
# before the next one is that transaction's own wrapped continuation
# text). This depends only on text position, not on the presence of
# ruling lines, so it isn't fooled by a statement that skips them.
_HEADER_TOKEN_FIELDS = {
    "transaction": "txn_date",
    "value": "value_date",
    "description": "description",
    "reference": "reference",
    "number": "reference",
    "cheque": "reference",
    "withdrawal": "debit",
    "withdrawals": "debit",
    "debit": "debit",
    "debits": "debit",
    "deposit": "credit",
    "deposits": "credit",
    "credit": "credit",
    "credits": "credit",
    "running": "balance",
    "balance": "balance",
}

# These two fields must be present for a page's header to be usable at
# all — a header with no discernible date or description column can't be
# safely mapped, so the page is skipped entirely rather than guessed at.
_REQUIRED_HEADER_FIELDS = ("txn_date", "description")

_FIELD_ORDER_FOR_ROW = ["txn_date", "value_date", "description", "reference", "credit", "debit", "balance"]

# Words on the same visual line rarely differ in `top` by more than this
# (font size variance / sub-pixel rendering) — used to cluster words into
# lines and to decide whether a header continuation line (e.g. wrapped
# "Date" below "Transaction") still belongs to the header block.
_LINE_TOLERANCE = 3.0
_HEADER_BLOCK_TOLERANCE = 25.0


def _normalize_token(text: str) -> str:
    return re.sub(r"[^a-z]", "", text.lower())


def _cluster_lines(words: list[dict]) -> list[list[dict]]:
    """Group words into visual lines by their `top` (y) position."""
    lines: list[list[dict]] = []
    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(word["top"] - lines[-1][0]["top"]) <= _LINE_TOLERANCE:
            lines[-1].append(word)
        else:
            lines.append([word])
    for line in lines:
        line.sort(key=lambda w: w["x0"])
    return lines


def _detect_header_columns(words: list[dict]) -> dict[str, tuple[float, float]] | None:
    """Find the header line(s) on this page and return each recognized
    field's x-position range: {field: (start_x, end_x)}, sorted left to
    right, with the last field's range extending to infinity. Returns
    None if no usable header (with at least Transaction Date and
    Description) is found on this page."""
    lines = _cluster_lines(words)

    anchor_index = None
    for i, line in enumerate(lines):
        if any(_normalize_token(w["text"]) == "description" for w in line):
            anchor_index = i
            break

    if anchor_index is None:
        return None

    anchor_top = lines[anchor_index][0]["top"]
    header_words = list(lines[anchor_index])

    # Pull in a wrapped continuation line (e.g. "Date" wrapping below
    # "Transaction") if it immediately follows and isn't itself a data row.
    if anchor_index + 1 < len(lines):
        next_line = lines[anchor_index + 1]
        if next_line[0]["top"] - anchor_top <= _HEADER_BLOCK_TOLERANCE:
            if not any(is_valid_date(w["text"]) for w in next_line):
                header_words.extend(next_line)

    # Classify every header word by keyword, tracking each field's x0s so
    # an ambiguous wrapped "Date" token can be assigned to whichever of
    # Transaction/Value it's positioned closer to.
    field_x0s: dict[str, list[float]] = {}
    date_tokens: list[dict] = []
    for word in header_words:
        token = _normalize_token(word["text"])
        if token == "date":
            date_tokens.append(word)
            continue
        field = _HEADER_TOKEN_FIELDS.get(token)
        if field:
            field_x0s.setdefault(field, []).append(word["x0"])

    for word in date_tokens:
        candidates = {
            field: min(xs) for field, xs in field_x0s.items() if field in ("txn_date", "value_date")
        }
        if not candidates:
            field_x0s.setdefault("txn_date", []).append(word["x0"])
            continue
        nearest_field = min(candidates, key=lambda f: abs(candidates[f] - word["x0"]))
        field_x0s.setdefault(nearest_field, []).append(word["x0"])

    if any(field not in field_x0s for field in _REQUIRED_HEADER_FIELDS):
        return None

    field_start = {field: min(xs) for field, xs in field_x0s.items()}
    ordered_fields = sorted(field_start, key=lambda f: field_start[f])

    # Boundaries sit at the MIDPOINT between two adjacent fields' header
    # start positions, not at each field's own start — a wrapped
    # continuation word's left edge can land a couple points to either
    # side of the header label it belongs under (e.g. a wrapped
    # reference-number fragment starting slightly left of where
    # "Reference" itself started), so a boundary drawn exactly at the
    # header position misclassifies it into the previous column.
    ranges: dict[str, tuple[float, float]] = {}
    for i, field in enumerate(ordered_fields):
        start = (
            float("-inf") if i == 0
            else (field_start[ordered_fields[i - 1]] + field_start[field]) / 2
        )
        end = (
            float("inf") if i + 1 == len(ordered_fields)
            else (field_start[field] + field_start[ordered_fields[i + 1]]) / 2
        )
        ranges[field] = (start, end)

    return ranges


def _bucket_line(line: list[dict], column_ranges: dict[str, tuple[float, float]]) -> dict[str, str]:
    buckets: dict[str, list[str]] = {field: [] for field in column_ranges}
    for word in line:
        for field, (start, end) in column_ranges.items():
            if start <= word["x0"] < end:
                buckets[field].append(word["text"])
                break
    return {field: " ".join(texts) for field, texts in buckets.items()}


def should_skip_row(row_text: str) -> bool:
    lowered = row_text.lower()
    return any(pattern in lowered for pattern in EXCLUDE_PATTERNS)


def extract_transactions_from_pdf(pdf_path: Path) -> list[list[str]]:
    """Reconstruct transaction rows from each page's word positions.

    Returns a list of 7-field rows already in EXPECTED_COLUMNS order
    (Transaction Date, Value Date, Description, Reference, Credits,
    Debits, Balance).
    """
    log.info("Opening PDF: %s", pdf_path)

    transactions: list[list[str]] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        log.info("Processing %d pages...", len(pdf.pages))

        for page_num, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            if not words:
                continue

            column_ranges = _detect_header_columns(words)
            if column_ranges is None:
                log.debug("Page %d: no recognizable transaction header found.", page_num)
                continue

            header_bottom = max(
                w["bottom"] for w in words
                if any(start <= w["x0"] < end for start, end in column_ranges.values())
                and _normalize_token(w["text"]) in _HEADER_TOKEN_FIELDS
            )

            data_words = [w for w in words if w["top"] > header_bottom]
            lines = _cluster_lines(data_words)

            current: dict[str, str] | None = None
            for line in lines:
                fields = _bucket_line(line, column_ranges)
                line_text = " ".join(fields.get(f, "") for f in _FIELD_ORDER_FOR_ROW)

                if should_skip_row(line_text):
                    continue

                txn_date = fields.get("txn_date", "").strip()

                if is_valid_date(txn_date):
                    if current is not None:
                        transactions.append([current[f] for f in _FIELD_ORDER_FOR_ROW])
                    current = {f: fields.get(f, "").strip() for f in _FIELD_ORDER_FOR_ROW}
                    current["txn_date"] = txn_date
                elif current is not None:
                    # Continuation line — append any wrapped text (mainly
                    # Description, occasionally Reference) to the
                    # transaction already in progress.
                    for f in _FIELD_ORDER_FOR_ROW:
                        extra = fields.get(f, "").strip()
                        if extra:
                            current[f] = (current[f] + " " + extra).strip() if current[f] else extra

            if current is not None:
                transactions.append([current[f] for f in _FIELD_ORDER_FOR_ROW])

    return transactions


def build_dataframe(transactions: list[list[str]]) -> pd.DataFrame:
    log.info("Valid transactions found: %d", len(transactions))
    print(f"Retained {len(transactions)} data row(s) after filtering.")

    return pd.DataFrame(transactions, columns=EXPECTED_COLUMNS)


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

    transactions = extract_transactions_from_pdf(input_path)

    if not transactions:
        raise ValueError("No rows found in PDF")

    df = build_dataframe(transactions)

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