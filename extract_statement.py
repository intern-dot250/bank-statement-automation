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
    "b/f",          # Brought Forward (opening balance marker)
    "c/f",          # Carried Forward (closing balance marker)
    "toll free",    # YES Bank footer contact block
]

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("extract_statement")


_DATE_PATTERNS = (
    re.compile(r"^\d{2}-[A-Za-z]{3}-\d{4}$"),  # 22-Jun-2026
    re.compile(r"^\d{4}-\d{2}-\d{2}$"),         # 2026-06-22 (ISO format)
)


def is_valid_date(text):
    if not text:
        return False
    stripped = str(text).strip()
    return any(pattern.match(stripped) for pattern in _DATE_PATTERNS)


# Word-position-based column detection. pdfplumber's grid-based table
# detection is unreliable on real bank PDFs: many statements omit
# horizontal ruling lines between transactions, so the grid detector
# silently merges or drops cells. Instead we:
#   1. find the header row by looking for the "Description" keyword,
#   2. map every recognised header token to a field (txn_date, description …),
#   3. assign each data word to a field by its x-position.
#
# Amount columns (credit/debit/balance) are right-aligned: the same
# number "0.00" printed under Deposits vs Withdrawals has different x0
# values depending on the number's width. Using the word's RIGHT edge (x1)
# for amount columns — and the header label's x1 as the column anchor —
# gives stable column membership regardless of number width.
_HEADER_TOKEN_FIELDS = {
    "transaction": "txn_date",
    "value":       "value_date",
    "description": "description",
    "narration":   "description",
    "particulars": "description",
    "reference":   "reference",
    "number":      "reference",
    "cheque":      "reference",
    "withdrawal":  "debit",
    "withdrawals": "debit",
    "debit":       "debit",
    "debits":      "debit",
    "deposit":     "credit",
    "deposits":    "credit",
    "credit":      "credit",
    "credits":     "credit",
    "running":     "balance",
    "balance":     "balance",
}

# Fields in this set use right-edge (x1) for column anchoring and bucketing.
_AMOUNT_FIELDS = frozenset({"credit", "debit", "balance"})

_REQUIRED_HEADER_FIELDS = ("txn_date", "description")

_FIELD_ORDER_FOR_ROW = [
    "txn_date", "value_date", "description", "reference",
    "credit", "debit", "balance",
]

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


def _detect_header_columns(
    words: list[dict],
) -> tuple[dict[str, tuple[float, float]], float] | tuple[None, None]:
    """Scan the page for the transaction-table header and return
    (column_ranges, header_bottom_y), or (None, None) if not found.

    column_ranges maps each field name to (start_x, end_x). For text
    fields (txn_date, value_date, description, reference) both the
    anchor and the per-word bucket position use the word's LEFT edge
    (x0). For amount fields (credit, debit, balance) they use the
    word's RIGHT edge (x1), which is stable for right-aligned numbers.

    We try every line that contains the "description" keyword and accept
    the first one for which the required fields (txn_date + description)
    can be resolved. This avoids a false positive when "description"
    appears in account-summary text (e.g. "Account Variant / description:
    Freedom Flexi 100") before the real transaction-table header.

    We also look at the line immediately preceding the anchor, which in
    some YES Bank DAILY formats contains "Transaction" and "Cheque" on
    one line while "Description" is on the line below.
    """
    lines = _cluster_lines(words)

    for anchor_index, anchor_line in enumerate(lines):
        if not any(_normalize_token(w["text"]) == "description" for w in anchor_line):
            continue

        anchor_top = anchor_line[0]["top"]
        header_words = list(anchor_line)

        # Preceding line: "Transaction"/"Cheque" may sit above "Description"
        # (YES Bank YPR-daily format).  Only look backward when the anchor
        # line itself already contains a recognised financial-column keyword
        # (deposits / withdrawals / running / balance) — that proves we're in
        # the real transaction-table header and not in account-summary prose
        # such as "A/C Opening Date" which would otherwise inject a spurious
        # Date → txn_date mapping.
        anchor_has_financial = any(
            _HEADER_TOKEN_FIELDS.get(_normalize_token(w["text"])) in _AMOUNT_FIELDS
            for w in anchor_line
        )
        if anchor_has_financial and anchor_index > 0:
            prev_line = lines[anchor_index - 1]
            if anchor_top - prev_line[0]["top"] <= _HEADER_BLOCK_TOLERANCE:
                if not any(is_valid_date(w["text"]) for w in prev_line):
                    header_words.extend(prev_line)

        # Following line: wrapped continuation, e.g. "Date" below "Transaction"
        if anchor_index + 1 < len(lines):
            next_line = lines[anchor_index + 1]
            if next_line[0]["top"] - anchor_top <= _HEADER_BLOCK_TOLERANCE:
                if not any(is_valid_date(w["text"]) for w in next_line):
                    header_words.extend(next_line)

        field_x0s: dict[str, list[float]] = {}
        field_x1s: dict[str, list[float]] = {}
        date_tokens: list[dict] = []

        for word in header_words:
            token = _normalize_token(word["text"])
            if token == "date":
                date_tokens.append(word)
                continue
            field = _HEADER_TOKEN_FIELDS.get(token)
            if field:
                field_x0s.setdefault(field, []).append(word["x0"])
                field_x1s.setdefault(field, []).append(word.get("x1", word["x0"]))

        for word in date_tokens:
            candidates = {
                f: min(xs)
                for f, xs in field_x0s.items()
                if f in ("txn_date", "value_date")
            }
            if not candidates:
                field_x0s.setdefault("txn_date", []).append(word["x0"])
                field_x1s.setdefault("txn_date", []).append(word.get("x1", word["x0"]))
                continue
            nearest = min(candidates, key=lambda f: abs(candidates[f] - word["x0"]))
            field_x0s.setdefault(nearest, []).append(word["x0"])
            field_x1s.setdefault(nearest, []).append(word.get("x1", word["x0"]))

        if any(f not in field_x0s for f in _REQUIRED_HEADER_FIELDS):
            continue  # try the next line with "description"

        # A real transaction table must have at least one debit/credit column.
        # Account-info text (e.g. "Available Balance") may contain "balance"
        # but never "deposits" / "withdrawals", so this extra guard rejects it.
        if not any(f in field_x0s for f in ("credit", "debit")):
            continue

        # Column anchor: left edge for text fields, right edge for amounts
        field_anchor: dict[str, float] = {}
        for field in field_x0s:
            if field in _AMOUNT_FIELDS:
                field_anchor[field] = max(field_x1s[field])
            else:
                field_anchor[field] = min(field_x0s[field])

        ordered = sorted(field_anchor, key=lambda f: field_anchor[f])

        ranges: dict[str, tuple[float, float]] = {}
        for i, field in enumerate(ordered):
            start = (
                float("-inf") if i == 0
                else (field_anchor[ordered[i - 1]] + field_anchor[field]) / 2
            )
            end = (
                float("inf") if i + 1 == len(ordered)
                else (field_anchor[field] + field_anchor[ordered[i + 1]]) / 2
            )
            ranges[field] = (start, end)

        header_bottom = max(w["bottom"] for w in header_words)
        return ranges, header_bottom

    return None, None


def _bucket_line(
    line: list[dict],
    column_ranges: dict[str, tuple[float, float]],
) -> dict[str, str]:
    """Assign each word to a field column.

    Uses x1 (right edge) for amount columns — right-aligned numbers
    have a stable right edge regardless of number width. Uses x0 for
    text columns.
    """
    buckets: dict[str, list[str]] = {field: [] for field in column_ranges}
    for word in line:
        for field, (start, end) in column_ranges.items():
            pos = word.get("x1", word["x0"]) if field in _AMOUNT_FIELDS else word["x0"]
            if start <= pos < end:
                buckets[field].append(word["text"])
                break
    return {field: " ".join(texts) for field, texts in buckets.items()}


def should_skip_row(row_text: str) -> bool:
    lowered = row_text.lower()
    return any(pattern in lowered for pattern in EXCLUDE_PATTERNS)


def extract_transactions_from_pdf(
    pdf_path: Path,
    password: str = "",
) -> list[list[str]]:
    """Reconstruct transaction rows from each page's word positions.

    Returns a list of 7-field rows in EXPECTED_COLUMNS order:
    Transaction Date, Value Date, Description, Reference, Credits,
    Debits, Balance.
    """
    log.info("Opening PDF: %s", pdf_path)
    transactions: list[list[str]] = []

    open_kwargs: dict = {"password": password} if password else {}
    with pdfplumber.open(str(pdf_path), **open_kwargs) as pdf:
        log.info("Processing %d pages...", len(pdf.pages))

        for page_num, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            if not words:
                continue

            column_ranges, header_bottom = _detect_header_columns(words)
            if column_ranges is None:
                log.debug("Page %d: no recognizable transaction header found.", page_num)
                continue

            log.debug("Page %d: columns=%s header_bottom=%.1f",
                      page_num, list(column_ranges), header_bottom)

            data_words = [w for w in words if w["top"] > header_bottom]
            lines = _cluster_lines(data_words)

            # Pre-compute field buckets for every line so look-ahead is cheap.
            fields_list = [_bucket_line(line, column_ranges) for line in lines]

            current: dict[str, str] | None = None
            pending_pre: list[str] = []  # description fragments before next date
            last_top: float = -1.0

            for idx, fields in enumerate(fields_list):
                current_top = lines[idx][0]["top"]
                # A gap > 50pt between adjacent lines signals a section
                # boundary (page footer, disclaimer block) — stop here so
                # that footer text doesn't get appended to the last transaction.
                if last_top >= 0 and current_top - last_top > 50.0:
                    break
                prev_top = last_top
                last_top = current_top

                line_text = " ".join(fields.get(f, "") for f in _FIELD_ORDER_FOR_ROW)

                if should_skip_row(line_text):
                    # A "Closing Balance" summary row marks the end of the
                    # transaction table — stop here so disclaimer text that
                    # immediately follows doesn't get appended to the last txn.
                    if "closing balance" in line_text.lower():
                        break
                    continue

                txn_date = fields.get("txn_date", "").strip()

                if is_valid_date(txn_date):
                    if current is not None:
                        transactions.append(
                            [current[f] for f in _FIELD_ORDER_FOR_ROW]
                        )
                    current = {f: fields.get(f, "").strip() for f in _FIELD_ORDER_FOR_ROW}
                    if pending_pre:
                        pre = " ".join(pending_pre)
                        desc = current.get("description", "")
                        current["description"] = (
                            (pre + " " + desc).strip() if desc else pre
                        )
                        pending_pre = []

                else:
                    desc_text = fields.get("description", "").strip()
                    has_amount = any(fields.get(f, "").strip() for f in _AMOUNT_FIELDS)
                    has_date_field = bool(
                        fields.get("txn_date", "").strip() or fields.get("value_date", "").strip()
                    )

                    # Find the next non-skipped line's y-position (whatever
                    # kind of line it is), to compare against this line's
                    # distance from the previous one.
                    next_top = None
                    for j in range(idx + 1, len(fields_list)):
                        nf = fields_list[j]
                        nt = " ".join(nf.get(f, "") for f in _FIELD_ORDER_FOR_ROW)
                        if should_skip_row(nt):
                            continue
                        next_top = lines[j][0]["top"]
                        break

                    gap_above = (current_top - prev_top) if prev_top >= 0 else None
                    gap_below = (next_top - current_top) if next_top is not None else None

                    # A continuation fragment belongs to whichever
                    # neighboring transaction block it sits visually
                    # closer to. Wrapped description text within one
                    # transaction is tightly, uniformly spaced; the gap to
                    # a genuinely different transaction's block is
                    # measurably larger. This single geometric comparison
                    # replaces a format-specific "is the next line a
                    # date?" guess — that guess breaks down whenever a
                    # description wraps across 2+ lines AFTER its own date
                    # row (YES Bank monthly/CR format): the *last* of
                    # those lines is followed by the next transaction's
                    # date row too, but it still belongs to the
                    # transaction above it, not the one below.
                    belongs_below = current is None or (
                        gap_below is not None and (gap_above is None or gap_below < gap_above)
                    )

                    if belongs_below and desc_text and not has_amount and not has_date_field:
                        # Pre-line for the next transaction
                        pending_pre.append(desc_text)
                    elif current is not None:
                        # Genuine continuation lines only carry description (and
                        # occasionally reference) text — amounts always appear on
                        # the date line that opened the transaction. Footer / banner
                        # text (phone numbers, legal notices, contact info) lands
                        # across ALL column areas. Skip any continuation line that
                        # has text in a date column or an amount column.
                        if has_date_field or has_amount:
                            continue
                        # Post-line: continuation of the current transaction
                        for f in _FIELD_ORDER_FOR_ROW:
                            extra = fields.get(f, "").strip()
                            if extra:
                                current[f] = (
                                    (current[f] + " " + extra).strip()
                                    if current[f] else extra
                                )

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


def extract_statement(input_path, output_path, password: str = ""):
    if not input_path.exists():
        raise FileNotFoundError(f"Input PDF not found: {input_path}")

    log.info("=" * 50)
    log.info("Starting extraction")
    log.info("=" * 50)

    transactions = extract_transactions_from_pdf(input_path, password=password)

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
    parser.add_argument("-i", "--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("-p", "--password", default="")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    try:
        extract_statement(args.input, args.output, password=args.password)
    except Exception as exc:
        log.exception("Error: %s", exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
