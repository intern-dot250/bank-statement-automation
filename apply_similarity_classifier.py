"""Similarity-based AI classifier using existing classified transactions as reference.

How it works:
  1. Reads all classified rows from Google Sheet (250+ transactions)
  2. Builds a TF-IDF index on their DESCRIPTION text
  3. For each '?' row, finds the top 5 most similar past transactions
  4. If 3 or more of the top 5 agree on the same HEAD -> auto-classify
  5. Fills BU / Type / TCP using the same account-stage logic as the rules
  6. Marks AI-classified rows in YELLOW text (distinct from red=rules, orange=Claude)

Run after resolve_unknowns.py:
  py -3 apply_similarity_classifier.py

No API key or internet AI service required.
"""

from __future__ import annotations

import re
from collections import Counter
from typing import Any, Optional

import gspread
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

from classify_transactions import (
    BUSINESS_UNIT_COLUMN,
    HEAD_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
    NARRATION_COLUMN,
    UNKNOWN_MAPPING_VALUE,
    _get_accounts_by_number,
    _ACCOUNT_BU_OVERRIDES,
    _ACCOUNT_STAGE_OVERRIDES,
    _HO_ADMIN_DEFAULTS,
    STAGE_VENDOR_DEFAULTS,
    _get_cell,
    _is_row_empty,
    _to_float,
    _parse_amount,
)
from narration import generate_narration
from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_gspread_client,
    get_account_worksheets,
)

# Yellow text = AI-classified via similarity (needs team verification)
SIMILARITY_TEXT_COLOR = {"red": 0.8, "green": 0.7, "blue": 0.0}

# Heads excluded from training data — handled by specific rules, not similarity
SKIP_HEADS = {"Internal", "Collection", "?", ""}

# Similarity thresholds
TOP_N = 5          # how many similar past rows to look at
MIN_VOTES = 3      # how many of the top N must agree on the same HEAD
MIN_SIMILARITY = 0.25  # minimum cosine similarity to count a vote at all

# Regex to strip transaction reference codes and IFSC patterns
# (unique per transaction, add noise to similarity)
_REF_RE = re.compile(r'\b(YESME|YESB|HDFC|SBIN|CNRB|UBIN|ICIC|KKBK|UTIB|BARB|PUNB|IDIB|FDRL|IDFC|MAHB)[A-Z0-9]{6,}\b', re.IGNORECASE)
_DIGITS_RE = re.compile(r'\b\d{4,}\b')  # long digit sequences (account/ref numbers)


def _preprocess(description: str) -> str:
    """Clean a description for TF-IDF: remove ref codes, long numbers,
    normalize separators to spaces so word tokens are clean."""
    text = description.upper()
    text = _REF_RE.sub(' ', text)
    text = _DIGITS_RE.sub(' ', text)
    text = re.sub(r'[-/]', ' ', text)   # split on hyphens and slashes
    text = re.sub(r'\s+', ' ', text).strip()
    return text


class SimilarityClassifier:
    """TF-IDF similarity index over existing classified transactions."""

    def __init__(self) -> None:
        self.descriptions: list[str] = []   # raw (for display)
        self.heads: list[str] = []
        self.accounts: list[str] = []
        self.vectorizer = TfidfVectorizer(
            analyzer='word',
            ngram_range=(1, 2),   # unigrams + bigrams catch "SALARY SITE", "CIVIL WORKS"
            min_df=1,
            sublinear_tf=True,    # log(1+tf) — reduces weight of very common terms
        )
        self._matrix = None

    def fit(self, spreadsheet: gspread.Spreadsheet) -> int:
        """Load all classified rows and build the TF-IDF index.
        Returns the number of training rows loaded."""
        for ws in get_account_worksheets(spreadsheet):
            rows = ws.get_all_values()
            if not rows or 'HEAD' not in rows[0]:
                continue
            hdr = rows[0]
            for row in rows[1:]:
                if _is_row_empty(row):
                    continue
                head = _get_cell(row, hdr, 'HEAD')
                desc = _get_cell(row, hdr, 'DESCRIPTION')
                acc = _get_cell(row, hdr, 'Account Number')
                if not desc or head in SKIP_HEADS:
                    continue
                self.descriptions.append(desc)
                self.heads.append(head)
                self.accounts.append(acc)

        if not self.descriptions:
            return 0

        processed = [_preprocess(d) for d in self.descriptions]
        self._matrix = self.vectorizer.fit_transform(processed)
        return len(self.descriptions)

    def predict(self, description: str) -> Optional[dict[str, Any]]:
        """Return prediction dict {head, top_matches, confidence} or None
        if no confident match found."""
        if self._matrix is None or not self.descriptions:
            return None

        query_vec = self.vectorizer.transform([_preprocess(description)])
        sims = cosine_similarity(query_vec, self._matrix)[0]

        # Get indices of top N results above minimum similarity
        top_indices = np.argsort(sims)[::-1][:TOP_N]
        top_matches = [
            {
                "description": self.descriptions[i],
                "head": self.heads[i],
                "similarity": float(sims[i]),
            }
            for i in top_indices
            if sims[i] >= MIN_SIMILARITY
        ]

        if not top_matches:
            return None

        # Vote: count how many qualifying matches agree on the same HEAD
        vote_counter: Counter = Counter()
        for m in top_matches:
            vote_counter[m["head"]] += 1

        best_head, best_votes = vote_counter.most_common(1)[0]

        if best_votes < MIN_VOTES:
            return None   # not enough agreement

        confidence = "high" if best_votes >= 4 else "medium"
        return {
            "head": best_head,
            "votes": best_votes,
            "confidence": confidence,
            "top_matches": top_matches[:3],   # show top 3 for logging
        }


def _resolve_fields_for_head(
    head: str,
    own_bu: str,
    own_stage: Optional[str],
) -> dict[str, str]:
    """Determine BU / Type / TCP for a similarity-predicted head."""
    if head in ("Salary HO", "Professional", "Statutory Dues"):
        return {
            "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
            "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
            "tcp_head": _HO_ADMIN_DEFAULTS["tcp_head"],
        }
    if head == "HO - Advert/Mkt":
        return {
            "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
            "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
            "tcp_head": "Other-Selling Expenses",
        }
    if head == "Bank Charges":
        return {
            "business_unit": own_bu,
            "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
            "tcp_head": "Other- Others",
        }
    if head == "Collection":
        return {
            "business_unit": own_bu,
            "type_rera_idw": "Customer Collection",
            "tcp_head": "Credit- no effect",
        }
    if head == "Cancellation":
        return {
            "business_unit": own_bu,
            "type_rera_idw": "Cust Cancellation",
            "tcp_head": UNKNOWN_MAPPING_VALUE,
        }
    if head == "Internal":
        return {
            "business_unit": own_bu,
            "type_rera_idw": "Internal",
            "tcp_head": "Internal transfer",
        }
    # Vendor / Contractor / Imprest / Salary Site
    if own_stage == "Free":
        return {
            "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
            "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
            "tcp_head": _HO_ADMIN_DEFAULTS["tcp_head"],
        }
    defaults = STAGE_VENDOR_DEFAULTS.get(own_stage or "", {})
    return {
        "business_unit": own_bu,
        "type_rera_idw": defaults.get("type_rera_idw", UNKNOWN_MAPPING_VALUE),
        "tcp_head": defaults.get("tcp_head", UNKNOWN_MAPPING_VALUE),
    }


def _mark_similarity_rows(
    worksheet: gspread.Worksheet,
    row_numbers: list[int],
    col_indices: dict[str, int],
) -> None:
    """Color similarity-classified cells yellow."""
    if not row_numbers:
        return
    target_cols = [
        BUSINESS_UNIT_COLUMN, HEAD_COLUMN,
        TYPE_RERA_IDW_COLUMN, TCP_HEAD_COLUMN, NARRATION_COLUMN,
    ]
    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": row - 1,
                    "endRowIndex": row,
                    "startColumnIndex": col_indices[col] - 1,
                    "endColumnIndex": col_indices[col],
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"foregroundColor": SIMILARITY_TEXT_COLOR}
                    }
                },
                "fields": "userEnteredFormat.textFormat.foregroundColor",
            }
        }
        for row in row_numbers
        for col in target_cols
        if col in col_indices
    ]
    try:
        worksheet.spreadsheet.batch_update({"requests": requests})
    except Exception as exc:
        print(f"  [WARN] Could not apply yellow color: {exc}")


def process_worksheet(
    worksheet: gspread.Worksheet,
    classifier: SimilarityClassifier,
) -> tuple[int, int]:
    """Classify all ? rows in one worksheet. Returns (resolved, still_unknown)."""
    all_values = worksheet.get_all_values()
    if not all_values:
        return 0, 0

    hdr = all_values[0]
    required = [HEAD_COLUMN, BUSINESS_UNIT_COLUMN, TYPE_RERA_IDW_COLUMN,
                TCP_HEAD_COLUMN, NARRATION_COLUMN]
    if any(c not in hdr for c in required):
        return 0, 0

    col_idx = {c: hdr.index(c) + 1 for c in required}
    updates: list[gspread.cell.Cell] = []
    sim_rows: list[int] = []
    resolved = 0
    still_unknown = 0

    for offset, row in enumerate(all_values[1:]):
        sheet_row = offset + 2

        if _is_row_empty(row):
            continue

        if _get_cell(row, hdr, HEAD_COLUMN) != UNKNOWN_MAPPING_VALUE:
            continue

        description = _get_cell(row, hdr, 'DESCRIPTION')
        if not description:
            continue

        account_number = _get_cell(row, hdr, 'Account Number')
        deposits  = _to_float(_get_cell(row, hdr, 'CREDITS'))
        withdrawals = _to_float(_get_cell(row, hdr, 'DEBITS'))
        amount = _parse_amount(
            _get_cell(row, hdr, 'CREDITS'),
            _get_cell(row, hdr, 'DEBITS'),
        )

        own_account = _get_accounts_by_number().get(account_number, {})
        own_bu = own_account.get('business_unit') or next(
            (bu for sfx, bu in _ACCOUNT_BU_OVERRIDES.items()
             if account_number.endswith(sfx)),
            UNKNOWN_MAPPING_VALUE,
        )
        own_stage = own_account.get('account_stage') or next(
            (s for sfx, s in _ACCOUNT_STAGE_OVERRIDES.items()
             if account_number.endswith(sfx)),
            None,
        )

        prediction = classifier.predict(description)

        if prediction is None:
            print(f"  [{worksheet.title}] row {sheet_row}: no confident match")
            print(f"    DESC: {description[:80]}")
            still_unknown += 1
            continue

        ai_head = prediction['head']
        votes    = prediction['votes']
        confidence = prediction['confidence']

        print(f"  [{worksheet.title}] row {sheet_row}: {description[:70]}")
        print(f"    -> {ai_head} ({votes}/{TOP_N} votes, {confidence} confidence)")
        for m in prediction['top_matches']:
            print(f"       sim={m['similarity']:.2f} | {m['head']} | {m['description'][:60]}")

        fields = _resolve_fields_for_head(ai_head, own_bu, own_stage)

        new_narration = generate_narration(
            description, ai_head, amount,
            business_unit=fields['business_unit'],
            type_rera_idw=fields['type_rera_idw'],
            deposits=deposits,
            withdrawals=withdrawals,
            own_account_number=account_number,
        )

        for col_name, value in [
            (BUSINESS_UNIT_COLUMN, fields['business_unit']),
            (HEAD_COLUMN,          ai_head),
            (TYPE_RERA_IDW_COLUMN, fields['type_rera_idw']),
            (TCP_HEAD_COLUMN,      fields['tcp_head']),
            (NARRATION_COLUMN,     new_narration),
        ]:
            updates.append(
                gspread.cell.Cell(row=sheet_row, col=col_idx[col_name], value=value)
            )
        sim_rows.append(sheet_row)
        resolved += 1

    if updates:
        worksheet.update_cells(updates, value_input_option='RAW')
        _mark_similarity_rows(worksheet, sim_rows, col_idx)

    return resolved, still_unknown


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    print("Building similarity index from existing classified transactions...")
    classifier = SimilarityClassifier()
    n = classifier.fit(spreadsheet)
    print(f"Index built: {n} training rows loaded.\n")

    if n == 0:
        print("No training data found. Run classify_transactions.py first.")
        return

    total_resolved = 0
    total_unknown  = 0

    for worksheet in get_account_worksheets(spreadsheet):
        resolved, unknown = process_worksheet(worksheet, classifier)
        total_resolved += resolved
        total_unknown  += unknown

    print()
    print(f"Total AI-resolved (similarity): {total_resolved}")
    print(f"Total still unknown (needs manual): {total_unknown}")
    if total_resolved > 0:
        print("Yellow text = similarity-classified. Team should verify these rows.")


if __name__ == "__main__":
    main()
