"""Priority 3: RAG-based AI classifier using Groq (free tier).

Architecture
------------
1. Retrieval  — TF-IDF finds the 5 most similar already-classified
                transactions from the sheet.
2. Augment    — Those 5 examples are injected into the LLM prompt so
                the model reasons with real DPL data, not just training
                knowledge.
3. Generate   — Groq (GPT-OSS-120B, free tier) returns the head and
                optionally suggests adding the beneficiary to the master.

Only fires for rows still showing HEAD = '?' after all rule-based and
keyword classification has run.  For typical monthly volumes this will
be 5-15 rows — well inside Groq's 14,400-request/day free limit.

Color coding after this runs
  Red    = auto-classified by rules
  Yellow = similarity (TF-IDF votes, no LLM)
  Orange = RAG AI (Groq)  ← this script
  Black  = verified by accounts team

Usage
-----
  py -3 rag_classifier.py

Environment variable required
  GROQ_API_KEY=gsk_...   (free at console.groq.com)
"""

from __future__ import annotations

import json
import os
import re
from collections import Counter
from pathlib import Path
from typing import Any, Optional

import gspread
import numpy as np
from groq import Groq
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

from classify_transactions import (
    BUSINESS_UNIT_COLUMN,
    HEAD_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
    NARRATION_COLUMN,
    CONFIDENCE_COLUMN,
    REASON_COLUMN,
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

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

GROQ_MODEL = "openai/gpt-oss-120b"

# Retrieval settings
TOP_N = 5          # how many similar past rows to retrieve
MIN_SIMILARITY = 0.15  # lower than similarity-only classifier — LLM does reasoning

# Orange text = RAG AI classified
AI_TEXT_COLOR = {"red": 1.0, "green": 0.5, "blue": 0.0}

# Heads excluded from training data
SKIP_HEADS = {"Internal", "Collection", "?", ""}

VALID_HEADS = [
    "Vendor", "Contractor", "Professional", "Imprest",
    "Salary HO", "Salary Site", "Internal", "Collection",
    "Cancellation", "Statutory Dues", "HO - Advert/Mkt", "Bank Charges",
]

# Heads that were classified by a keyword the user typed — these can be wrong
# if the user typed the wrong remark. RAG will re-verify these rows for named
# payees not yet confirmed in the Beneficiary Master.
KEYWORD_CLASSIFIED_HEADS = {"Vendor", "Contractor", "Professional"}

MASTER_TAB_NAME = "Beneficiary Master"
MASTER_STATUS_CONFIRMED = "Confirmed"
MASTER_STATUS_AI = "AI Suggested"

# ---------------------------------------------------------------------------
# Text preprocessing (reused from apply_similarity_classifier.py)
# ---------------------------------------------------------------------------

_REF_RE = re.compile(
    r'\b(YESME|YESB|HDFC|SBIN|CNRB|UBIN|ICIC|KKBK|UTIB|BARB|PUNB|IDIB|FDRL|IDFC|MAHB)'
    r'[A-Z0-9]{6,}\b',
    re.IGNORECASE,
)
_DIGITS_RE = re.compile(r'\b\d{4,}\b')
_ROLE_SUFFIX = re.compile(
    r'\s+(IMPREST|SALARY|CONTRACTOR|PROFESSIONAL|VENDOR|ADVANCE|REFUND)$',
    re.IGNORECASE,
)
_IFSC_LIKE = re.compile(r'^[A-Z]{4}0[A-Z0-9]{6}$')


def _preprocess(description: str) -> str:
    text = description.upper()
    text = _REF_RE.sub(' ', text)
    text = _DIGITS_RE.sub(' ', text)
    text = re.sub(r'[-/]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def _extract_name_from_desc(desc: str) -> Optional[str]:
    """Best-effort beneficiary name extraction for master list suggestion."""
    upper = desc.upper()
    if upper.startswith("YIB-NEFT") or upper.startswith("YIB-TPT"):
        parts = desc.split("-")
        if len(parts) >= 4:
            name = parts[3].strip()
            if name and not _IFSC_LIKE.match(name) and not name.isdigit():
                return _ROLE_SUFFIX.sub("", name).strip().upper() or None
    if upper.startswith("IMPS/"):
        parts = desc.split("/")
        if len(parts) >= 3:
            name = parts[-2].strip()
            if name and not name.upper().startswith("RRN") and not name.isdigit():
                return _ROLE_SUFFIX.sub("", name).strip().upper() or None
    return None


# ---------------------------------------------------------------------------
# Training data loader
# ---------------------------------------------------------------------------

def load_training_data(
    spreadsheet: gspread.Spreadsheet,
) -> tuple[list[str], list[str]]:
    """Load all classified rows → (descriptions, heads) for TF-IDF index."""
    descriptions: list[str] = []
    heads: list[str] = []

    for ws in get_account_worksheets(spreadsheet):
        rows = ws.get_all_values()
        if not rows or "HEAD" not in rows[0]:
            continue
        hdr = rows[0]
        for row in rows[1:]:
            if _is_row_empty(row):
                continue
            head = _get_cell(row, hdr, "HEAD")
            desc = _get_cell(row, hdr, "DESCRIPTION")
            if not desc or head in SKIP_HEADS:
                continue
            descriptions.append(desc)
            heads.append(head)

    return descriptions, heads


# ---------------------------------------------------------------------------
# Retrieval layer
# ---------------------------------------------------------------------------

class Retriever:
    def __init__(self, descriptions: list[str], heads: list[str]) -> None:
        self.descriptions = descriptions
        self.heads = heads
        self._vectorizer = TfidfVectorizer(
            analyzer="word",
            ngram_range=(1, 2),
            min_df=1,
            sublinear_tf=True,
        )
        processed = [_preprocess(d) for d in descriptions]
        self._matrix = self._vectorizer.fit_transform(processed)

    def retrieve(self, query: str) -> list[dict[str, Any]]:
        """Return up to TOP_N most similar past rows above MIN_SIMILARITY."""
        if not self.descriptions:
            return []
        query_vec = self._vectorizer.transform([_preprocess(query)])
        sims = cosine_similarity(query_vec, self._matrix)[0]
        top_idx = np.argsort(sims)[::-1][:TOP_N]
        return [
            {
                "description": self.descriptions[i],
                "head": self.heads[i],
                "similarity": float(sims[i]),
            }
            for i in top_idx
            if sims[i] >= MIN_SIMILARITY
        ]


# ---------------------------------------------------------------------------
# LLM layer
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = """\
You are a financial transaction classifier for DPL (Dwarkadhis Projects Limited), \
an Indian real estate company under CIRP (insolvency proceedings). \
Classify bank transactions into the correct Head category.

Valid Head values:
- Vendor: material/hardware/consumables suppliers
- Contractor: construction contractors, sub-contractors, daily-wage labourers
- Professional: CA firms, architects, legal consultants
- Imprest: petty cash advance given to an employee (transaction type, not identity)
- Salary HO: head office staff monthly salary
- Salary Site: site/construction staff monthly salary
- Internal: transfer between DPL's own bank accounts
- Collection: incoming payment from apartment buyers
- Cancellation: refund to a customer who cancelled their booking
- Statutory Dues: PF, ESI, TDS, Professional Tax, PTAX
- HO - Advert/Mkt: marketing, advertising, hoarding, branding
- Bank Charges: bank fees, locker charges, service charges

Rules:
1. Use the similar past transactions as your primary guide — they are real DPL data.
2. If 3 or more similar transactions agree on a head, follow them unless the new \
description clearly contradicts.
3. IMPREST in the description always means Imprest head (cash advance), \
even if the person's name appears in other heads historically.
4. If the direction is Credit (incoming), prefer Collection unless it is clearly internal.
5. If genuinely unclear, return head = "?" — do not guess randomly.
6. suggest_master = true only if the payee is a named individual or company \
(not a bank charge or statutory payment) and you are confident about the head.

Respond ONLY with valid JSON, no markdown, no explanation:
{
  "head": "<exact head from list or ?>",
  "reason": "<one sentence>",
  "suggest_master": <true|false>,
  "beneficiary_name": "<UPPER CASE clean name or null>"
}"""


def call_groq(
    client: Groq,
    description: str,
    direction: str,
    amount: float,
    retrieved: list[dict[str, Any]],
) -> Optional[dict[str, Any]]:
    """Call Groq LLM with RAG context. Returns parsed JSON or None."""
    examples_text = ""
    if retrieved:
        examples_text = "\n\nSIMILAR PAST TRANSACTIONS (from DPL's own data):\n"
        for i, ex in enumerate(retrieved, 1):
            examples_text += (
                f"{i}. [{ex['head']}] sim={ex['similarity']:.2f} | {ex['description'][:80]}\n"
            )

    user_msg = (
        f"Direction: {direction} | Amount: Rs {amount:,.0f}\n"
        f"Description: {description}"
        f"{examples_text}\n\n"
        "Classify this transaction. Return JSON only."
    )

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user", "content": user_msg},
            ],
            temperature=0.1,
            max_tokens=600,
            # GPT-OSS is a reasoning model: it spends output tokens on a
            # hidden reasoning pass before the final answer. Left at
            # defaults, that reasoning pass alone can consume the whole
            # token budget, leaving message.content empty (observed as
            # "Expecting value: line 1 column 1 (char 0)" from json.loads).
            # reasoning_effort="low" keeps that pass short (a single-label
            # classification doesn't need deep reasoning), and
            # reasoning_format="hidden" drops the reasoning trace from the
            # response entirely so content always holds just the answer.
            reasoning_effort="low",
            reasoning_format="hidden",
            response_format={"type": "json_object"},
        )
        raw = response.choices[0].message.content.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        return json.loads(raw)
    except Exception as exc:
        print(f"  [WARN] Groq API error: {exc}")
        return None


# ---------------------------------------------------------------------------
# BU / Type / TCP resolution
# ---------------------------------------------------------------------------

def _resolve_fields(
    head: str,
    own_bu: str,
    own_stage: Optional[str],
) -> dict[str, str]:
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
    if head == "Imprest" or own_stage == "Free":
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


# ---------------------------------------------------------------------------
# Beneficiary Master suggestion
# ---------------------------------------------------------------------------

def _get_existing_master_names(spreadsheet: gspread.Spreadsheet) -> set[str]:
    """Return all names already in the Beneficiary Master tab (any status)."""
    try:
        ws = spreadsheet.worksheet(MASTER_TAB_NAME)
        rows = ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        return set()
    if len(rows) < 2:
        return set()
    hdr = rows[0]
    if "BENEFICIARY NAME" not in hdr:
        return set()
    ni = hdr.index("BENEFICIARY NAME")
    return {row[ni].strip().upper() for row in rows[1:] if len(row) > ni and row[ni].strip()}


def _ensure_status_column(spreadsheet: gspread.Spreadsheet) -> None:
    """Add STATUS column to Beneficiary Master if it doesn't exist yet."""
    try:
        ws = spreadsheet.worksheet(MASTER_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        return
    hdr = ws.row_values(1)
    if "STATUS" not in hdr:
        next_col = len(hdr) + 1
        ws.update_cell(1, next_col, "STATUS")
        # Mark all existing rows as Confirmed
        existing = ws.get_all_values()
        if len(existing) > 1:
            updates = [
                gspread.cell.Cell(row=i + 2, col=next_col, value=MASTER_STATUS_CONFIRMED)
                for i in range(len(existing) - 1)
                if any(c.strip() for c in existing[i + 1])
            ]
            if updates:
                ws.update_cells(updates, value_input_option="RAW")
        print(f"  [MASTER] Added STATUS column, marked {len(updates)} existing rows as Confirmed.")


def suggest_to_master(
    spreadsheet: gspread.Spreadsheet,
    suggestions: list[dict[str, str]],
) -> None:
    """Write AI-suggested beneficiaries to the master tab in light blue."""
    if not suggestions:
        return

    try:
        ws = spreadsheet.worksheet(MASTER_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  [WARN] '{MASTER_TAB_NAME}' tab not found — skipping master suggestions.")
        return

    hdr = ws.row_values(1)
    existing_names = _get_existing_master_names(spreadsheet)

    import datetime
    today = datetime.date.today().strftime("%d-%b-%Y")

    # Ensure STATUS column exists
    if "STATUS" not in hdr:
        _ensure_status_column(spreadsheet)
        ws = spreadsheet.worksheet(MASTER_TAB_NAME)
        hdr = ws.row_values(1)

    # Build each row by header name, not fixed position - a positional list
    # here previously assumed a 6-column layout ([name, head, "AI
    # Suggested", "AI (RAG)", today, status]) that predates Head 2/Head 3
    # being added to the sheet, so every value after Head 1 landed 2
    # columns too early (e.g. "AI Suggested" written into Head 2, the
    # STATUS value written into ADDED BY). Any column not set below (Head
    # 2/Head 3, ACCOUNT NUMBER, etc.) is simply left blank.
    new_rows = []
    for s in suggestions:
        name = s.get("name", "").strip().upper()
        head = s.get("head", "").strip()
        if not name or not head or name in existing_names:
            continue
        row_values = {
            "BENEFICIARY NAME": name,
            "Head 1": head,
            "ADDED BY": "AI (RAG)",
            "DATE ADDED": today,
            "STATUS": MASTER_STATUS_AI,
        }
        new_row = [""] * len(hdr)
        for col_name, value in row_values.items():
            if col_name in hdr:
                new_row[hdr.index(col_name)] = value
        new_rows.append(new_row)
        existing_names.add(name)

    if not new_rows:
        return

    start_row = len(ws.get_all_values()) + 1
    ws.append_rows(new_rows, value_input_option="RAW")

    # Color AI-suggested rows light blue
    light_blue = {"red": 0.78, "green": 0.87, "blue": 0.95}
    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": start_row - 1 + i,
                    "endRowIndex": start_row + i,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(hdr),
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": light_blue
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
            }
        }
        for i in range(len(new_rows))
    ]
    try:
        ws.spreadsheet.batch_update({"requests": requests})
    except Exception as exc:
        print(f"  [WARN] Could not apply blue color to suggestions: {exc}")

    print(f"  [MASTER] Added {len(new_rows)} AI-suggested beneficiaries (light blue) for team review.")


# ---------------------------------------------------------------------------
# Color AI rows orange
# ---------------------------------------------------------------------------

def _mark_ai_rows(
    worksheet: gspread.Worksheet,
    row_numbers: list[int],
    col_indices: dict[str, int],
) -> None:
    if not row_numbers:
        return
    target_cols = [
        BUSINESS_UNIT_COLUMN, HEAD_COLUMN,
        TYPE_RERA_IDW_COLUMN, TCP_HEAD_COLUMN, NARRATION_COLUMN,
        CONFIDENCE_COLUMN, REASON_COLUMN,
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
                        "textFormat": {"foregroundColor": AI_TEXT_COLOR}
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
        print(f"  [WARN] Could not apply orange color: {exc}")


# ---------------------------------------------------------------------------
# Per-worksheet processor
# ---------------------------------------------------------------------------

def process_worksheet(
    worksheet: gspread.Worksheet,
    retriever: Retriever,
    groq_client: Groq,
    existing_master_names: Optional[set] = None,
) -> tuple[int, int, list[dict[str, str]]]:
    """Classify ? rows and verify keyword-classified rows for named payees.

    Two passes:
    1. ? rows — standard RAG classification.
    2. Keyword-classified rows (Vendor/Contractor/Professional) where a
       beneficiary name can be extracted and is NOT yet in the master.
       If RAG disagrees with the keyword, the row is corrected.

    Returns (resolved, still_unknown, master_suggestions).
    """
    all_values = worksheet.get_all_values()
    if not all_values:
        return 0, 0, []

    hdr = all_values[0]
    required = [HEAD_COLUMN, BUSINESS_UNIT_COLUMN, TYPE_RERA_IDW_COLUMN,
                TCP_HEAD_COLUMN, NARRATION_COLUMN]
    if any(c not in hdr for c in required):
        return 0, 0, []

    optional = [CONFIDENCE_COLUMN, REASON_COLUMN]
    tracked = required + [c for c in optional if c in hdr]
    col_idx = {c: hdr.index(c) + 1 for c in tracked}
    updates: list[gspread.cell.Cell] = []
    ai_rows: list[int] = []
    master_suggestions: list[dict[str, str]] = []
    resolved = 0
    still_unknown = 0

    if existing_master_names is None:
        existing_master_names = set()

    for offset, row in enumerate(all_values[1:]):
        sheet_row = offset + 2
        if _is_row_empty(row):
            continue

        current_head = _get_cell(row, hdr, HEAD_COLUMN)
        is_unknown = current_head == UNKNOWN_MAPPING_VALUE

        # Pass 2: verify keyword-classified rows for named payees not in master
        is_keyword_verify = (
            current_head in KEYWORD_CLASSIFIED_HEADS
            and _extract_name_from_desc(_get_cell(row, hdr, "DESCRIPTION") or "") is not None
            and (_extract_name_from_desc(_get_cell(row, hdr, "DESCRIPTION") or "") or "").upper()
            not in existing_master_names
        )

        if not is_unknown and not is_keyword_verify:
            continue

        description = _get_cell(row, hdr, "DESCRIPTION")
        if not description:
            continue

        account_number = _get_cell(row, hdr, "Account Number")
        deposits = _to_float(_get_cell(row, hdr, "CREDITS"))
        withdrawals = _to_float(_get_cell(row, hdr, "DEBITS"))
        amount = _parse_amount(
            _get_cell(row, hdr, "CREDITS"),
            _get_cell(row, hdr, "DEBITS"),
        )
        direction = "Credit (incoming)" if deposits > 0 else "Debit (outgoing)"

        own_account = _get_accounts_by_number().get(account_number, {})
        own_bu = own_account.get("business_unit") or next(
            (bu for sfx, bu in _ACCOUNT_BU_OVERRIDES.items()
             if account_number.endswith(sfx)),
            UNKNOWN_MAPPING_VALUE,
        )
        own_stage = own_account.get("account_stage") or next(
            (s for sfx, s in _ACCOUNT_STAGE_OVERRIDES.items()
             if account_number.endswith(sfx)),
            None,
        )

        # --- Retrieval ---
        retrieved = retriever.retrieve(description)

        print(f"  [{worksheet.title}] row {sheet_row}: {description[:70]}")
        if retrieved:
            print(f"    Top match: [{retrieved[0]['head']}] sim={retrieved[0]['similarity']:.2f}"
                  f" | {retrieved[0]['description'][:55]}")

        # --- Generation ---
        result = call_groq(groq_client, description, direction, amount, retrieved)

        if result is None or result.get("head") in (None, UNKNOWN_MAPPING_VALUE, ""):
            print(f"    -> AI: no result")
            if is_unknown:
                still_unknown += 1
            continue

        ai_head = result["head"]
        if ai_head not in VALID_HEADS:
            print(f"    -> AI: invalid head '{ai_head}' — skipping")
            if is_unknown:
                still_unknown += 1
            continue

        # For keyword-verify rows, skip if AI agrees with the existing head
        if is_keyword_verify and ai_head == current_head:
            print(f"    -> AI confirms: {ai_head} (keyword was correct)")
            continue

        if is_keyword_verify and ai_head != current_head:
            print(f"    -> AI OVERRIDE: '{current_head}' → '{ai_head}' | {result.get('reason', '')}")
        else:
            print(f"    -> AI: {ai_head} | {result.get('reason', '')}")

        # --- Master suggestion ---
        if result.get("suggest_master"):
            bname = result.get("beneficiary_name") or _extract_name_from_desc(description)
            if bname:
                master_suggestions.append({"name": bname, "head": ai_head})
                print(f"    -> Suggest master: {bname} = {ai_head}")

        fields = _resolve_fields(ai_head, own_bu, own_stage)
        new_narration = generate_narration(
            description, ai_head, amount,
            business_unit=fields["business_unit"],
            type_rera_idw=fields["type_rera_idw"],
            deposits=deposits,
            withdrawals=withdrawals,
            own_account_number=account_number,
        )

        ai_reason = f"RAG AI (Groq Llama): {result.get('reason', 'classified by AI')} — verify this row"
        if is_keyword_verify:
            ai_reason = f"RAG AI OVERRIDE from '{current_head}': {result.get('reason', '')} — verify this row"

        for col_name, value in [
            (BUSINESS_UNIT_COLUMN, fields["business_unit"]),
            (HEAD_COLUMN, ai_head),
            (TYPE_RERA_IDW_COLUMN, fields["type_rera_idw"]),
            (TCP_HEAD_COLUMN, fields["tcp_head"]),
            (NARRATION_COLUMN, new_narration),
            (CONFIDENCE_COLUMN, "Low"),
            (REASON_COLUMN, ai_reason),
        ]:
            if col_name in col_idx:
                updates.append(
                    gspread.cell.Cell(row=sheet_row, col=col_idx[col_name], value=value)
                )
        ai_rows.append(sheet_row)
        resolved += 1

    if updates:
        worksheet.update_cells(updates, value_input_option="RAW")
        _mark_ai_rows(worksheet, ai_rows, col_idx)

    return resolved, still_unknown, master_suggestions


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run_rag_classifier(
    credentials_path: Path = DEFAULT_CREDENTIALS,
    spreadsheet: Optional[gspread.Spreadsheet] = None,
) -> tuple[int, int]:
    """Run the RAG classifier on all account tabs.

    Returns (total_resolved, total_still_unknown).
    Can accept a pre-opened spreadsheet to skip re-authentication.
    """
    api_key = os.environ.get("GROQ_API_KEY")
    if not api_key:
        print("ERROR: GROQ_API_KEY environment variable not set.")
        print("Get a free key at: https://console.groq.com")
        return 0, 0

    groq_client = Groq(api_key=api_key)

    if spreadsheet is None:
        client = get_gspread_client(credentials_path)
        spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    # Ensure Beneficiary Master has STATUS column
    _ensure_status_column(spreadsheet)

    # Build retrieval index from all classified rows
    print("Building retrieval index...")
    descriptions, heads = load_training_data(spreadsheet)
    if not descriptions:
        print("No training data found. Run classify_transactions.py first.")
        return 0, 0

    retriever = Retriever(descriptions, heads)
    print(f"Index built: {len(descriptions)} training rows.\n")

    # Load master names once — used by process_worksheet() to skip payees
    # already confirmed so keyword-verify pass doesn't touch master rows.
    existing_master_names = _get_existing_master_names(spreadsheet)

    all_suggestions: list[dict[str, str]] = []
    total_resolved = 0
    total_unknown = 0

    for ws in get_account_worksheets(spreadsheet):
        resolved, unknown, suggestions = process_worksheet(
            ws, retriever, groq_client, existing_master_names=existing_master_names
        )
        total_resolved += resolved
        total_unknown += unknown
        all_suggestions.extend(suggestions)

    # Write master suggestions
    if all_suggestions:
        print()
        suggest_to_master(spreadsheet, all_suggestions)

    print()
    print(f"Total AI-resolved (RAG): {total_resolved}")
    print(f"Total still unknown:     {total_unknown}")
    if total_resolved > 0:
        print("Orange text = AI-classified. Team should verify these rows.")
    if all_suggestions:
        print("Light blue rows added to Beneficiary Master tab for team review.")

    return total_resolved, total_unknown


def main() -> None:
    run_rag_classifier()


if __name__ == "__main__":
    main()
