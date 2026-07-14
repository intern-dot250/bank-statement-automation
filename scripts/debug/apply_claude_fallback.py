"""Phase 3: Claude API fallback for rows that still have HEAD = '?'
after all rule-based classification and beneficiary master lookup.

Color coding after this runs:
  Red    = auto-classified by rules (high confidence)
  Orange = AI-suggested by Claude (needs team verification)
  Black  = verified by accounts team

Run this after resolve_unknowns.py when ? rows remain:
  py -3 apply_claude_fallback.py

Requires environment variable: ANTHROPIC_API_KEY=sk-ant-...
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __future__ import annotations

import json
import os
from typing import Any, Optional

import anthropic
import gspread

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

# Orange text = AI-suggested, needs team verification
AI_TEXT_COLOR = {"red": 1.0, "green": 0.5, "blue": 0.0}

VALID_HEADS = [
    "Vendor",
    "Contractor",
    "Professional",
    "Imprest",
    "Salary HO",
    "Salary Site",
    "Internal",
    "Collection",
    "Cancellation",
    "Statutory Dues",
    "HO - Advert/Mkt",
]

_SYSTEM_PROMPT = """\
You are a financial transaction classifier for DPL (Dwarkadhis Projects Limited), \
an Indian real estate company. Your job is to read a bank transaction description \
and assign the correct Head category.

Valid Head values and when to use them:
- Vendor: Payment to material/hardware/consumables suppliers (e.g. hardware store, paint shop)
- Contractor: Payment to construction contractors, sub-contractors, or daily-wage laborers
- Professional: Payment to CA firms, architects, consultants, legal professionals
- Imprest: Petty cash advance given to an employee for small expenses
- Salary HO: Head office staff monthly salary
- Salary Site: Site/construction staff monthly salary
- Internal: Transfer between the company's own bank accounts
- Collection: Incoming payment received from apartment buyers (customer)
- Cancellation: Refund issued to a customer who cancelled their apartment booking
- Statutory Dues: Government statutory payments — PF, ESI, TDS, Professional Tax, PTAX
- HO - Advert/Mkt: Marketing, advertising, publicity, hoarding, branding expenses

Rules:
1. If the description contains a person's name followed by a role (e.g. "RAM KUMAR VENDOR"), \
use that role's Head.
2. If the description mentions a government/statutory payment, use Statutory Dues.
3. If it is an incoming payment (UPI credit, NEFT CR, customer transfer), use Collection.
4. If the role is completely unclear, return head = "?" — do not guess randomly.

Respond ONLY with valid JSON (no explanation, no markdown):
{"head": "<exact head value from list above or ?>", "reason": "<one short sentence>"}
"""


def _build_user_message(
    description: str,
    account_number: str,
    deposits: float,
    withdrawals: float,
    own_stage: Optional[str],
    own_bu: str,
) -> str:
    direction = "Credit (incoming)" if deposits > 0 else "Debit (outgoing)"
    amount = deposits if deposits > 0 else withdrawals
    return (
        f"Account: ...{account_number[-4:]} | Stage: {own_stage or 'Unknown'} | "
        f"BU: {own_bu}\n"
        f"Direction: {direction} | Amount: Rs {amount:,.0f}\n"
        f"Description: {description}\n\n"
        f"Classify this transaction. Return JSON only."
    )


def call_claude(
    client: anthropic.Anthropic,
    description: str,
    account_number: str,
    deposits: float,
    withdrawals: float,
    own_stage: Optional[str],
    own_bu: str,
) -> Optional[dict[str, str]]:
    """Call Claude API and return parsed JSON result, or None on failure."""
    user_msg = _build_user_message(
        description, account_number, deposits, withdrawals, own_stage, own_bu
    )
    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=150,
            system=_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )
        raw = response.content[0].text.strip()
        # Strip markdown code fences if present
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        return json.loads(raw)
    except Exception as exc:
        print(f"  [WARN] Claude API error: {exc}")
        return None


def _resolve_fields_for_head(
    head: str,
    own_bu: str,
    own_stage: Optional[str],
) -> dict[str, str]:
    """Determine BU / Type / TCP for a Claude-suggested head using the same
    logic as rule-based classification."""
    if head in ("Salary HO", "Professional", "Statutory Dues", "HO - Advert/Mkt"):
        tcp = (
            "Other-Selling Expenses"
            if head == "HO - Advert/Mkt"
            else _HO_ADMIN_DEFAULTS["tcp_head"]
        )
        return {
            "business_unit": _HO_ADMIN_DEFAULTS["business_unit"],
            "type_rera_idw": _HO_ADMIN_DEFAULTS["type_rera_idw"],
            "tcp_head": tcp,
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


def _mark_ai_rows(
    worksheet: gspread.Worksheet,
    row_numbers: list[int],
    col_indices: dict[str, int],
) -> None:
    """Color AI-classified cells orange so accounts team knows to verify."""
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


def process_worksheet(
    worksheet: gspread.Worksheet,
    claude_client: anthropic.Anthropic,
) -> tuple[int, int]:
    """Process all ? rows in one worksheet. Returns (resolved, still_unknown)."""
    all_values = worksheet.get_all_values()
    if not all_values:
        return 0, 0

    hdr = all_values[0]
    required = [HEAD_COLUMN, BUSINESS_UNIT_COLUMN, TYPE_RERA_IDW_COLUMN,
                TCP_HEAD_COLUMN, NARRATION_COLUMN]
    if any(c not in hdr for c in required):
        print(f"  [SKIP] {worksheet.title}: missing classification columns.")
        return 0, 0

    col_idx = {c: hdr.index(c) + 1 for c in required}

    updates: list[gspread.cell.Cell] = []
    ai_rows: list[int] = []
    resolved = 0
    still_unknown = 0

    for offset, row in enumerate(all_values[1:]):
        sheet_row = offset + 2

        if _is_row_empty(row):
            continue

        current_head = _get_cell(row, hdr, HEAD_COLUMN)
        if current_head != UNKNOWN_MAPPING_VALUE:
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

        # Resolve own BU and stage using same logic as classify_transactions.py
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

        print(f"  [{worksheet.title}] row {sheet_row}: {description[:70]}...")
        result = call_claude(
            claude_client, description, account_number,
            deposits, withdrawals, own_stage, own_bu,
        )

        if result is None or result.get("head") in (None, UNKNOWN_MAPPING_VALUE, ""):
            print(f"    -> Claude: still unknown ({result})")
            still_unknown += 1
            continue

        ai_head = result["head"]
        if ai_head not in VALID_HEADS:
            print(f"    -> Claude returned invalid head '{ai_head}' — skipping.")
            still_unknown += 1
            continue

        fields = _resolve_fields_for_head(ai_head, own_bu, own_stage)
        reason = result.get("reason", "")
        print(f"    -> Claude: {ai_head} | {reason}")

        new_narration = generate_narration(
            description,
            ai_head,
            amount,
            business_unit=fields["business_unit"],
            type_rera_idw=fields["type_rera_idw"],
            deposits=deposits,
            withdrawals=withdrawals,
            own_account_number=account_number,
        )

        for col_name, value in [
            (BUSINESS_UNIT_COLUMN, fields["business_unit"]),
            (HEAD_COLUMN, ai_head),
            (TYPE_RERA_IDW_COLUMN, fields["type_rera_idw"]),
            (TCP_HEAD_COLUMN, fields["tcp_head"]),
            (NARRATION_COLUMN, new_narration),
        ]:
            updates.append(
                gspread.cell.Cell(
                    row=sheet_row,
                    col=col_idx[col_name],
                    value=value,
                )
            )
        ai_rows.append(sheet_row)
        resolved += 1

    if updates:
        worksheet.update_cells(updates, value_input_option="RAW")
        _mark_ai_rows(worksheet, ai_rows, col_idx)

    return resolved, still_unknown


def main() -> None:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: ANTHROPIC_API_KEY environment variable is not set.")
        print("Set it with:  set ANTHROPIC_API_KEY=sk-ant-...")
        return

    claude_client = anthropic.Anthropic(api_key=api_key)

    sheets_client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = sheets_client.open_by_key(MASTER_SHEET_ID)

    total_resolved = 0
    total_unknown = 0

    for worksheet in get_account_worksheets(spreadsheet):
        resolved, unknown = process_worksheet(worksheet, claude_client)
        total_resolved += resolved
        total_unknown += unknown
        if resolved > 0 or unknown > 0:
            print(
                f"[OK] {worksheet.title}: "
                f"AI-resolved {resolved}, still unknown {unknown}"
            )

    print()
    print(f"Total AI-resolved: {total_resolved}")
    print(f"Total still unknown (needs manual): {total_unknown}")
    if total_resolved > 0:
        print("Orange text = AI-suggested. Team should verify these rows.")


if __name__ == "__main__":
    main()
