"""Generate a human-readable Narration matching the accounts team's own
sheet format (see their Google Sheets LET() formula, reference: "Copy of
DPL Bank Statements 2026-27.xlsx"). Reproduces the same three narration
styles they use — Internal Transfer, Receipt Credit, Payment Disbursement —
built from the same fields their formula reads (Ref, Type for RERA IDW /
"Type", Business Unit / "BU", Head), rather than the old free-text
template system.

Public API:

    generate_narration(
        description, head, amount, *,
        business_unit="?", type_rera_idw="?",
        deposits=0.0, withdrawals=0.0,
        own_account_number="", reference=None,
    ) -> str
"""

from __future__ import annotations

import logging
from typing import Optional

from description_parser import parse_description

logger = logging.getLogger(__name__)

UNKNOWN_VALUE = "?"

# Our own group companies' names, as they appear (verbatim, upper-cased)
# inside transaction descriptions when a transfer is between two of our
# own tracked accounts — mirrors the accounts team's own formula, which
# hard-codes a SEARCH("DWARKADHIS", ...) check for this same purpose.
OWN_COMPANY_KEYWORDS = ["DWARKADHIS", "AMBITION COLONISERS"]


def _format_amount(amount: float) -> str:
    try:
        return f"{amount:,.2f}"
    except (TypeError, ValueError):
        return "0.00"


def _last4(text: str) -> str:
    """Last 4 characters of the trimmed text, matching the accounts
    formula's RIGHT(TRIM(...), 4) — used to identify the counterparty
    account by the tail digits of its account number embedded in the
    description."""
    trimmed = (text or "").strip()
    return trimmed[-4:] if trimmed else ""


def _is_internal_company_transfer(description: str) -> bool:
    upper = (description or "").upper()
    return any(keyword in upper for keyword in OWN_COMPANY_KEYWORDS)


def generate_narration(
    description: str,
    head: str,
    amount: float,
    *,
    business_unit: str = UNKNOWN_VALUE,
    type_rera_idw: str = UNKNOWN_VALUE,
    deposits: float = 0.0,
    withdrawals: float = 0.0,
    own_account_number: str = "",
    reference: Optional[str] = None,
) -> str:
    """Generate a narration in the accounts team's own format.

    Args:
        description: Raw bank transaction description.
        head: Business Head assigned to the transaction.
        amount: Transaction amount (deposit or withdrawal) — used only for
            the emergency fallback text; the accounts team's own format
            does not include the amount in the narration itself.
        business_unit: Value already resolved for the Business Unit column.
        type_rera_idw: Value already resolved for the Type for RERA IDW column.
        deposits: Credit amount for this row (0.0 if this is a debit row).
        withdrawals: Debit amount for this row (0.0 if this is a credit row).
        own_account_number: This row's own account number, used to label
            "our side" of an internal transfer.
        reference: Reference/Cheque No value already available for this
            row, if any — falls back to description_parser's extracted
            reference when not provided.

    Returns:
        A human-readable narration string, following the same three
        styles (Internal Transfer / Receipt Credit / Payment
        Disbursement) as the accounts team's own sheet. Never raises.
    """
    try:
        parsed = parse_description(description) if description else {}
    except Exception as exc:
        logger.warning("description_parser raised %s — continuing with no parsed fields.", exc)
        parsed = {}

    party = parsed.get("party") or "Party"
    ref = reference or parsed.get("reference") or "N/A"
    last4 = _last4(description)
    own_last4 = own_account_number[-4:] if own_account_number else ""
    own_label = f"x{own_last4}" if own_last4 else "our account"

    if head == "Internal":
        if _is_internal_company_transfer(description):
            if deposits > 0:
                text = f"Internal Fund Transfer (From x{last4} to {own_label})"
            else:
                text = f"Internal Fund Transfer (From {own_label} to x{last4})"
        else:
            text = "Internal Transfer"
        text += f" | Ref: {ref} | Type: {type_rera_idw} | BU: {business_unit} | Head: Internal"

    elif deposits > 0:
        if type_rera_idw == "Customer Collection":
            text = f"Receipt Credit (Collection) | Ref: {ref} | From: {party} | BU: {business_unit}"
        else:
            text = (
                f"Receipt Credit from x{last4} (Purpose: {business_unit}) "
                f"| Ref: {ref} | BU: {business_unit} | Head: {head}"
            )

    else:
        purpose = "Salary" if "salary" in (head or "").lower() else business_unit
        text = (
            f"Payment Disbursement (Purpose: {purpose}) | To: {party} "
            f"| Ref: {ref} | BU: {business_unit} | Head: {head}"
        )

    logger.debug(
        "Generated narration for head=%r amount=%r description=%r -> %r",
        head, amount, description, text,
    )

    return text
