"""Phase 2: Classify a bank transaction into a business Head.

Strict rule: payment modes (IMPS/UPI/NEFT/RTGS) are NOT head identifiers —
they only describe how money moved, not why. A Head is only assigned when
Description (+ Direction, where required) gives strong confidence. Anything
uncertain falls back to Others rather than being guessed.

Business heads:
    1. Salary
    2. Customer Cancellation
    3. Bank Charges
    4. Customer Collection
    5. Vendor Payment
    6. Internal
    7. Others (fallback when nothing matches confidently)
"""

from __future__ import annotations

import logging
import re

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Head labels
# ---------------------------------------------------------------------------

HEAD_SALARY = "Salary"
HEAD_CUSTOMER_CANCELLATION = "Customer Cancellation"
HEAD_BANK_CHARGES = "Bank Charges"
HEAD_CUSTOMER_COLLECTION = "Customer Collection"
HEAD_VENDOR_PAYMENT = "Vendor Payment"
HEAD_INTERNAL = "Internal"
HEAD_OTHERS = "Others"

# ---------------------------------------------------------------------------
# Pattern keywords
# ---------------------------------------------------------------------------

SALARY_KEYWORDS = ["SALARY"]

CANCELLATION_KEYWORDS = ["CANCEL", "REFUND", "REVERSAL"]

BANK_CHARGES_KEYWORDS = ["CHARGE", "GST", "FEE"]

# Requires Deposits > 0
COLLECTION_KEYWORDS = [
    "ABHIJIT",
    "DWARKADHIS",
    "BOOKING",
    "INSTALLMENT",
    "CUSTOMER",
    "PROJECT",
    "RERA",
]

# Requires Withdrawals > 0
VENDOR_KEYWORDS = [
    "CARRIER",
    "PAINT",
    "SUPPLIER",
    "VENDOR",
    "CONTRACTOR",
    "TRANSPORT",
]

# Only explicit self/internal-transfer language — never inferred from
# payment mode alone (IMPS/UPI/NEFT/RTGS are excluded on purpose).
INTERNAL_KEYWORDS = [
    "SELF",
    "OWN",
    "INTERNAL",
    "TO SELF",
    "TRANSFER TO OWN",
]


def _contains_word(desc_upper: str, keyword: str) -> bool:
    """Match keyword as a whole word/phrase, avoiding substring false
    positives (e.g. "OWN" must not match inside "UNKNOWN")."""
    pattern = r"\b" + re.escape(keyword) + r"\b"
    return re.search(pattern, desc_upper) is not None


def get_head(description: str, deposits: float, withdrawals: float) -> str:
    """Return the business Head for a transaction.

    Args:
        description: Raw bank transaction description.
        deposits: Deposit amount for the transaction (0 if none).
        withdrawals: Withdrawal amount for the transaction (0 if none).

    Priority order:
        1. Salary
        2. Customer Cancellation
        3. Bank Charges
        4. Customer Collection (requires Deposits > 0 + known pattern)
        5. Vendor Payment (requires Withdrawals > 0 + known pattern)
        6. Internal (only explicit self/internal-transfer language)
        7. Others — no confident match
    """
    if not description:
        logger.debug("Empty description — defaulting to %s.", HEAD_OTHERS)
        return HEAD_OTHERS

    desc_upper = description.upper()

    logger.debug(
        "Classifying description=%r deposits=%r withdrawals=%r",
        description,
        deposits,
        withdrawals,
    )

    # Rule 1: Salary
    if any(kw in desc_upper for kw in SALARY_KEYWORDS):
        return HEAD_SALARY

    # Rule 2: Customer Cancellation
    if any(kw in desc_upper for kw in CANCELLATION_KEYWORDS):
        return HEAD_CUSTOMER_CANCELLATION

    # Rule 3: Bank Charges
    if any(kw in desc_upper for kw in BANK_CHARGES_KEYWORDS):
        return HEAD_BANK_CHARGES

    # Rule 4: Customer Collection — Deposit + known collection pattern
    if deposits > 0 and any(kw in desc_upper for kw in COLLECTION_KEYWORDS):
        return HEAD_CUSTOMER_COLLECTION

    # Rule 5: Vendor Payment — Withdrawal + known vendor pattern
    if withdrawals > 0 and any(kw in desc_upper for kw in VENDOR_KEYWORDS):
        return HEAD_VENDOR_PAYMENT

    # Rule 6: Internal — only explicit self/internal-transfer language
    if any(_contains_word(desc_upper, kw) for kw in INTERNAL_KEYWORDS):
        return HEAD_INTERNAL

    # Rule 7: No confident match
    logger.debug("No confident match for description=%r — defaulting to Others.", description)
    return HEAD_OTHERS


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

    test_cases = [
        # (description, deposits, withdrawals, expected)
        ("SALARY CREDIT FOR JUNE", 60000.0, 0.0, HEAD_SALARY),
        ("BOOKING CANCEL REFUND", 0.0, 25000.0, HEAD_CUSTOMER_CANCELLATION),
        ("REVERSAL OF WRONG CREDIT", 0.0, 1200.0, HEAD_CUSTOMER_CANCELLATION),
        ("BANK CHARGES GST DEBIT", 0.0, 150.0, HEAD_BANK_CHARGES),
        ("IMPS/ABHIJIT SHARMA/XXX2427", 15000.0, 0.0, HEAD_CUSTOMER_COLLECTION),
        ("RERA INSTALLMENT FROM CUSTOMER", 250000.0, 0.0, HEAD_CUSTOMER_COLLECTION),
        ("DWARKADHIS PROJECT PAYMENT RECEIVED", 50000.0, 0.0, HEAD_CUSTOMER_COLLECTION),
        ("NEFT TO SUPPLIER FOR PAINT", 0.0, 40000.0, HEAD_VENDOR_PAYMENT),
        ("RTGS CONTRACTOR TRANSPORT CHARGES PAID", 0.0, 20000.0, HEAD_BANK_CHARGES),
        ("SELF TRANSFER TO OWN ACCOUNT", 0.0, 20000.0, HEAD_INTERNAL),
        ("IMPS/UNKNOWN PERSON/XXX1234", 15000.0, 0.0, HEAD_OTHERS),
        ("NEFT RANDOM TEXT NO PATTERN", 0.0, 5000.0, HEAD_OTHERS),
        ("UNKNOWN TEXT NO KEYWORDS", 3000.0, 0.0, HEAD_OTHERS),
        ("", 0.0, 0.0, HEAD_OTHERS),
    ]

    passed = 0
    for description, deposits, withdrawals, expected in test_cases:
        result = get_head(description, deposits, withdrawals)
        status = "PASS" if result == expected else "FAIL"
        if status == "PASS":
            passed += 1
        print(f"[{status}] {description!r} (D={deposits}, W={withdrawals}) -> {result} (expected {expected})")

    print(f"\n{passed}/{len(test_cases)} test cases passed.")
