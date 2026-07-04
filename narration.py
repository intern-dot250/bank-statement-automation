"""Phase 2: Generate a human-readable Narration from a transaction's
Description, Head, and Amount."""

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)


def generate_narration(description: str, head: str, amount: float) -> str:
    """Generate a human-readable narration for a transaction.

    Args:
        description: Raw bank transaction description.
        head: Business Head assigned to the transaction (see heads.py).
        amount: Transaction amount (deposit or withdrawal).

    Returns:
        A human-readable narration string describing the transaction.
    """
    logger.debug(
        "Generating narration for head=%r, amount=%r, description=%r",
        head,
        amount,
        description,
    )

    if head == "Internal":
        return "Internal fund movement recorded."

    if head == "Customer Collection":
        return "Customer payment received."

    if head == "Customer Cancellation":
        return "Customer refund/cancellation processed."

    return "General banking transaction."


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

    test_cases = [
        ("IMPS/ABHIJIT SHARMA/XXX2427", "Internal", 15000.0),
        ("RERA BOOKING INSTALLMENT", "Customer Collection", 250000.0),
        ("BOOKING CANCEL REFUND", "Customer Cancellation", 5000.0),
        ("RANDOM TEXT", "Others", 100.0),
    ]

    for description, head, amount in test_cases:
        narration = generate_narration(description, head, amount)
        print(f"{description!r} | {head} | {amount} -> {narration}")
