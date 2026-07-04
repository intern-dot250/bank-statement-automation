"""Phase 2: Classify a bank transaction Description into a business Head."""

# Order matters: more specific categories are checked before generic ones.
HEAD_RULES = [
    ("Cancellation", ["CANCEL", "REFUND", "REVERSAL"]),
    ("RERA Collection", ["RERA", "BOOKING", "INSTALLMENT"]),
    ("Salary", ["SALARY"]),
    ("Bank Charges", ["CHARGE", "GST", "FEE"]),
    ("Vendor Payment", ["VENDOR", "PAYMENT"]),
    ("Internal Collection", ["IMPS", "UPI", "NEFT", "RTGS"]),
]

DEFAULT_HEAD = "Others"


def get_head(description: str) -> str:
    """Return the business Head for a transaction Description."""
    if not description:
        return DEFAULT_HEAD

    desc_upper = description.upper()

    for head, keywords in HEAD_RULES:
        if any(keyword in desc_upper for keyword in keywords):
            return head

    return DEFAULT_HEAD
