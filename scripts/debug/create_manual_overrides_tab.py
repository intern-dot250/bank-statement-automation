"""One-off: create the "Manual Overrides" tab in the master spreadsheet,
with header row, if it doesn't already exist.

Lets the accounts team define recurring classification corrections
themselves (see classify_transactions.py's Rule 0 / _load_manual_overrides_cache),
without needing a developer to change code and redeploy.
"""
from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS, MASTER_SHEET_ID

TAB_NAME = "Manual Overrides"

HEADER = [
    "ACCOUNT NUMBER",
    "DESCRIPTION KEYWORD",
    "HEAD",
    "BUSINESS UNIT",
    "TYPE FOR RERA IDW",
    "TCP Head",
    "STATUS",
    "ADDED BY",
    "DATE ADDED",
    "NOTES",
]

STATUS_VALUES = ["Active", "Disabled"]

# Every HEAD value the pipeline can actually produce — Beneficiary Master's
# curated payee-role heads (web_app.py's BENEFICIARY_MASTER_HEADS) plus the
# non-payee-role heads classify_transactions.py's rules and
# config/heads_config.json also assign (Collection, Internal, Bank Charges,
# etc.) — so an override can correct ANY transaction's HEAD, not just a
# payee-identity one.
HEAD_VALUES = [
    "Vendor", "Contractor", "Salary Site", "Salary HO", "Professional",
    "Imprest", "Internal", "Legal & Proff.", "Statutory Dues",
    "Collection", "Bank Charges", "Credit Card", "HO - Advert/Mkt",
    "Cancellation", "Refund", "Refundable Security", "Commission",
    "Bounce", "Full & Final", "ROC Fees", "Wages", "Tax", "Loan",
    "Office Rent", "Others",
]


def _apply_validations(ws) -> None:
    import gspread.utils as gu
    ws.add_validation(
        "C2:C200",
        gu.ValidationConditionType.one_of_list,
        HEAD_VALUES,
        showCustomUi=True,
    )
    ws.add_validation(
        "G2:G200",
        gu.ValidationConditionType.one_of_list,
        STATUS_VALUES,
        showCustomUi=True,
    )


def main():
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(MASTER_SHEET_ID)

    existing = [ws.title for ws in ss.worksheets()]
    if TAB_NAME in existing:
        print(f"'{TAB_NAME}' tab already exists — applying/refreshing dropdowns only.")
        ws = ss.worksheet(TAB_NAME)
        _apply_validations(ws)
        print("Applied HEAD and STATUS dropdowns.")
        return

    ws = ss.add_worksheet(title=TAB_NAME, rows=200, cols=len(HEADER))
    ws.update(range_name="A1", values=[HEADER])
    ws.format("A1:J1", {"textFormat": {"bold": True}})
    _apply_validations(ws)

    print(f"Created '{TAB_NAME}' tab with header row, HEAD dropdown, and STATUS dropdown.")


if __name__ == "__main__":
    main()
