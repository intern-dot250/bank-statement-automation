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


def main():
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(MASTER_SHEET_ID)

    existing = [ws.title for ws in ss.worksheets()]
    if TAB_NAME in existing:
        print(f"'{TAB_NAME}' tab already exists — nothing to do.")
        return

    ws = ss.add_worksheet(title=TAB_NAME, rows=200, cols=len(HEADER))
    ws.update(range_name="A1", values=[HEADER])
    ws.format("A1:J1", {"textFormat": {"bold": True}})

    # Dropdown validation on STATUS (column G)
    import gspread.utils as gu
    ws.add_validation(
        "G2:G200",
        gu.ValidationConditionType.one_of_list,
        STATUS_VALUES,
        showCustomUi=True,
    )

    print(f"Created '{TAB_NAME}' tab with header row and STATUS dropdown.")


if __name__ == "__main__":
    main()
