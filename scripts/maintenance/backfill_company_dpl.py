"""One-off: set company='DPL' on any account_credentials row with a NULL
or blank `company` field.

Multi-company routing (see upload_to_sheets.get_spreadsheet_id_for_company())
already treats a blank/NULL company as "DPL" everywhere, so this script is
not required for correctness — it's a cleanliness pass so every existing
account has an explicit company on file, matching what new accounts are
required to have going forward.

Safe to re-run: an account that already has a company set is left untouched.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import credentials_store
from upload_to_sheets import DEFAULT_COMPANY

RECORDS_PATH = Path(__file__).resolve().parent.parent.parent / "data" / "records.json"


def main() -> None:
    accounts = credentials_store.list_credentials(RECORDS_PATH)

    updated = 0
    for acc in accounts:
        if acc.get("id") is None:
            continue  # file-fallback account, no company field to backfill
        if acc.get("company"):
            continue  # already set

        credentials_store.update_credential(
            credential_id=acc["id"],
            bank_name=acc["bank_name"],
            account_number=acc["account_number"],
            password=None,  # leave unchanged
            business_unit=acc.get("business_unit"),
            company=DEFAULT_COMPANY,
            financial_year=acc.get("financial_year"),
            account_stage=acc.get("account_stage"),
            update_account_stage=True,
        )
        print(f"[OK] {acc['bank_name']} - {acc['account_number']}: company set to '{DEFAULT_COMPANY}'.")
        updated += 1

    print(f"\nDone. {updated} account(s) backfilled.")


if __name__ == "__main__":
    main()
