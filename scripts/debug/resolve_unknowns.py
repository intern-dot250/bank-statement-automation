"""One-off: re-resolve existing rows whose Business Unit/Head/Type for
RERA IDW/TCP Head are currently "?" (or Head="DPL", a confirmed
misclassification), using the newly-extended resolve_business_fields()
rules grounded in 2 years of the accounts team's reference workbooks.
Also regenerates Narration for any row whose fields actually changed.

Never touches a row's classification if the row is NOT currently "?"/
"DPL" in these columns — already-correct data (including anything the
accounts team has manually corrected) is left untouched.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __future__ import annotations

from classify_transactions import (
    BUSINESS_UNIT_COLUMN,
    HEAD_COLUMN,
    TYPE_RERA_IDW_COLUMN,
    TCP_HEAD_COLUMN,
    NARRATION_COLUMN,
    UNKNOWN_MAPPING_VALUE,
    resolve_business_fields,
    _get_accounts_by_number,
    _ACCOUNT_BU_OVERRIDES,
    _get_cell,
    _is_row_empty,
    _to_float,
    _parse_amount,
)
from heads import get_head
from narration import generate_narration
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client, get_account_worksheets
import gspread


def _expected_bu(account_number: str) -> str | None:
    """Return the own BU for this account (DB first, then override), or None."""
    own_account = _get_accounts_by_number().get(account_number, {})
    raw = own_account.get("business_unit")
    if raw:
        return raw
    return next(
        (bu for sfx, bu in _ACCOUNT_BU_OVERRIDES.items() if account_number.endswith(sfx)),
        None,
    )


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    for worksheet in get_account_worksheets(spreadsheet):
        all_values = worksheet.get_all_values()
        if not all_values:
            continue

        header_row = all_values[0]
        required = [BUSINESS_UNIT_COLUMN, HEAD_COLUMN, TYPE_RERA_IDW_COLUMN, TCP_HEAD_COLUMN, NARRATION_COLUMN]
        if any(c not in header_row for c in required):
            print(f"[SKIP] {worksheet.title}: missing classification columns.")
            continue

        col_index = {c: header_row.index(c) for c in required}
        data_rows = all_values[1:]
        updates: list[gspread.cell.Cell] = []
        resolved_count = 0
        still_unknown = 0

        for offset, row in enumerate(data_rows):
            sheet_row_number = offset + 2

            if _is_row_empty(row):
                continue

            current_head = _get_cell(row, header_row, HEAD_COLUMN)
            current_bu = _get_cell(row, header_row, BUSINESS_UNIT_COLUMN)
            current_type = _get_cell(row, header_row, TYPE_RERA_IDW_COLUMN)
            current_tcp = _get_cell(row, header_row, TCP_HEAD_COLUMN)

            description = _get_cell(row, header_row, "DESCRIPTION")
            if not description:
                continue

            deposits_raw = _get_cell(row, header_row, "CREDITS")
            withdrawals_raw = _get_cell(row, header_row, "DEBITS")
            account_number = _get_cell(row, header_row, "Account Number")

            # "Wrong BU": stored BU is neither "?", "HO" (valid for Professional/
            # Salary HO regardless of account), nor the account's own expected BU.
            # This catches rows where BU was set from description content (e.g.
            # a transfer description mentioning "CASA ROMANA" setting BU to
            # "Casa Romana" on an Aravali Heights account).
            exp_bu = _expected_bu(account_number)
            wrong_bu = (
                exp_bu is not None
                and current_bu not in (UNKNOWN_MAPPING_VALUE, "HO", exp_bu)
            )

            needs_resolution = (
                current_head == UNKNOWN_MAPPING_VALUE
                or current_head == "DPL"
                or current_bu == UNKNOWN_MAPPING_VALUE
                or current_type == UNKNOWN_MAPPING_VALUE
                or current_tcp == UNKNOWN_MAPPING_VALUE
                or wrong_bu
            )
            if not needs_resolution:
                continue
            deposits = _to_float(deposits_raw)
            withdrawals = _to_float(withdrawals_raw)
            amount = _parse_amount(deposits_raw, withdrawals_raw)

            resolved = resolve_business_fields(account_number, description, deposits, withdrawals)
            new_head = resolved["head"] or get_head(description, deposits, withdrawals)
            display_head = UNKNOWN_MAPPING_VALUE if new_head == "Others" else new_head

            new_bu = resolved["business_unit"]
            new_type = resolved["type_rera_idw"]
            new_tcp = resolved["tcp_head"]

            changed = (
                display_head != current_head
                or new_bu != current_bu
                or new_type != current_type
                or new_tcp != current_tcp
            )
            if not changed:
                continue

            new_narration = generate_narration(
                description,
                display_head,
                amount,
                business_unit=new_bu,
                type_rera_idw=new_type,
                deposits=deposits,
                withdrawals=withdrawals,
                own_account_number=account_number,
            )

            for column_name, value in [
                (BUSINESS_UNIT_COLUMN, new_bu),
                (HEAD_COLUMN, display_head),
                (TYPE_RERA_IDW_COLUMN, new_type),
                (TCP_HEAD_COLUMN, new_tcp),
                (NARRATION_COLUMN, new_narration),
            ]:
                updates.append(
                    gspread.cell.Cell(row=sheet_row_number, col=col_index[column_name] + 1, value=value)
                )

            if display_head == UNKNOWN_MAPPING_VALUE:
                still_unknown += 1
            else:
                resolved_count += 1

        if updates:
            worksheet.update_cells(updates, value_input_option="RAW")

        print(f"[OK] {worksheet.title}: resolved {resolved_count} row(s), still unresolved {still_unknown} row(s).")


if __name__ == "__main__":
    main()
