"""Sweep all account tabs for transactions involving the 6 dual-head
Beneficiary Master entries (Diksha Sharma, Mukesh Kumar, Ram Kishan, Ravi
Vats, Sher Singh, Yogesh Singh), re-resolve HEAD/BU/Type/TCP/CONFIDENCE/
REASON via the updated classify_transactions rules, write corrections, and
mark all matched rows navy blue."""
import gspread.utils as gu
import classify_transactions as ct
from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS, MASTER_SHEET_ID

ACCOUNT_TABS = [
    "YES BANK - 0264",
    "YES BANK - 2477",
    "YES BANK - 0490",
    "YES BANK - 0377",
    "YES BANK - 2314",
    "YES BANK - 2457",
]

DUAL_HEAD_NAMES = {
    "DIKSHA SHARM A", "DIKSHA SHARMA",
    "MUKESH KUMAR",
    "RAM KISHAN",
    "RAVI VATS",
    "SHER SINGH",
    "YOGESH SING H", "YOGESH SINGH",
}


def main():
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(MASTER_SHEET_ID)

    for tab in ACCOUNT_TABS:
        ws = ss.worksheet(tab)
        hdr = ws.row_values(1)
        col_indices = {name: i + 1 for i, name in enumerate(hdr)}
        all_vals = ws.get_all_values()

        i_desc = hdr.index("DESCRIPTION")
        i_cr = hdr.index("CREDITS")
        i_db = hdr.index("DEBITS")
        i_acct = hdr.index("Account Number")

        dual_rows = []
        requests = []

        for r_idx, row in enumerate(all_vals[1:], start=2):
            if len(row) <= i_desc or not row[i_desc]:
                continue
            description = row[i_desc]
            name = ct._extract_beneficiary_name(description)
            if not name or name not in DUAL_HEAD_NAMES:
                continue

            deposits = ct._to_float(row[i_cr]) if len(row) > i_cr else 0.0
            withdrawals = ct._to_float(row[i_db]) if len(row) > i_db else 0.0
            account_number = row[i_acct] if len(row) > i_acct else ""

            resolved = ct.resolve_business_fields(
                account_number, description, deposits, withdrawals, spreadsheet=ss
            )
            head = resolved["head"] or ct.get_head(description, deposits, withdrawals)
            display_head = ct.UNKNOWN_MAPPING_VALUE if head == "Others" else head
            reason_text = ct._build_reason_text(display_head, resolved)

            row_values = {
                "HEAD": display_head,
                "BUSINESS UNIT": resolved["business_unit"],
                "TYPE FOR RERA IDW": resolved["type_rera_idw"],
                "TCP Head": resolved["tcp_head"],
                "CONFIDENCE": resolved.get("confidence", "Low"),
                "REASON": reason_text,
            }
            for col_name, value in row_values.items():
                if col_name in col_indices:
                    a1 = gu.rowcol_to_a1(r_idx, col_indices[col_name])
                    requests.append({"range": a1, "values": [[value]]})

            dual_rows.append((r_idx, name, display_head))

        if requests:
            ws.batch_update(requests)
        if dual_rows:
            ct._mark_dual_head_rows(ws, [r for r, _, _ in dual_rows], col_indices)

        print(f"{tab}: {len(dual_rows)} dual-head rows updated")
        for r_idx, name, head in dual_rows:
            print(f"   row {r_idx}: {name} -> {head}")


if __name__ == "__main__":
    main()
