"""Creates (or replaces) a 'Rules' worksheet in the AMB Google Sheet that
documents the AMB-specific auto-classification logic
(classify_transactions._resolve_amb_business_fields) so the accounts team
can verify each rule. Mirrors scripts/debug/create_rules_tab.py's
structure (DPL's own Rules tab), scoped to AMB's 4 confirmed accounts —
AMB CA AXIS-280 is listed as pending, not yet covered.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import gspread
from upload_to_sheets import DEFAULT_CREDENTIALS, get_gspread_client

AMB_SHEET_ID = "1kVMuah99dU8g3q9zsHxiTtBh7l7E8xIPrFoSmGQpywc"

# ---------------------------------------------------------------------------
# Rule data — matches classify_transactions._resolve_amb_business_fields()
# ---------------------------------------------------------------------------

RULES: list[dict] = [
    {
        "account": "AMB MASTER KVB-6375",
        "rule": "Bank Charges exception",
        "trigger": "DESCRIPTION contains 'SMS Charge' / 'Service Chrg' / 'Monthly Service' — checked before the credit/debit rule below.",
        "examples": "SMS Charges for JUN2026",
        "bu": "SW", "head": "Bank Charges", "type_rera_idw": "HO - Admin", "tcp_head": "(blank)",
        "notes": "Confirmed by accounts team.",
    },
    {
        "account": "AMB MASTER KVB-6375",
        "rule": "Credit is always Collection",
        "trigger": "CREDITS > 0.",
        "examples": "RTGS CR-PSIB0020974-DIMPAL-AMBITION COLONISERS PRIVATE LTD-...",
        "bu": "SW", "head": "Collection", "type_rera_idw": "Customer Collection", "tcp_head": "Credit- no effect",
        "notes": "No exceptions confirmed.",
    },
    {
        "account": "AMB MASTER KVB-6375",
        "rule": "Debit is always Internal",
        "trigger": "DEBITS > 0. TYPE FOR RERA IDW derived from which AMB account number appears in the description (target account's stage).",
        "examples": "KVBLH00258953931-Ambition Colonisers Private Limited-60073932667 (target = RERA account) -> Master 2 RERA",
        "bu": "SW", "head": "Internal", "type_rera_idw": "Master 2 RERA / Master to Free / Internal (by target stage)", "tcp_head": "Internal transfer",
        "notes": "Same Master 2 RERA / Master to Free vocabulary DPL already uses.",
    },
    {
        "account": "AMB RERA BOM 667",
        "rule": "Always Internal, direction picks the label",
        "trigger": "HEAD is always Internal on this account. CREDIT -> Master 2 RERA. DEBIT -> RERA 2 IDW.",
        "examples": "Credit: NEFT ...; Debit: NEFT ...",
        "bu": "SW", "head": "Internal", "type_rera_idw": "Master 2 RERA (credit) / RERA 2 IDW (debit)", "tcp_head": "Internal transfer",
        "notes": "Accounts team: 'only these rules for now — one more rule may be added in future.'",
    },
    {
        "account": "IDW KVB-6535",
        "rule": "Credit from Bank of Maharashtra (RERA account)",
        "trigger": "CREDITS > 0 and description contains 'MAH' (Bank of Maharashtra / MAHB IFSC).",
        "examples": "NEFT CR-MAHB0001461-AMBITION COLONISERS PRIVATE LIMITED-...",
        "bu": "SW", "head": "Internal", "type_rera_idw": "RERA 2 IDW", "tcp_head": "Internal transfer",
        "notes": "Every real credit row on this account matches this pattern.",
    },
    {
        "account": "IDW KVB-6535",
        "rule": "Debit internal transfer to another AMB account",
        "trigger": "DEBITS > 0 and another AMB account number appears in the description. Takes priority over the site-expense rule below.",
        "examples": "...-4114135000001050-internal (target = Free account)",
        "bu": "SW", "head": "Internal", "type_rera_idw": "Free & IDW Loan (target=Free) / Internal (other targets)", "tcp_head": "Internal transfer",
        "notes": "Master's and IDW's own account numbers appear with an extra digit (16 vs 15) when referenced by OTHER accounts' statements — both forms are matched.",
    },
    {
        "account": "IDW KVB-6535",
        "rule": "Internal fund transfer (AMB company named as counterparty)",
        "trigger": "Description names 'Ambition Colonisers' as counterparty, but the target account isn't one of the 4 confirmed AMB accounts (e.g. routed through AMB CA AXIS-280).",
        "examples": "KVBLH00259426389-Ambition Colonisers Pvt Ltd-925020010722280-for pf",
        "bu": "SW", "head": "Internal", "type_rera_idw": "Internal", "tcp_head": "Internal transfer",
        "notes": "AXIS-280's own rules aren't confirmed yet — this only catches the counterparty being AMB itself.",
    },
    {
        "account": "IDW KVB-6535",
        "rule": "Site-expense bucket (everything else)",
        "trigger": "DEBITS > 0, not an internal transfer. Head from Beneficiary Master (named identity) first, then keyword: salary/imprest/contractor/vendor/card/bank-charges. 'Card' is checked before the Beneficiary Master (transaction type, not identity).",
        "examples": "Salary-Site, Vendor -Site, Imprest, Contractor, Card, S K G Buildcon (land payment)",
        "bu": "SW", "head": "(from Beneficiary Master / keyword)", "type_rera_idw": "Dev- Infra (Land Payment for S K G Buildcon)", "tcp_head": "IDW Civil Works (Other- Land Cost for S K G Buildcon)",
        "notes": "This is a site account — accounts team: 'contractor vendor or salary these are for site... check beneficiary master.'",
    },
    {
        "account": "KVB FREE 1050",
        "rule": "S K G Buildcon credit",
        "trigger": "CREDITS > 0 and description names S K G Buildcon (either 'S K G BUILDCON' or 'S.K.G. BUILDCON' spelling).",
        "examples": "RTGS CR-UTIB0003622-S.K.G. BUILDCON PRIVATE LIMITED-...",
        "bu": "SW", "head": "Loan", "type_rera_idw": "Promoter Contribution", "tcp_head": "Credit- no effect",
        "notes": "Resolves an earlier ambiguity in the reference data (same description pattern also labeled 'SKG Buildcon' in some historical rows) — HEAD=Loan confirmed by accounts team for the credit case.",
    },
    {
        "account": "KVB FREE 1050",
        "rule": "Hoarding",
        "trigger": "Description contains 'Hoarding'.",
        "examples": "KVBLH00259704850-Surender Kumar-50100736339069-Hoarding",
        "bu": "SW", "head": "MKT/ADVER", "type_rera_idw": "HO - Advert/ Mkt", "tcp_head": "Other- Selling Expenses",
        "notes": "A later batch of 17 reference rows shows TYPE='HO - Admin' instead — accounts team confirmed this is a reference-data inconsistency; 'HO - Advert/ Mkt' is the rule.",
    },
    {
        "account": "KVB FREE 1050",
        "rule": "Internal transfer from Master account",
        "trigger": "Description contains Master's account number (either digit form).",
        "examples": "070426BB4552144A-AMBITION COLONISERS-4114135000006375-master to free",
        "bu": "SW", "head": "Internal", "type_rera_idw": "Master to Free", "tcp_head": "Internal transfer",
        "notes": "Confirmed: SKG, Hoarding, and 'master to free' are the only SW cases on this account.",
    },
    {
        "account": "KVB FREE 1050",
        "rule": "Internal fund transfer (AMB company named as counterparty)",
        "trigger": "Description names 'Ambition Colonisers' as counterparty, target isn't the Master account (e.g. pf/esi/tds routed through AXIS-280).",
        "examples": "KVBLH00259426391-Ambition Colonisers Pvt Ltd-925020010722280-for esi",
        "bu": "HO", "head": "Internal", "type_rera_idw": "Internal", "tcp_head": "Internal transfer",
        "notes": "Everything else on this account defaults to HO — this is the internal-transfer exception within that default.",
    },
    {
        "account": "KVB FREE 1050",
        "rule": "HO-default bucket (everything else)",
        "trigger": "Card/Imprest/Commission (incl. 'broker'/'incentive') keywords checked FIRST (transaction type, not identity — a Beneficiary Master name tagged Salary-HO can still receive a one-off Imprest/Commission payment). Then Beneficiary Master, then salary/vendor/electricity/rent/maintenance/professional keywords.",
        "examples": "Salary-HO, Vendor - Ho, Card, Commission (incl. broker), Imprest, Legal & Proff., Bank Charges",
        "bu": "HO", "head": "(from keyword / Beneficiary Master)", "type_rera_idw": "HO - Admin", "tcp_head": "Other- Administrative Expenses",
        "notes": "Confirmed: broker-keyword Commission rows and one Imprest exception in the reference data are still HO (not SW) — accounts team confirmed these are reference-data inconsistencies, not real rules.",
    },
]

DEFERRED: list[dict] = [
    {
        "item": "AMB CA AXIS-280",
        "issue": "Rules not yet confirmed by accounts team.",
        "action": "This account's transactions keep using the generic keyword + Beneficiary Master behavior until its rules are shared and confirmed.",
    },
    {
        "item": "Generic '-tfr' suffix (KVB FREE 1050 / IDW KVB-6535)",
        "issue": "Used across Contractor, Loan, Full & Final, Wages, and plain Internal with no reliable keyword signal.",
        "action": "Resolved via the Beneficiary Master only — per explicit accounts-team instruction, not a hardcoded keyword rule.",
    },
    {
        "item": "IMPS 'Loan Repayment' / 'Money Transfer' (KVB FREE 1050)",
        "issue": "A few Shobha Jain rows use these exact phrases for Loan-head transactions — not yet a confirmed general rule.",
        "action": "Left for Beneficiary Master / manual review.",
    },
]

# ---------------------------------------------------------------------------
# Colors (same palette as DPL's create_rules_tab.py)
# ---------------------------------------------------------------------------
COLOR_HEADER = {"red": 0.145, "green": 0.208, "blue": 0.353}
COLOR_SECTION = {"red": 0.235, "green": 0.408, "blue": 0.627}
COLOR_ALT_ROW = {"red": 0.918, "green": 0.937, "blue": 0.976}
COLOR_WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
COLOR_PENDING_BG = {"red": 1.0, "green": 0.949, "blue": 0.8}
COLOR_TEXT_WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
COLOR_TEXT_DARK = {"red": 0.1, "green": 0.1, "blue": 0.1}


def cell_format(bg: dict, bold: bool = False, size: int = 10,
                fg: dict | None = None, wrap: str = "WRAP",
                halign: str = "LEFT", valign: str = "TOP") -> dict:
    return {
        "backgroundColor": bg,
        "textFormat": {"bold": bold, "fontSize": size, "foregroundColor": fg or COLOR_TEXT_DARK},
        "wrapStrategy": wrap,
        "horizontalAlignment": halign,
        "verticalAlignment": valign,
    }


def build_rules_sheet(ws: gspread.Worksheet) -> None:
    ws.clear()

    rows: list[list[str]] = []
    formats: list[dict] = []
    merges: list[dict] = []
    NCOLS = 8

    def add_row(cells: list[str], fmt_per_cell: list[dict] | None = None, default_fmt: dict | None = None) -> None:
        rows.append(cells + [""] * (NCOLS - len(cells)))
        r = len(rows)
        if fmt_per_cell:
            for c, fmt in enumerate(fmt_per_cell, 1):
                formats.append({"row": r, "col": c, "fmt": fmt})
        elif default_fmt:
            for c in range(1, NCOLS + 1):
                formats.append({"row": r, "col": c, "fmt": default_fmt})

    def merge_row(row_1indexed: int, start_col: int = 0, end_col: int = NCOLS) -> None:
        merges.append({"sheetId": ws.id, "startRowIndex": row_1indexed - 1, "endRowIndex": row_1indexed,
                       "startColumnIndex": start_col, "endColumnIndex": end_col})

    add_row(["AMB BANK STATEMENT — AUTO-CLASSIFICATION RULES"],
            default_fmt=cell_format(COLOR_HEADER, bold=True, size=13, fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))
    merge_row(len(rows))

    add_row(["Confirmed directly by the accounts team, account by account, and cross-checked "
             "against 391 real reference transactions. Covers 4 of AMB's 5 accounts — "
             "AMB CA AXIS-280 is pending (see below)."],
            default_fmt=cell_format({"red": 0.18, "green": 0.18, "blue": 0.18},
                                    fg={"red": 0.75, "green": 0.75, "blue": 0.75}, halign="CENTER"))
    merge_row(len(rows))

    add_row([])

    cols = ["ACCOUNT", "RULE", "TRIGGER", "EXAMPLE", "BUSINESS UNIT", "HEAD", "TYPE FOR RERA IDW", "TCP HEAD"]
    add_row(cols, default_fmt=cell_format(COLOR_SECTION, bold=True, size=10, fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))

    prev_account = None
    for i, rule in enumerate(RULES):
        if rule["account"] != prev_account:
            add_row([rule["account"]], default_fmt=cell_format({"red": 0.878, "green": 0.878, "blue": 0.878}, bold=True, halign="LEFT", valign="MIDDLE"))
            merge_row(len(rows))
            prev_account = rule["account"]

        row_bg = COLOR_ALT_ROW if i % 2 == 0 else COLOR_WHITE
        add_row([
            rule["account"], rule["rule"], rule["trigger"], rule["examples"],
            rule["bu"], rule["head"], rule["type_rera_idw"], rule["tcp_head"],
        ], default_fmt=cell_format(row_bg))
        # Notes as its own row underneath, spanning the width, italic-ish via dark bg
        add_row([f"Note: {rule['notes']}"], default_fmt=cell_format({"red": 0.96, "green": 0.96, "blue": 0.96}, size=9))
        merge_row(len(rows))

    add_row([])
    add_row([])
    add_row(["DEFERRED / PENDING — NOT AUTO-CLASSIFIED"],
            default_fmt=cell_format({"red": 0.8, "green": 0.3, "blue": 0.0}, bold=True, size=11, fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))
    merge_row(len(rows))

    add_row(["Item", "Issue", "Action", "", "", "", "", ""],
            default_fmt=cell_format(COLOR_SECTION, bold=True, fg=COLOR_TEXT_WHITE))

    for p in DEFERRED:
        add_row([p["item"], p["issue"], p["action"]], default_fmt=cell_format(COLOR_PENDING_BG))

    ws.update(values=rows, range_name="A1", value_input_option="RAW")

    requests = []
    for m in merges:
        requests.append({"mergeCells": {"range": m, "mergeType": "MERGE_ALL"}})
    for item in formats:
        r, c, fmt = item["row"], item["col"], item["fmt"]
        requests.append({
            "repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": r - 1, "endRowIndex": r, "startColumnIndex": c - 1, "endColumnIndex": c},
                "cell": {"userEnteredFormat": fmt},
                "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat,userEnteredFormat.wrapStrategy,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment",
            }
        })

    col_widths = [160, 220, 340, 260, 110, 160, 180, 200]
    for ci, width in enumerate(col_widths):
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": ci, "endIndex": ci + 1},
                "properties": {"pixelSize": width}, "fields": "pixelSize",
            }
        })

    requests.append({
        "updateSheetProperties": {
            "properties": {"sheetId": ws.id, "gridProperties": {"frozenRowCount": 4}},
            "fields": "gridProperties.frozenRowCount",
        }
    })

    ws.spreadsheet.batch_update({"requests": requests})
    print(f"[OK] AMB Rules tab rebuilt — {len(rows)} rows, {len(formats)} formats, {len(merges)} merges.")


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(AMB_SHEET_ID)

    existing = next((ws for ws in spreadsheet.worksheets() if ws.title == "Rules"), None)
    if existing:
        spreadsheet.del_worksheet(existing)
        print("[INFO] Deleted old Rules tab.")

    ws = spreadsheet.add_worksheet(title="Rules", rows=200, cols=8)
    build_rules_sheet(ws)
    print("[DONE] AMB Rules tab is ready.")


if __name__ == "__main__":
    main()
