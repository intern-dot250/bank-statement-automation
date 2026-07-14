"""Creates (or replaces) a 'Rules' worksheet in the master Google Sheet
that documents the full auto-classification logic so the accounts team
can verify each rule and mark it as correct or incorrect.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from __future__ import annotations

import gspread
from upload_to_sheets import DEFAULT_CREDENTIALS, MASTER_SHEET_ID, get_gspread_client

# ---------------------------------------------------------------------------
# Rule data — matches the EXACT priority order in classify_transactions.py
# ---------------------------------------------------------------------------

RULES: list[dict] = [
    # ── Priority 1: Internal transfers between DPL own YES Bank accounts ─────
    {
        "priority": 1,
        "rule": "Internal Transfer — DPL Own YES Bank Account",
        "trigger": (
            "DESCRIPTION contains one of our own YES Bank account numbers "
            "(ends with 0264 / 0377 / 0490 / 2477 / 2314 / 2457). "
            "Spaces inside the number are ignored (PDF extraction sometimes splits them)."
        ),
        "examples": "YIB-TPT-DWARKADHIS PROJECTS PVT LTD...TFR-045563400002477",
        "bu": "This account's own BU (see Account Reference Table below)",
        "head": "Internal",
        "type_rera_idw": (
            "Master + RERA  →  Master 2 RERA\n"
            "Master + Free  →  Master to Free\n"
            "Free  + IDW    →  Free & IDW Loan\n"
            "RERA  + IDW    →  Rera 2 IDW  (direction decided by Debit/Credit — see Notes)\n"
            "Any   + AH     →  Internal"
        ),
        "tcp_head": (
            "RERA→IDW Debit  →  Internal transfer\n"
            "RERA→IDW Credit →  Rera to IDW\n"
            "All others      →  Internal transfer"
        ),
        "notes": (
            "RERA↔IDW direction rule: if DEBITS > 0 (money leaving) → TCP = 'Internal transfer'. "
            "If CREDITS > 0 (money arriving) → TCP = 'Rera to IDW'. "
            "BU comes from account setup — see BU Priority section below."
        ),
    },
    # ── Priority 2: Bank of Maharashtra (BOM) internal transfer ──────────────
    {
        "priority": 2,
        "rule": "Internal Transfer — Bank of Maharashtra (BOM)",
        "trigger": (
            "DESCRIPTION contains any MAHB IFSC code (pattern: MAHB + 7 characters, "
            "e.g. MAHB0001461) OR the text 'BANK OF MAHARASHTRA'. "
            "Does NOT apply if description starts with UPI/ NEFT CR- IMPS/ RTGS CR- "
            "(those are incoming payments from external parties who happen to bank with BOM)."
        ),
        "examples": "NEFT-MAHB0001461-DWARKADHIS..., MAHB0001347...",
        "bu": "This account's own BU",
        "head": "Internal",
        "type_rera_idw": "Internal",
        "tcp_head": "Internal transfer",
        "notes": "Any MAHB IFSC = DPL's own BOM account. No 'Dwarkadhis' keyword needed.",
    },
    # ── Priority 3: Salary ────────────────────────────────────────────────────
    {
        "priority": 3,
        "rule": "Salary — Site (accounts 0490 and 2457)",
        "trigger": (
            "DESCRIPTION contains 'salary' as a hyphen/slash segment OR as the last word "
            "of any segment (e.g. 'BHARAT SINGH SALARY' or '-SALARY-'). "
            "Account ends with 0490 OR 2457. "
            "Note: account 0490 is hardcoded as Salary Site even if admin panel stage is blank."
        ),
        "examples": "YIB-NEFT-...-BHARAT SINGH SALARY-..., -SALARY-",
        "bu": "This account's own BU (Casa Romana for 0490, Aravali Heights for 2457)",
        "head": "Salary Site",
        "type_rera_idw": "Dev- Apt",
        "tcp_head": "IDW Civil Works",
        "notes": (
            "RAM KISHAN — April entries say 'SALARY' explicitly so go through this rule. "
            "May onwards the description says 'CONTRACTOR' so he goes through Priority 6."
        ),
    },
    {
        "priority": 3,
        "rule": "Salary — HO (all other accounts)",
        "trigger": (
            "DESCRIPTION contains 'salary' as a segment or last word. "
            "Account is NOT 0490 or 2457 (i.e. any Free / Master / RERA account)."
        ),
        "examples": "VISHAL KUMAR SALARY, -SALARY-HO-",
        "bu": "HO",
        "head": "Salary HO",
        "type_rera_idw": "HO - Admin",
        "tcp_head": "Other- Administrative Expenses",
        "notes": "BU is always 'HO' (Head Office) regardless of which bank account it's paid from.",
    },
    # ── Priority 4: Statutory Dues (PF / ESI / TDS) ───────────────────────────
    {
        "priority": 4,
        "rule": "Statutory Dues — PF / ESI / TDS",
        "trigger": (
            "DESCRIPTION contains any of: PROVIDENT FUND, EPF, ESIC, ESI, PF, TDS, "
            "TAX DEDUCTED, PTAX, PROFESSIONAL TAX."
        ),
        "examples": "ARAVALI HEIGHT RESIDENT WELFARE-EPF..., TDS PAYMENT...",
        "bu": "HO",
        "head": "Statutory Dues",
        "type_rera_idw": "HO - Admin",
        "tcp_head": "Other- Administrative Expenses",
        "notes": (
            "ALWAYS HO-Admin regardless of account or project. Even site-staff PF/ESI is "
            "routed through the Free/HO account and expensed as Admin — never capitalized as IDW. "
            "Confirmed by accounts team (VJ rulebook)."
        ),
    },
    # ── Priority 5: Marketing / Advertising ───────────────────────────────────
    {
        "priority": 5,
        "rule": "Marketing / Advertising",
        "trigger": (
            "DESCRIPTION contains any of: MARKETING, ADVERTISEMENT, ADVERTISING, "
            "-MKT-, /MKT, PUBLICITY, BRANDING, HOARDING."
        ),
        "examples": "MKA DECORATOR-...-MKT-, ADVERTISING PAYMENT...",
        "bu": "HO",
        "head": "HO - Advert/Mkt",
        "type_rera_idw": "HO - Admin",
        "tcp_head": "Other-Selling Expenses",
        "notes": "Confirmed from VJ rulebook analysis. TCP Head is 'Other-Selling Expenses' (different from Admin).",
    },
    # ── Priority 6: Role keywords in description ──────────────────────────────
    {
        "priority": 6,
        "rule": "Cancellation",
        "trigger": (
            "DESCRIPTION contains a segment starting with 'CANCELLATION' "
            "(e.g. -CANCELLATION D6126- or -CANCELLATION 033-). "
            "Matched even if a reference code follows immediately (e.g. 'CANCELLATIOND2025')."
        ),
        "examples": "YIB-NEFT-...-CANCELLATION D2025-ICICI BANK",
        "bu": "This account's own BU",
        "head": "Cancellation",
        "type_rera_idw": "Cust Cancellation",
        "tcp_head": "? (Never recorded in 2 years of historical data — leave blank)",
        "notes": "TCP Head for Cancellation is genuinely unknown — accounts team always left it blank.",
    },
    {
        "priority": 6,
        "rule": "Professional — Known Firm or 'Professional' Keyword",
        "trigger": (
            "DESCRIPTION contains 'professional' as a segment OR "
            "contains a known professional firm name anywhere in the text "
            "(spaces removed for matching, handles PDF mid-word wraps).\n"
            "Known firms currently in system: NARESH K JAIN.\n"
            "To add a new firm: update KNOWN_PROFESSIONAL_FIRMS in classify_transactions.py."
        ),
        "examples": "NARESH K JAIN AND CO, -PROFESSIONAL-",
        "bu": "HO",
        "head": "Professional",
        "type_rera_idw": "HO - Admin",
        "tcp_head": "Other- Administrative Expenses",
        "notes": "Always HO regardless of which account pays. Add new CA/legal firm names to the code.",
    },
    {
        "priority": 6,
        "rule": "Vendor",
        "trigger": "DESCRIPTION contains '-VENDOR-' or '/VENDOR' as a segment (case-insensitive).",
        "examples": "-VENDOR-MATERIAL SUPPLY-, /VENDOR HDFC BANK",
        "bu": (
            "IDW / AH-IDW stage  →  This account's own BU\n"
            "Free stage          →  HO\n"
            "Other stages        →  This account's own BU"
        ),
        "head": "Vendor",
        "type_rera_idw": "IDW/AH-IDW → Dev- Apt  |  Free/HO → HO - Admin",
        "tcp_head": "IDW/AH-IDW → IDW Civil Works  |  Free/HO → Other- Administrative Expenses",
        "notes": "Free-stage Vendor uses HO-Admin (same as Professional) per accounts team reference.",
    },
    {
        "priority": 6,
        "rule": "Contractor — Keyword or Known Individual",
        "trigger": (
            "DESCRIPTION contains '-CONTRACTOR-' / '-CONTRACT-' / '/CONTRACTOR' as a segment, "
            "OR contains the full name of a known contractor anywhere in the text.\n"
            "Known contractors currently in system: RAM KISHAN, SHER SINGH.\n"
            "RAM KISHAN exception: April descriptions say 'SALARY' → goes through Priority 3 (Salary). "
            "May onwards descriptions say 'CONTRACTOR' → comes here.\n"
            "To add a new contractor: update KNOWN_CONTRACTORS in classify_transactions.py."
        ),
        "examples": "-CONTRACTOR-CIVIL WORK-, RAM KISHAN, SHER SINGH, /MUKESH KUMAR/CONTRACTOR",
        "bu": (
            "IDW / AH-IDW stage  →  This account's own BU\n"
            "Free stage          →  HO\n"
            "Other stages        →  This account's own BU"
        ),
        "head": "Contractor",
        "type_rera_idw": "IDW/AH-IDW → Dev- Apt  |  Others → ?",
        "tcp_head": "IDW/AH-IDW → IDW Civil Works  |  Others → ?",
        "notes": "IMPS/NA format uses /CONTRACTOR (slash-separated) — also detected.",
    },
    {
        "priority": 6,
        "rule": "Imprest",
        "trigger": "DESCRIPTION contains '-IMPREST-' as a segment (case-insensitive).",
        "examples": "-IMPREST-SITE EXPENSES-, RAVI VATS-IMPREST-UNION BANK",
        "bu": (
            "IDW / AH-IDW stage  →  This account's own BU\n"
            "Free stage          →  HO\n"
            "Other stages        →  This account's own BU"
        ),
        "head": "Imprest",
        "type_rera_idw": "IDW/AH-IDW → Dev- Apt  |  Others → ?",
        "tcp_head": "IDW/AH-IDW → IDW Civil Works  |  Others → ?",
        "notes": "Imprest from Free/HO account → HO-Admin (same behaviour as Free-stage Vendor/Contractor).",
    },
    # ── Priority 7: CHQ DEP / Cheque deposit (Collection) ────────────────────
    {
        "priority": 7,
        "rule": "Collection — Cheque Deposit (CHQ DEP / BY CLG)",
        "trigger": (
            "CREDITS > 0 AND DESCRIPTION contains 'CHQ DEP', 'CHEQ DEP', or 'BY CLG' "
            "(spaces removed for matching)."
        ),
        "examples": "CHQ DEP-123456, BY CLG-REF-789",
        "bu": "This account's own BU",
        "head": "Collection",
        "type_rera_idw": "Customer Collection",
        "tcp_head": "Credit- no effect",
        "notes": "Separate from Priority 8 (UPI/NEFT). Cheque deposits are detected by description keyword, not prefix.",
    },
    # ── Priority 8: Incoming customer payment (Collection) ───────────────────
    {
        "priority": 8,
        "rule": "Collection — Incoming Digital Payment",
        "trigger": (
            "CREDITS > 0 AND DESCRIPTION starts with one of: "
            "UPI/  |  NEFT CR-  |  IMPS/  |  RTGS CR-  |  NET-TPT-  |  NET-"
        ),
        "examples": "UPI/...-RAHUL SHARMA..., NEFT CR-...-CUSTOMER NAME, IMPS/NA/...",
        "bu": "This account's own BU",
        "head": "Collection",
        "type_rera_idw": "Customer Collection",
        "tcp_head": "Credit- no effect",
        "notes": "These prefixes identify money coming IN from a customer, not internal transfers.",
    },
    # ── Priority 9: Fallback ──────────────────────────────────────────────────
    {
        "priority": 9,
        "rule": "Fallback — Keyword Heuristic (heads.py)",
        "trigger": (
            "None of Priorities 1–8 matched. System runs a keyword scan of the "
            "full DESCRIPTION text against a dictionary of known terms (heads.py)."
        ),
        "examples": "Any description not matched by rules above",
        "bu": "?",
        "head": "? (or keyword match if heuristic finds one)",
        "type_rera_idw": "?",
        "tcp_head": "?",
        "notes": (
            "If HEAD = '?' after this step, the row is flagged RED for manual review. "
            "To fix permanently: add the keyword/pattern to classify_transactions.py."
        ),
    },
]

# ---------------------------------------------------------------------------
# Pending rules (confirmed needed but Head/TCP not yet decided by accounts team)
# ---------------------------------------------------------------------------

PENDING_RULES: list[dict] = [
    {
        "rule": "Bank Charges",
        "trigger": "Description contains BANK CHARGE, SERVICE CHARGE, GST ON, CHGS, CHRG etc.",
        "issue": (
            "Contradictory in source data — same account (IDW 0490), same Sub Head, "
            "two different outcomes observed: HO-Admin in some rows, IDW Other in others. "
            "Cannot reliably auto-classify without a consistent rule from accounts team."
        ),
        "action": "Accounts team to decide: always HO-Admin, or follow account stage?",
    },
    {
        "rule": "EDC / IDC (External / Internal Development Charges)",
        "trigger": "Description contains EDC, IDC, EXTERNAL DEVELOPMENT, INTERNAL DEVELOPMENT.",
        "issue": "Head and TCP not yet confirmed by accounts team.",
        "action": "Accounts team to provide Head, Type for RERA IDW, and TCP Head for EDC/IDC payments.",
    },
    {
        "rule": "RERA / Legal Fees",
        "trigger": "Description contains RERA, LEGAL FEE, ADVOCATE, SOLICITOR etc.",
        "issue": "Unclear if this maps to 'Professional' or a separate Head. Not confirmed.",
        "action": "Accounts team to confirm: is this 'Professional' (HO-Admin) or a separate Head?",
    },
]

# ---------------------------------------------------------------------------
# Account reference table
# ---------------------------------------------------------------------------

ACCOUNTS: list[dict] = [
    {"suffix": "0264", "bu": "Casa Romana",     "stage": "Free",    "salary_type": "Salary HO"},
    {"suffix": "0377", "bu": "Casa Romana",     "stage": "RERA",    "salary_type": "Salary HO"},
    {"suffix": "0490", "bu": "Casa Romana",     "stage": "IDW",     "salary_type": "Salary Site"},
    {"suffix": "2477", "bu": "Casa Romana",     "stage": "Free",    "salary_type": "Salary HO"},
    {"suffix": "2314", "bu": "Aravali Heights", "stage": "Master",  "salary_type": "Salary HO"},
    {"suffix": "2457", "bu": "Aravali Heights", "stage": "AH-IDW",  "salary_type": "Salary Site"},
]

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------

COLOR_HEADER     = {"red": 0.145, "green": 0.208, "blue": 0.353}
COLOR_SECTION    = {"red": 0.235, "green": 0.408, "blue": 0.627}
COLOR_ALT_ROW    = {"red": 0.918, "green": 0.937, "blue": 0.976}
COLOR_WHITE      = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
COLOR_PENDING_BG = {"red": 1.0,   "green": 0.949, "blue": 0.8}
COLOR_TEXT_WHITE = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
COLOR_TEXT_DARK  = {"red": 0.1,   "green": 0.1,   "blue": 0.1}
COLOR_TEXT_ORANGE= {"red": 0.8,   "green": 0.4,   "blue": 0.0}


def cell_format(bg: dict, bold: bool = False, size: int = 10,
                fg: dict | None = None, wrap: str = "WRAP",
                halign: str = "LEFT", valign: str = "TOP") -> dict:
    return {
        "backgroundColor": bg,
        "textFormat": {"bold": bold, "fontSize": size,
                       "foregroundColor": fg or COLOR_TEXT_DARK},
        "wrapStrategy": wrap,
        "horizontalAlignment": halign,
        "verticalAlignment": valign,
    }


def build_rules_sheet(ws: gspread.Worksheet) -> None:
    ws.clear()

    rows: list[list[str]] = []
    formats: list[dict] = []
    merges: list[dict] = []

    NCOLS = 10

    def add_row(cells: list[str], fmt_per_cell: list[dict] | None = None,
                default_fmt: dict | None = None) -> None:
        rows.append(cells + [""] * (NCOLS - len(cells)))
        r = len(rows)
        if fmt_per_cell:
            for c, fmt in enumerate(fmt_per_cell, 1):
                formats.append({"row": r, "col": c, "fmt": fmt})
        elif default_fmt:
            for c in range(1, NCOLS + 1):
                formats.append({"row": r, "col": c, "fmt": default_fmt})

    def merge_row(row_1indexed: int, start_col: int = 0, end_col: int = NCOLS) -> None:
        merges.append({"sheetId": ws.id,
                       "startRowIndex": row_1indexed - 1, "endRowIndex": row_1indexed,
                       "startColumnIndex": start_col, "endColumnIndex": end_col})

    # ── Title ────────────────────────────────────────────────────────────────
    cols = ["RULE #", "RULE NAME", "HOW IT IS DETECTED (TRIGGER)", "EXAMPLE DESCRIPTION",
            "BUSINESS UNIT", "HEAD", "TYPE FOR RERA IDW", "TCP HEAD",
            "NOTES / EDGE CASES", "ACCOUNTS TEAM: CORRECT? (Y/N)"]

    add_row(["DPL BANK STATEMENT — AUTO-CLASSIFICATION RULES"],
            default_fmt=cell_format(COLOR_HEADER, bold=True, size=13,
                                    fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))
    merge_row(len(rows))

    add_row(["Rules are applied in PRIORITY ORDER — first rule that matches wins. "
             "Lower priority number = checked first."],
            default_fmt=cell_format({"red": 0.18, "green": 0.18, "blue": 0.18},
                                    fg={"red": 0.75, "green": 0.75, "blue": 0.75},
                                    halign="CENTER"))
    merge_row(len(rows))

    add_row([])  # spacer

    add_row(cols,
            default_fmt=cell_format(COLOR_SECTION, bold=True, size=10,
                                    fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))

    # ── Rule rows ─────────────────────────────────────────────────────────────
    prev_priority = None
    for i, rule in enumerate(RULES):
        if rule["priority"] != prev_priority:
            sep = {
                1: "PRIORITY 1 — INTERNAL TRANSFERS (DPL Own YES Bank Accounts)",
                2: "PRIORITY 2 — BANK OF MAHARASHTRA (BOM) TRANSFERS",
                3: "PRIORITY 3 — SALARY PAYMENTS",
                4: "PRIORITY 4 — STATUTORY DUES  (PF / ESI / TDS)",
                5: "PRIORITY 5 — MARKETING / ADVERTISING",
                6: "PRIORITY 6 — ROLE KEYWORD RULES  (Cancellation / Professional / Vendor / Contractor / Imprest)",
                7: "PRIORITY 7 — CHEQUE DEPOSIT COLLECTION  (CHQ DEP / BY CLG)",
                8: "PRIORITY 8 — INCOMING DIGITAL PAYMENT COLLECTION  (UPI / NEFT / IMPS / RTGS)",
                9: "PRIORITY 9 — FALLBACK  (No Rule Matched)",
            }.get(rule["priority"], f"Priority {rule['priority']}")
            add_row([sep],
                    default_fmt=cell_format({"red": 0.878, "green": 0.878, "blue": 0.878},
                                            bold=True, halign="LEFT", valign="MIDDLE"))
            merge_row(len(rows))
            prev_priority = rule["priority"]

        row_bg = COLOR_ALT_ROW if i % 2 == 0 else COLOR_WHITE
        add_row([
            f"Priority {rule['priority']}",
            rule["rule"],
            rule["trigger"],
            rule["examples"],
            rule["bu"],
            rule["head"],
            rule["type_rera_idw"],
            rule["tcp_head"],
            rule["notes"],
            "",
        ], default_fmt=cell_format(row_bg))

    # ── BU Priority explanation ───────────────────────────────────────────────
    add_row([])
    add_row([])
    add_row(["BUSINESS UNIT (BU) SOURCE PRIORITY"],
            default_fmt=cell_format(COLOR_HEADER, bold=True, size=11,
                                    fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))
    merge_row(len(rows))

    add_row(["Step", "Source", "How it works", "When it applies", "", "", "", "", "", ""],
            default_fmt=cell_format(COLOR_SECTION, bold=True, fg=COLOR_TEXT_WHITE, halign="CENTER"))

    bu_steps = [
        ("1 (first)", "Admin Panel (Database)",
         "BU set by admin in the web app for this account.",
         "Always checked first. If set here, this value is used."),
        ("2 (fallback)", "Hardcoded override by account suffix",
         "0264 → Casa Romana, 0377 → Casa Romana, 0490 → Casa Romana, "
         "2314 → Aravali Heights, 2457 → Aravali Heights.",
         "Used only when admin panel BU is blank."),
        ("3 (last resort)", "Unknown",
         "BU is set to '?' and flagged RED for manual review.",
         "Used when neither admin panel nor override has a value."),
    ]
    for k, (step, source, how, when) in enumerate(bu_steps):
        row_bg = COLOR_ALT_ROW if k % 2 == 0 else COLOR_WHITE
        add_row([step, source, how, when], default_fmt=cell_format(row_bg))

    # ── Account reference table ───────────────────────────────────────────────
    add_row([])
    add_row([])
    add_row(["ACCOUNT REFERENCE TABLE"],
            default_fmt=cell_format(COLOR_HEADER, bold=True, size=11,
                                    fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))
    merge_row(len(rows))

    add_row(["Account (last 4)", "Business Unit", "Stage",
             "Salary Type → HEAD", "Stage determines Type/TCP for Vendor / Contractor / Imprest rows",
             "", "", "", "", ""],
            default_fmt=cell_format(COLOR_SECTION, bold=True, fg=COLOR_TEXT_WHITE, halign="CENTER"))

    for j, acc in enumerate(ACCOUNTS):
        row_bg = COLOR_ALT_ROW if j % 2 == 0 else COLOR_WHITE
        add_row([
            f"...{acc['suffix']}",
            acc["bu"],
            acc["stage"],
            acc["salary_type"],
            "IDW/AH-IDW → Dev- Apt / IDW Civil Works  |  Free/Master/RERA → HO-Admin / Other-Admin Expenses",
        ], default_fmt=cell_format(row_bg))

    # ── Known names lists ─────────────────────────────────────────────────────
    add_row([])
    add_row([])
    add_row(["KNOWN NAMES — ADD NEW ENTRIES HERE (notify developer to update code)"],
            default_fmt=cell_format(COLOR_HEADER, bold=True, size=11,
                                    fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))
    merge_row(len(rows))

    add_row(["Type", "Names currently in system", "Head assigned", "How to add more", "", "", "", "", "", ""],
            default_fmt=cell_format(COLOR_SECTION, bold=True, fg=COLOR_TEXT_WHITE))

    known = [
        ("Professional Firms", "NARESH K JAIN", "Professional → HO-Admin",
         "Tell developer firm name exactly as it appears in DESCRIPTION. Add to KNOWN_PROFESSIONAL_FIRMS."),
        ("Contractors (Individuals)", "RAM KISHAN, SHER SINGH", "Contractor → IDW/HO based on account stage",
         "Tell developer name exactly as in DESCRIPTION. Add to KNOWN_CONTRACTORS. "
         "Note: RAM KISHAN April = Salary (description says SALARY); May+ = Contractor."),
    ]
    for k, (typ, names, head, how) in enumerate(known):
        row_bg = COLOR_ALT_ROW if k % 2 == 0 else COLOR_WHITE
        add_row([typ, names, head, how], default_fmt=cell_format(row_bg))

    # ── Pending rules ─────────────────────────────────────────────────────────
    add_row([])
    add_row([])
    add_row(["PENDING RULES — ACCOUNTS TEAM ACTION REQUIRED"],
            default_fmt=cell_format({"red": 0.8, "green": 0.3, "blue": 0.0},
                                    bold=True, size=11,
                                    fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))
    merge_row(len(rows))

    add_row(["Rule Name", "How Detected", "Issue / Why Not Implemented",
             "Action Needed from Accounts Team", "", "", "", "", "", ""],
            default_fmt=cell_format(COLOR_SECTION, bold=True, fg=COLOR_TEXT_WHITE))

    for k, p in enumerate(PENDING_RULES):
        row_bg = COLOR_PENDING_BG
        add_row([p["rule"], p["trigger"], p["issue"], p["action"]],
                default_fmt=cell_format(row_bg))

    # ── Reason for ? guide ───────────────────────────────────────────────────
    add_row([])
    add_row([])
    add_row(["REASON FOR ? — QUICK REFERENCE"],
            default_fmt=cell_format(COLOR_HEADER, bold=True, size=11,
                                    fg=COLOR_TEXT_WHITE, halign="CENTER", valign="MIDDLE"))
    merge_row(len(rows))

    add_row(["REASON TEXT IN SHEET", "WHAT IT MEANS", "HOW TO FIX", "", "", "", "", "", "", ""],
            default_fmt=cell_format(COLOR_SECTION, bold=True, fg=COLOR_TEXT_WHITE))

    reasons = [
        ("description format not recognized by any existing rule",
         "None of the 9 priorities matched. Description pattern never seen before.",
         "Classify manually. Tell developer the description so a new rule can be added."),
        ("this account has no Business Unit configured",
         "BU is blank in admin panel and no hardcoded override exists for this account.",
         "Go to admin panel → find account → set Business Unit and Account Stage."),
        ("RERA<->IDW transfer — direction (debit/credit) could not be determined",
         "Both DEBITS and CREDITS are 0 for this row — cannot tell which direction.",
         "Manually set TCP Head to 'Internal transfer' or 'Rera to IDW'."),
        ("not recorded in 2 years of historical data for Cancellation transactions",
         "TCP Head for Cancellation was always blank in reference data. This is expected.",
         "Leave as-is, or fill manually if the accounts team has decided a standard label."),
        ("Head: description format not recognized by any existing rule",
         "Both the rules engine AND the keyword heuristic (heads.py) failed to assign a Head.",
         "Classify HEAD manually. If this description recurs, notify developer."),
        ("no historical data for ... payments from this account's stage",
         "Rule matched a role keyword (Vendor/Contractor/Imprest) but no Type/TCP data exists for this account's stage.",
         "Manually fill Type for RERA IDW and TCP Head. Report to developer."),
    ]
    for k, (reason, meaning, fix) in enumerate(reasons):
        row_bg = COLOR_ALT_ROW if k % 2 == 0 else COLOR_WHITE
        add_row([reason, meaning, fix], default_fmt=cell_format(row_bg))

    # ── Write data ────────────────────────────────────────────────────────────
    ws.update(values=rows, range_name="A1", value_input_option="RAW")

    # ── Batch format ──────────────────────────────────────────────────────────
    requests = []

    for m in merges:
        requests.append({"mergeCells": {"range": m, "mergeType": "MERGE_ALL"}})

    for item in formats:
        r, c, fmt = item["row"], item["col"], item["fmt"]
        requests.append({
            "repeatCell": {
                "range": {"sheetId": ws.id,
                          "startRowIndex": r - 1, "endRowIndex": r,
                          "startColumnIndex": c - 1, "endColumnIndex": c},
                "cell": {"userEnteredFormat": fmt},
                "fields": ("userEnteredFormat.backgroundColor,"
                           "userEnteredFormat.textFormat,"
                           "userEnteredFormat.wrapStrategy,"
                           "userEnteredFormat.horizontalAlignment,"
                           "userEnteredFormat.verticalAlignment"),
            }
        })

    col_widths = [90, 210, 370, 220, 180, 130, 190, 210, 260, 170]
    for ci, width in enumerate(col_widths):
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                          "startIndex": ci, "endIndex": ci + 1},
                "properties": {"pixelSize": width},
                "fields": "pixelSize",
            }
        })

    requests.append({
        "updateSheetProperties": {
            "properties": {"sheetId": ws.id, "gridProperties": {"frozenRowCount": 4}},
            "fields": "gridProperties.frozenRowCount",
        }
    })

    ws.spreadsheet.batch_update({"requests": requests})
    print(f"[OK] Rules tab rebuilt — {len(rows)} rows, {len(formats)} formats, {len(merges)} merges.")


def main() -> None:
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    existing = next((ws for ws in spreadsheet.worksheets() if ws.title == "Rules"), None)
    if existing:
        spreadsheet.del_worksheet(existing)
        print("[INFO] Deleted old Rules tab.")

    ws = spreadsheet.add_worksheet(title="Rules", rows=250, cols=10)
    build_rules_sheet(ws)
    print("[DONE] Rules tab is ready.")


if __name__ == "__main__":
    main()
