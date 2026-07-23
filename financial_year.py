"""Financial Year (FY) validation for bank statement processing.

A Financial Year runs 01-Apr to 31-Mar and is labeled "YYYY-YY" (e.g.
"2025-26" = 01-Apr-2025 to 31-Mar-2026). This module is intentionally
free of Flask/Google-Sheets dependencies so it can be reused from
run_pipeline.py (the single choke point shared by the manual-upload and
email-automation paths) as well as tested in isolation.
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd

_FY_LABEL_RE = re.compile(r"^(\d{4})-(\d{2})$")


def parse_fy_label(label: str) -> tuple[date, date]:
    """Parse "YYYY-YY" into (start, end) dates: 01-Apr-<YYYY> to
    31-Mar-<YYYY+1>. Raises ValueError if the label is malformed or the
    second half isn't (first year + 1) mod 100 (e.g. "2025-27" is
    rejected — that's not a valid single financial year)."""
    match = _FY_LABEL_RE.match((label or "").strip())
    if not match:
        raise ValueError(
            f"Invalid Financial Year format: {label!r}. Expected 'YYYY-YY', e.g. '2025-26'."
        )

    start_year = int(match.group(1))
    end_suffix = int(match.group(2))
    expected_end_suffix = (start_year + 1) % 100
    if end_suffix != expected_end_suffix:
        raise ValueError(
            f"Invalid Financial Year {label!r}: {start_year}-{end_suffix:02d} does not "
            f"describe a single financial year (expected {start_year}-{expected_end_suffix:02d})."
        )

    return date(start_year, 4, 1), date(start_year + 1, 3, 31)


def generate_fy_options(years_back: int = 2, years_forward: int = 3, today: date | None = None) -> list[str]:
    """Return a list of "YYYY-YY" labels for a dropdown, centered on
    whichever FY contains `today` (defaults to date.today())."""
    today = today or date.today()
    current_start_year = today.year if today.month >= 4 else today.year - 1
    return [
        f"{start_year}-{(start_year + 1) % 100:02d}"
        for start_year in range(current_start_year - years_back, current_start_year + years_forward + 1)
    ]


def resolve_financial_year(account_number: str, records_path: Path) -> str | None:
    """Return the FY label that governs this account, or None if none is
    configured (in which case validation is skipped, not failed).

    The account's own Financial Year (set on Admin -> Account Passwords)
    takes precedence; if unset, falls back to its company's Financial
    Year (set on Admin -> Company Sheet Links), matched via the
    account's `company` field.
    """
    import credentials_store
    import company_sheets_store

    accounts = credentials_store.list_credentials(records_path)
    account = next(
        (acc for acc in accounts if acc.get("account_number") == account_number),
        None,
    )
    if account is None:
        return None

    fy_label = account.get("financial_year")
    if fy_label:
        return fy_label

    company = account.get("company")
    if not company:
        return None

    company_sheets = company_sheets_store.list_company_sheets()
    company_sheet = next(
        (row for row in company_sheets if row.get("company") == company),
        None,
    )
    if company_sheet is None:
        return None

    return company_sheet.get("financial_year")


def compute_statement_period(excel_path: Path) -> tuple[date, date] | None:
    """Read the extracted statement's "Transaction Date" column and
    return (earliest, latest) as dates. Returns None if the column is
    missing or no row has a parseable date."""
    df = pd.read_excel(str(excel_path))
    if "Transaction Date" not in df.columns:
        return None

    parsed = pd.to_datetime(df["Transaction Date"], dayfirst=True, errors="coerce").dropna()
    if parsed.empty:
        return None

    return parsed.min().date(), parsed.max().date()


def validate_statement_period(fy_label: str, period: tuple[date, date]) -> tuple[bool, str]:
    """Compare a statement's (earliest, latest) transaction dates against
    the configured Financial Year's range. Returns (True, "") if fully
    inside; otherwise (False, <human-readable message>)."""
    fy_start, fy_end = parse_fy_label(fy_label)
    earliest, latest = period

    if fy_start <= earliest and latest <= fy_end:
        return True, ""

    message = (
        f"Statement dates ({earliest:%d-%b-%Y} to {latest:%d-%b-%Y}) fall outside the "
        f"configured Financial Year {fy_label} ({fy_start:%d-%b-%Y} to {fy_end:%d-%b-%Y}). "
        f"Update the account's Financial Year before processing this statement."
    )
    return False, message
