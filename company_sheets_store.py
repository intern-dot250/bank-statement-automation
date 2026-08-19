"""Storage for per-Company Google Sheet links, shown on the Admin ->
Account Passwords page.

Backend selection mirrors credentials_store.py / gmail_accounts_store.py:
  - If DATABASE_URL is set: reads/writes the company_sheets table in
    Postgres (Supabase). The table itself is created directly in
    Supabase (not auto-created here).
  - Otherwise: this is a DB-only feature, same as gmail_accounts_store.py
    — without DATABASE_URL, the list is simply empty.
"""

from __future__ import annotations

import logging
import os
from typing import Any

logger = logging.getLogger(__name__)

DATABASE_URL_ENV_VAR = "DATABASE_URL"


def _get_connection():
    """Return a psycopg2 connection to DATABASE_URL, or None if unset."""
    database_url = os.environ.get(DATABASE_URL_ENV_VAR)
    if not database_url:
        return None

    import psycopg2  # imported lazily; optional dependency for local use

    return psycopg2.connect(database_url)


def _connect_or_none():
    try:
        return _get_connection()
    except Exception as exc:
        logger.warning("Database unavailable for company_sheets: %s", exc)
        return None


def list_company_sheets() -> list[dict[str, Any]]:
    """Return all company sheet links as dicts with id, company,
    sheet_url, financial_year, added_at, is_active (oldest first). Every
    company can now have more than one row (one per Financial Year it's
    had a sheet created for, see create_company_sheet_for_fy()) — at
    most one of them is_active per company at any time."""
    conn = _connect_or_none()
    if conn is None:
        return []
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, company, sheet_url, financial_year, added_at, is_active "
                "FROM company_sheets ORDER BY id ASC"
            )
            cols = ["id", "company", "sheet_url", "financial_year", "added_at", "is_active"]
            return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception as exc:
        logger.warning("Could not read company_sheets from database: %s", exc)
        return []
    finally:
        conn.close()


def get_company_sheet_by_fy(company: str, financial_year: str) -> dict[str, Any] | None:
    """Return the row for this exact (company, financial_year) pair, or
    None if no sheet has been created for that combination yet. Used to
    prevent creating a duplicate spreadsheet for an FY that already has
    one."""
    return next(
        (
            row for row in list_company_sheets()
            if row.get("company") == company and row.get("financial_year") == financial_year
        ),
        None,
    )


def get_active_company_sheet(company: str) -> dict[str, Any] | None:
    """Return the row currently flagged active for this company (the
    sheet that should receive new data when no specific Financial Year
    is known), or None if the company has no rows at all."""
    return next(
        (
            row for row in list_company_sheets()
            if row.get("company") == company and row.get("is_active")
        ),
        None,
    )


def create_company_sheet_for_fy(company: str, financial_year: str, sheet_url: str) -> None:
    """Insert a new (company, financial_year) sheet link and make it the
    active one for that company, deactivating every other row for the
    same company — all in one transaction, so there's never a moment
    with zero or two active rows for a company. Requires DATABASE_URL.
    Caller (web_app.py's create_fy_sheet route) is responsible for
    checking get_company_sheet_by_fy() first to avoid duplicates, and for
    having already verified the spreadsheet was created/shared
    successfully — this only ever records an already-working sheet."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot save a company sheet link.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE company_sheets SET is_active = FALSE WHERE company = %s",
                (company,),
            )
            cur.execute(
                "INSERT INTO company_sheets (company, sheet_url, financial_year, is_active) "
                "VALUES (%s, %s, %s, TRUE)",
                (company, sheet_url, financial_year),
            )
        conn.commit()
    finally:
        conn.close()


def add_company_sheet(company: str, sheet_url: str, financial_year: str | None = None) -> None:
    """Insert a new company -> sheet URL mapping and make it the active
    one for that company (deactivating any other row for the same
    company, same invariant as create_company_sheet_for_fy() — this is
    the manual equivalent, e.g. for linking a company's existing
    externally-created spreadsheet rather than one this app created).
    Requires DATABASE_URL."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot add a company sheet link.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE company_sheets SET is_active = FALSE WHERE company = %s",
                (company,),
            )
            cur.execute(
                "INSERT INTO company_sheets (company, sheet_url, financial_year, is_active) "
                "VALUES (%s, %s, %s, TRUE)",
                (company, sheet_url, financial_year),
            )
        conn.commit()
    finally:
        conn.close()


def update_company_sheet(sheet_id: int, company: str, sheet_url: str, financial_year: str | None = None) -> None:
    """Update an existing company sheet link's company name/URL/Financial
    Year. DB-only, see add_company_sheet()."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot update a company sheet link.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE company_sheets SET company = %s, sheet_url = %s, financial_year = %s WHERE id = %s",
                (company, sheet_url, financial_year, sheet_id),
            )
        conn.commit()
    finally:
        conn.close()


def delete_company_sheet(sheet_id: int) -> None:
    """Delete a company sheet link by id. DB-only, see add_company_sheet()."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot delete a company sheet link.")

    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM company_sheets WHERE id = %s", (sheet_id,))
        conn.commit()
    finally:
        conn.close()
