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
    sheet_url, added_at (oldest first)."""
    conn = _connect_or_none()
    if conn is None:
        return []
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, company, sheet_url, added_at "
                "FROM company_sheets ORDER BY id ASC"
            )
            cols = ["id", "company", "sheet_url", "added_at"]
            return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception as exc:
        logger.warning("Could not read company_sheets from database: %s", exc)
        return []
    finally:
        conn.close()


def add_company_sheet(company: str, sheet_url: str) -> None:
    """Insert a new company -> sheet URL mapping. Requires DATABASE_URL."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot add a company sheet link.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO company_sheets (company, sheet_url) VALUES (%s, %s)",
                (company, sheet_url),
            )
        conn.commit()
    finally:
        conn.close()


def update_company_sheet(sheet_id: int, company: str, sheet_url: str) -> None:
    """Update an existing company sheet link's company name/URL. DB-only, see add_company_sheet()."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot update a company sheet link.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE company_sheets SET company = %s, sheet_url = %s WHERE id = %s",
                (company, sheet_url, sheet_id),
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
