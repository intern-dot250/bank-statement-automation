"""Storage for per-account Google Sheet links, shown on the Admin ->
Account Passwords page.

Backend selection mirrors company_sheets_store.py:
  - If DATABASE_URL is set: reads/writes the account_sheet_links table
    in Postgres (Supabase). The table itself is created directly in
    Supabase (not auto-created here).
  - Otherwise: this is a DB-only feature, same as company_sheets_store.py
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
        logger.warning("Database unavailable for account_sheet_links: %s", exc)
        return None


def list_account_sheet_links() -> list[dict[str, Any]]:
    """Return all account sheet links as dicts with id, account_number,
    sheet_url, added_at (oldest first)."""
    conn = _connect_or_none()
    if conn is None:
        return []
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, account_number, sheet_url, added_at "
                "FROM account_sheet_links ORDER BY id ASC"
            )
            cols = ["id", "account_number", "sheet_url", "added_at"]
            return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception as exc:
        logger.warning("Could not read account_sheet_links from database: %s", exc)
        return []
    finally:
        conn.close()


def set_account_sheet_link(account_number: str, sheet_url: str) -> None:
    """Insert or update the sheet link for this account number (one link
    per account). Requires DATABASE_URL."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot set an account sheet link.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM account_sheet_links WHERE account_number = %s",
                (account_number,),
            )
            existing = cur.fetchone()
            if existing:
                cur.execute(
                    "UPDATE account_sheet_links SET sheet_url = %s WHERE account_number = %s",
                    (sheet_url, account_number),
                )
            else:
                cur.execute(
                    "INSERT INTO account_sheet_links (account_number, sheet_url) VALUES (%s, %s)",
                    (account_number, sheet_url),
                )
        conn.commit()
    finally:
        conn.close()


def delete_account_sheet_link(account_number: str) -> None:
    """Delete the sheet link for this account number, if any. DB-only,
    see set_account_sheet_link()."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot delete an account sheet link.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "DELETE FROM account_sheet_links WHERE account_number = %s",
                (account_number,),
            )
        conn.commit()
    finally:
        conn.close()
