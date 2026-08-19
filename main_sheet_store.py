"""Storage for per-company Accounts Team Main Sheet sync destinations,
shown on the Admin -> Account Passwords page under "Main Sheet Sync
Settings".

Backend selection mirrors company_sheets_store.py / credentials_store.py:
  - If DATABASE_URL is set: reads/writes the main_sheet_links table in
    Postgres (Supabase). The table itself is created directly in
    Supabase (not auto-created here), matching this project's existing
    convention for admin-managed config tables.
  - Otherwise: this is a DB-only feature, same as company_sheets_store.py
    — without DATABASE_URL, the list is simply empty and
    sync_to_main_sheet.get_main_sheet_id_for_company() falls back to
    config/main_sheets.json.

Unlike company_sheets_store.py (which stores a raw sheet_url and lets
callers extract the ID each time), rows here are only ever written after
sync_to_main_sheet.py has already verified the service account can open
the sheet (see web_app.py's admin_main_sheet_save route) — so a saved
row is always known-good at save time.
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
        logger.warning("Database unavailable for main_sheet_links: %s", exc)
        return None


def list_main_sheet_links() -> list[dict[str, Any]]:
    """Return all company -> Main Sheet destination mappings as dicts
    with id, company, sheet_url, updated_at (oldest first). Empty list
    if DATABASE_URL is unset or the table can't be read — callers should
    treat that as "nothing configured yet" and fall back accordingly,
    never raise."""
    conn = _connect_or_none()
    if conn is None:
        return []
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, company, sheet_url, updated_at "
                "FROM main_sheet_links ORDER BY id ASC"
            )
            cols = ["id", "company", "sheet_url", "updated_at"]
            return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception as exc:
        logger.warning("Could not read main_sheet_links from database: %s", exc)
        return []
    finally:
        conn.close()


def get_main_sheet_link(company: str) -> dict[str, Any] | None:
    """Return the single row for *company*, or None if not configured."""
    row = next(
        (r for r in list_main_sheet_links() if r.get("company") == company),
        None,
    )
    return row


def upsert_main_sheet_link(company: str, sheet_url: str) -> None:
    """Insert or update the Main Sheet destination for *company*. One row
    per company — a second save for the same company replaces the URL
    rather than adding a duplicate. Requires DATABASE_URL. Only ever
    called after the caller has already verified access to the sheet
    (see web_app.py), so this never silently stores a broken link."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot save a Main Sheet link.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO main_sheet_links (company, sheet_url, updated_at) "
                "VALUES (%s, %s, NOW()) "
                "ON CONFLICT (company) DO UPDATE "
                "SET sheet_url = EXCLUDED.sheet_url, updated_at = NOW()",
                (company, sheet_url),
            )
        conn.commit()
    finally:
        conn.close()
