"""Storage for bank account number -> PDF-unlock password mappings.

Backend selection mirrors history_store.py:
  - If DATABASE_URL is set: reads/writes the account_credentials table
    in Postgres (Supabase). The table itself is created directly in
    Supabase (not auto-created here), so accounts can be added either
    through the admin page or straight in the Supabase Table Editor.
  - Otherwise: falls back to the accounts list in records.json, matching
    this project's existing local/non-serverless behavior.
"""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
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
        logger.warning("Database unavailable, falling back to local file: %s", exc)
        return None


def list_credentials(fallback_path: Path) -> list[dict[str, Any]]:
    """Return all accounts as dicts with id, bank_name, account_number,
    password, business_unit, account_stage, company (oldest first). "id" is
    None for file-fallback entries, and business_unit/account_stage/company
    are None when not set (e.g. file-fallback accounts don't have these
    fields)."""
    conn = _connect_or_none()
    if conn is not None:
        try:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT id, bank_name, account_number, password, "
                    "business_unit, account_stage, company "
                    "FROM account_credentials ORDER BY id ASC"
                )
                cols = ["id", "bank_name", "account_number", "password", "business_unit", "account_stage", "company"]
                return [dict(zip(cols, row)) for row in cur.fetchall()]
        except Exception as exc:
            logger.warning("Could not read account_credentials from database: %s", exc)
            return []
        finally:
            conn.close()

    if not fallback_path.exists():
        return []
    try:
        with open(fallback_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return [
            {
                "id": None,
                "bank_name": acc.get("bank_name", "Unknown"),
                "account_number": acc.get("account_number"),
                "password": acc.get("password"),
                "business_unit": None,
                "account_stage": None,
                "company": None,
            }
            for acc in data.get("accounts", [])
        ]
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning("Could not load accounts from local file: %s", exc)
        return []


def add_credential(
    bank_name: str,
    account_number: str,
    password: str,
    business_unit: str | None = None,
    account_stage: str | None = None,
    company: str | None = None,
) -> None:
    """Insert a new account credential. Requires DATABASE_URL — this is
    a DB-only operation, since the admin page needs immediate, shared
    visibility for the whole team (a local JSON write wouldn't be)."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot add accounts.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO account_credentials "
                "(bank_name, account_number, password, business_unit, account_stage, company) "
                "VALUES (%s, %s, %s, %s, %s, %s)",
                (bank_name, account_number, password, business_unit, account_stage, company),
            )
        conn.commit()
    finally:
        conn.close()


def update_business_fields(credential_id: int, business_unit: str, account_stage: str) -> None:
    """Update just the business_unit/account_stage fields for an existing
    account credential. DB-only, see add_credential()."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot update accounts.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE account_credentials SET business_unit = %s, account_stage = %s WHERE id = %s",
                (business_unit, account_stage, credential_id),
            )
        conn.commit()
    finally:
        conn.close()


def delete_credential(credential_id: int) -> None:
    """Delete an account credential by id. DB-only, see add_credential()."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot delete accounts.")

    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM account_credentials WHERE id = %s", (credential_id,))
        conn.commit()
    finally:
        conn.close()
