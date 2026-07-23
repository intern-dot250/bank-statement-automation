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


def get_credential_password(credential_id: int, fallback_path: Path) -> str | None:
    """Return the password for a single credential by id, or None if not found.

    Used by the admin password-reveal endpoint, which deliberately does
    NOT include the password when listing accounts (to keep it out of the
    HTML DOM). Also supports the JSON-file fallback so it works in local dev.
    """
    conn = _connect_or_none()
    if conn is not None:
        try:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT password FROM account_credentials WHERE id = %s",
                    (credential_id,),
                )
                row = cur.fetchone()
                return row[0] if row else None
        except Exception as exc:
            logger.warning("Could not read password from database: %s", exc)
            return None
        finally:
            conn.close()

    # No DB: file-fallback rows have id=None, so reveal isn't possible.
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


def update_credential(
    credential_id: int,
    bank_name: str,
    account_number: str,
    password: str | None = None,
    business_unit: str | None = None,
    company: str | None = None,
    account_stage: str | None = None,
    update_account_stage: bool = False,
) -> None:
    """Update all editable fields of an existing account credential in one
    write. DB-only, see add_credential().

    When *password* is None the existing password is left unchanged (used
    by the edit form, which deliberately does not send the password back
    unless the user explicitly changed it).

    account_stage is only written when update_account_stage=True — this
    field drives live classification behavior for existing accounts
    (see classify_transactions.py's stage-pair Type resolution), so a form
    submission that doesn't include an Account Type field at all must never
    silently clear or overwrite it."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot update accounts.")

    try:
        with conn.cursor() as cur:
            set_clauses = ["bank_name = %s", "account_number = %s", "business_unit = %s", "company = %s"]
            params: list[Any] = [bank_name, account_number, business_unit, company]
            if password is not None:
                set_clauses.append("password = %s")
                params.append(password)
            if update_account_stage:
                set_clauses.append("account_stage = %s")
                params.append(account_stage)
            params.append(credential_id)
            cur.execute(
                f"UPDATE account_credentials SET {', '.join(set_clauses)} WHERE id = %s",
                params,
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
