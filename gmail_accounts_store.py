"""Storage for connected Gmail accounts (email -> OAuth token), and which
one is currently active for "Check Bank Emails".

Backend selection mirrors credentials_store.py:
  - If DATABASE_URL is set: reads/writes the gmail_accounts table in
    Postgres (Supabase). The table itself is created directly in Supabase
    (not auto-created here).
  - Otherwise: there is no local-file fallback for this store (unlike
    credentials_store.py) — connecting a Gmail account is a DB-only
    operation, since it needs to be visible across serverless instances
    immediately. Without DATABASE_URL, authenticate_gmail() simply falls
    back to its pre-existing token.json/GOOGLE_TOKEN_JSON behavior.
"""

from __future__ import annotations

import logging
import os
from typing import Any, Optional

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
        logger.warning("Database unavailable for gmail_accounts: %s", exc)
        return None


def list_accounts() -> list[dict[str, Any]]:
    """Return all connected Gmail accounts as dicts with id, email,
    is_active, connected_at (oldest first). Never includes token_json —
    keeps the raw token out of anything rendered to a template."""
    conn = _connect_or_none()
    if conn is None:
        return []
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, email, is_active, connected_at "
                "FROM gmail_accounts ORDER BY id ASC"
            )
            cols = ["id", "email", "is_active", "connected_at"]
            return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception as exc:
        logger.warning("Could not read gmail_accounts from database: %s", exc)
        return []
    finally:
        conn.close()


def add_or_update_account(email: str, token_json: str) -> None:
    """Insert a newly-connected Gmail account, or refresh its stored token
    if that email was already connected before (upsert by email). New
    accounts are inserted as inactive — the admin explicitly activates one
    via set_active_account()."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot connect a Gmail account.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO gmail_accounts (email, token_json, is_active, connected_at)
                VALUES (%s, %s, FALSE, now())
                ON CONFLICT (email) DO UPDATE SET
                    token_json = EXCLUDED.token_json,
                    connected_at = now()
                """,
                (email, token_json),
            )
        conn.commit()
    finally:
        conn.close()


def set_active_account(account_id: int) -> None:
    """Mark this account active, all others inactive — single-active-
    account model, matching "which inbox Check Bank Emails reads from"
    being one choice at a time."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot activate a Gmail account.")

    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE gmail_accounts SET is_active = FALSE")
            cur.execute("UPDATE gmail_accounts SET is_active = TRUE WHERE id = %s", (account_id,))
        conn.commit()
    finally:
        conn.close()


def get_active_token() -> Optional[str]:
    """Return the active account's token_json, or None if none is active
    (or DATABASE_URL isn't configured) — callers should fall back to their
    own pre-existing token resolution in that case."""
    conn = _connect_or_none()
    if conn is None:
        return None
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT token_json FROM gmail_accounts WHERE is_active = TRUE LIMIT 1")
            row = cur.fetchone()
        return row[0] if row else None
    except Exception as exc:
        logger.warning("Could not read active Gmail account token: %s", exc)
        return None
    finally:
        conn.close()


def update_token(account_id: int, token_json: str) -> None:
    """Persist a silently-refreshed token back to the DB for this account
    (called from authenticate_gmail() after creds.refresh()) — critical on
    serverless, where the local token.json file write is ephemeral."""
    conn = _connect_or_none()
    if conn is None:
        return
    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE gmail_accounts SET token_json = %s WHERE id = %s",
                (token_json, account_id),
            )
        conn.commit()
    except Exception as exc:
        logger.warning("Could not persist refreshed Gmail token for account %s: %s", account_id, exc)
    finally:
        conn.close()


def get_active_account_id() -> Optional[int]:
    """Return the active account's id, or None if none is active."""
    conn = _connect_or_none()
    if conn is None:
        return None
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM gmail_accounts WHERE is_active = TRUE LIMIT 1")
            row = cur.fetchone()
        return row[0] if row else None
    except Exception as exc:
        logger.warning("Could not read active Gmail account id: %s", exc)
        return None
    finally:
        conn.close()


def delete_account(account_id: int) -> None:
    """Delete a connected Gmail account by id. DB-only, see add_or_update_account()."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot delete a Gmail account.")

    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM gmail_accounts WHERE id = %s", (account_id,))
        conn.commit()
    finally:
        conn.close()
