"""Storage for the admin-configured automatic daily email-processing
schedule, shown on the Admin -> Account Passwords page under
"Automatic Processing Schedule".

Backend selection mirrors main_sheet_store.py / company_sheets_store.py:
  - If DATABASE_URL is set: reads/writes the single-row scheduler_settings
    table in Postgres (Supabase). The table itself is created directly in
    Supabase (not auto-created here), matching this project's existing
    convention for admin-managed config tables.
  - Otherwise: this is a DB-only feature — without DATABASE_URL, settings
    are simply absent and the automatic schedule stays disabled (the
    cron-ping endpoint no-ops), while "Check Bank Emails" continues to
    work manually exactly as it always has.

There is intentionally only ever one row (id fixed at 1) — a single
global schedule, not per-company/per-account, matching the manager's
spec ("Automatic Processing Time: 11:00 AM", one setting for the whole
system).

Time zone: processing_time is interpreted in IST (Asia/Kolkata) — this
project's existing timestamp-display convention (see history_store.py's
`| ist` Jinja filter). India does not observe daylight saving, so no
DST handling is needed here.
"""

from __future__ import annotations

import logging
import os
import re
from datetime import date
from typing import Any

logger = logging.getLogger(__name__)

DATABASE_URL_ENV_VAR = "DATABASE_URL"
_TIME_RE = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")


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
        logger.warning("Database unavailable for scheduler_settings: %s", exc)
        return None


def is_valid_time(value: str) -> bool:
    """True if *value* is a well-formed 24-hour "HH:MM" string."""
    return bool(_TIME_RE.match(value or ""))


def get_scheduler_settings() -> dict[str, Any] | None:
    """Return {"processing_time", "enabled", "last_run_date", "updated_at"},
    or None if unconfigured / DATABASE_URL unset. Callers must treat None
    as "automatic schedule is off" — never raise."""
    conn = _connect_or_none()
    if conn is None:
        return None
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT processing_time, enabled, last_run_date, updated_at "
                "FROM scheduler_settings WHERE id = 1"
            )
            row = cur.fetchone()
            if row is None:
                return None
            cols = ["processing_time", "enabled", "last_run_date", "updated_at"]
            return dict(zip(cols, row))
    except Exception as exc:
        logger.warning("Could not read scheduler_settings from database: %s", exc)
        return None
    finally:
        conn.close()


def upsert_scheduler_settings(processing_time: str, enabled: bool) -> None:
    """Insert or update the single schedule row. Requires DATABASE_URL.
    Caller (web_app.py) is responsible for validating processing_time
    with is_valid_time() first — this function trusts its input."""
    conn = _get_connection()
    if conn is None:
        raise RuntimeError("DATABASE_URL is not configured; cannot save the schedule.")

    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO scheduler_settings (id, processing_time, enabled, updated_at) "
                "VALUES (1, %s, %s, NOW()) "
                "ON CONFLICT (id) DO UPDATE "
                "SET processing_time = EXCLUDED.processing_time, "
                "    enabled = EXCLUDED.enabled, "
                "    updated_at = NOW()",
                (processing_time, enabled),
            )
        conn.commit()
    finally:
        conn.close()


def try_claim_run_for_today(today: date) -> bool:
    """Atomically claim today's automatic run: only succeeds (returns True)
    if last_run_date isn't already today, and immediately marks it as
    today in the same statement — so two overlapping cron pings (or
    multiple serverless instances handling pings close together) can
    never both win this race. Returns False if today's run was already
    claimed (by this call or an earlier one), or if unconfigured."""
    conn = _get_connection()
    if conn is None:
        return False
    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE scheduler_settings SET last_run_date = %s "
                "WHERE id = 1 AND (last_run_date IS DISTINCT FROM %s) "
                "RETURNING id",
                (today, today),
            )
            claimed = cur.fetchone() is not None
        conn.commit()
        return claimed
    except Exception as exc:
        logger.warning("Could not claim scheduler run for %s: %s", today, exc)
        return False
    finally:
        conn.close()
