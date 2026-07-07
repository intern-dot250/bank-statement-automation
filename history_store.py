"""Persistent storage for processing history and latest-batch stats.

Backend selection:
  - If DATABASE_URL is set: uses Postgres (e.g. Supabase). This is the
    durable option for serverless deployments, where local JSON files
    under /tmp don't survive between invocations/cold starts.
  - Otherwise: falls back to local JSON files, matching this project's
    existing local/non-serverless behavior — no database is required to
    run the app locally.

Every public function accepts the same local JSON path the caller was
already using, so it can fall back to it transparently.
"""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

DATABASE_URL_ENV_VAR = "DATABASE_URL"

_schema_ready = False


def _get_connection():
    """Return a psycopg2 connection to DATABASE_URL, or None if unset.

    Raises on connection failure so callers can log and fall back.
    """
    database_url = os.environ.get(DATABASE_URL_ENV_VAR)
    if not database_url:
        logger.info("DATABASE_URL not set; using local JSON file storage.")
        return None

    import psycopg2  # imported lazily so this stays optional for local use

    conn = psycopg2.connect(database_url)
    _ensure_schema(conn)
    logger.info("Connected to Postgres database.")
    return conn


def _ensure_schema(conn) -> None:
    global _schema_ready
    if _schema_ready:
        return
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS processing_history (
                id SERIAL PRIMARY KEY,
                request_id TEXT,
                timestamp TEXT,
                file TEXT,
                bank TEXT,
                status TEXT,
                total_rows INTEGER DEFAULT 0,
                new_rows INTEGER DEFAULT 0,
                duplicates_skipped INTEGER DEFAULT 0,
                total_rows_in_pdf INTEGER DEFAULT 0,
                sheet_url TEXT,
                error TEXT,
                failed_stage INTEGER,
                source TEXT,
                created_at TIMESTAMPTZ DEFAULT now()
            );
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS latest_batch (
                id INTEGER PRIMARY KEY DEFAULT 1,
                processed INTEGER DEFAULT 0,
                success INTEGER DEFAULT 0,
                failed INTEGER DEFAULT 0,
                updated_at TIMESTAMPTZ DEFAULT now(),
                CHECK (id = 1)
            );
        """)
    conn.commit()
    _schema_ready = True


def _connect_or_none():
    try:
        return _get_connection()
    except Exception as exc:
        logger.warning("Database unavailable, falling back to local file: %s", exc)
        return None


# ---------------------------------------------------------------------------
# Processing history
# ---------------------------------------------------------------------------

_HISTORY_COLUMNS = [
    "request_id", "timestamp", "file", "bank", "status", "total_rows",
    "new_rows", "duplicates_skipped", "total_rows_in_pdf", "sheet_url",
    "error", "failed_stage", "source",
]


def load_history(fallback_path: Path) -> list[dict[str, Any]]:
    """Load all processing history entries, oldest first."""
    conn = _connect_or_none()
    if conn is not None:
        try:
            with conn.cursor() as cur:
                cur.execute(
                    f"SELECT {', '.join(_HISTORY_COLUMNS)} FROM processing_history ORDER BY id ASC"
                )
                return [dict(zip(_HISTORY_COLUMNS, row)) for row in cur.fetchall()]
        except Exception as exc:
            logger.warning("Could not read history from database: %s", exc)
            return []
        finally:
            conn.close()

    if not fallback_path.exists():
        return []
    try:
        with open(fallback_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning("Could not load history file: %s", exc)
        return []


def save_history_entry(entry: dict[str, Any], fallback_path: Path) -> None:
    """Append one processing-history entry."""
    conn = _connect_or_none()
    if conn is not None:
        try:
            with conn.cursor() as cur:
                cur.execute(
                    f"""
                    INSERT INTO processing_history ({', '.join(_HISTORY_COLUMNS)})
                    VALUES ({', '.join(['%s'] * len(_HISTORY_COLUMNS))})
                    """,
                    tuple(entry.get(col) for col in _HISTORY_COLUMNS),
                )
            conn.commit()
            logger.debug("History entry saved to database.")
            return
        except Exception as exc:
            logger.error("Could not save history entry to database: %s", exc)
            return
        finally:
            conn.close()

    history = load_history(fallback_path)
    history.append(entry)
    if len(history) > 500:
        history = history[-500:]
    fallback_path.parent.mkdir(parents=True, exist_ok=True)
    with open(fallback_path, "w", encoding="utf-8") as f:
        json.dump(history, f, indent=2, default=str, ensure_ascii=False)
    logger.debug("History entry saved to local file.")


def update_history_source(
    fallback_path: Path,
    request_id: str | None,
    filename: str | None,
    source: str,
) -> None:
    """Set the 'source' field on the most recent matching entry.

    Matches by request_id when available, otherwise by filename.
    """
    conn = _connect_or_none()
    if conn is not None:
        try:
            with conn.cursor() as cur:
                if request_id:
                    cur.execute(
                        """
                        UPDATE processing_history SET source = %s
                        WHERE id = (
                            SELECT id FROM processing_history
                            WHERE request_id = %s ORDER BY id DESC LIMIT 1
                        )
                        """,
                        (source, request_id),
                    )
                else:
                    cur.execute(
                        """
                        UPDATE processing_history SET source = %s
                        WHERE id = (
                            SELECT id FROM processing_history
                            WHERE file = %s ORDER BY id DESC LIMIT 1
                        )
                        """,
                        (source, filename),
                    )
            conn.commit()
        except Exception as exc:
            logger.warning("Could not update history source in database: %s", exc)
        finally:
            conn.close()
        return

    if not fallback_path.exists():
        return
    try:
        with open(fallback_path, "r", encoding="utf-8") as f:
            hist = json.load(f)
        for entry in reversed(hist):
            if (request_id and entry.get("request_id") == request_id) or entry.get("file") == filename:
                entry["source"] = source
                break
        with open(fallback_path, "w", encoding="utf-8") as f:
            json.dump(hist, f, indent=2, ensure_ascii=False)
    except Exception as exc:
        logger.warning("Could not update history source in local file: %s", exc)


# ---------------------------------------------------------------------------
# Latest batch
# ---------------------------------------------------------------------------

_DEFAULT_BATCH = {"processed": 0, "success": 0, "failed": 0}


def load_latest_batch(fallback_path: Path) -> dict[str, int]:
    conn = _connect_or_none()
    if conn is not None:
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT processed, success, failed FROM latest_batch WHERE id = 1")
                row = cur.fetchone()
            if row is None:
                return dict(_DEFAULT_BATCH)
            return {"processed": row[0], "success": row[1], "failed": row[2]}
        except Exception as exc:
            logger.warning("Could not read latest_batch from database: %s", exc)
            return dict(_DEFAULT_BATCH)
        finally:
            conn.close()

    if not fallback_path.exists():
        return dict(_DEFAULT_BATCH)
    try:
        with open(fallback_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("latest_batch", dict(_DEFAULT_BATCH))
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning("Could not read latest_batch from local file: %s", exc)
        return dict(_DEFAULT_BATCH)


def save_latest_batch(batch_stats: dict, fallback_path: Path) -> None:
    conn = _connect_or_none()
    if conn is not None:
        try:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO latest_batch (id, processed, success, failed, updated_at)
                    VALUES (1, %s, %s, %s, now())
                    ON CONFLICT (id) DO UPDATE SET
                        processed = EXCLUDED.processed,
                        success = EXCLUDED.success,
                        failed = EXCLUDED.failed,
                        updated_at = EXCLUDED.updated_at
                    """,
                    (
                        batch_stats.get("processed", 0),
                        batch_stats.get("success", 0),
                        batch_stats.get("failed", 0),
                    ),
                )
            conn.commit()
            logger.info("Saved latest_batch to database: %s", batch_stats)
            return
        except Exception as exc:
            logger.error("Could not save latest_batch to database: %s", exc)
            return
        finally:
            conn.close()

    data = {}
    if fallback_path.exists():
        try:
            with open(fallback_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except json.JSONDecodeError:
            logger.error("Malformed JSON in records.json — latest_batch will overwrite it.")
            data = {}
    data["latest_batch"] = batch_stats
    try:
        with open(fallback_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        logger.info("Saved latest_batch to local file: %s", batch_stats)
    except Exception as exc:
        logger.error("Could not save latest_batch to local file: %s", exc)
