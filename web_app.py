"""Bank Statement Processor — Web Application.

A Flask-based web application for uploading and processing bank statement PDFs.
Non-technical employees can use this through a simple web browser.

Architecture:
- Flask web app handles uploads, routing, and display
- Calls existing run_pipeline.py as subprocess
- Tracks processing status in logs/app.log
- Never exposes credentials.json to the web

Usage:
    py web_app.py
    py web_app.py --port 8080 --host 0.0.0.0
"""

from __future__ import annotations

import functools
import json
import logging
import os
import shutil
import sys
import threading
import time
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import gspread

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

from upload_to_sheets import (
    DEFAULT_CREDENTIALS,
    MASTER_SHEET_ID,
    get_account_worksheets,
    get_gspread_client,
)
from email_reader import save_latest_batch, process_emails
from run_pipeline import run_pipeline as run_pipeline_fn
from runtime_paths import base_data_dir
import auth
import credentials_store
import history_store

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = base_data_dir(SCRIPT_DIR)
CONFIG_PATH = SCRIPT_DIR / "config.json"  # config.json ships with the code; read-only is fine
RECORDS_PATH = DATA_DIR / "records.json"
HISTORY_PATH = DATA_DIR / "logs" / "processing_history.json"
STATUS_PATH = DATA_DIR / "logs" / "processing_status.json"
LOG_PATH = DATA_DIR / "logs" / "web_app.log"
INPUT_DIR = DATA_DIR / "input"
PROCESSED_DIR = DATA_DIR / "processed"
FAILED_DIR = DATA_DIR / "failed"


# Persisted (Postgres-backed on Vercel, JSON file locally) rather than an
# in-memory dict — the HTTP request that starts a background thread and
# the later polling requests checking its progress can each land on a
# DIFFERENT serverless instance, so an in-memory dict populated by one
# instance is invisible to the others.
def _get_status(filename: str) -> dict[str, Any] | None:
    return history_store.load_processing_status(filename, STATUS_PATH)


def _set_status(filename: str, status: dict[str, Any]) -> None:
    history_store.save_processing_status(filename, status, STATUS_PATH)


def _update_status(filename: str, updates: dict[str, Any]) -> None:
    current = _get_status(filename) or {}
    current.update(updates)
    _set_status(filename, current)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
# File logging can fail on a read-only filesystem (e.g. a serverless
# deployment such as Vercel), where LOG_PATH's parent directory can't be
# created. Fall back to stdout-only logging rather than crashing on import.
_log_handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
try:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    _log_handlers.append(logging.FileHandler(str(LOG_PATH), encoding="utf-8"))
except OSError:
    pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    handlers=_log_handlers,
    force=True,
)
log = logging.getLogger("web_app")

# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------
app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "CHANGE_ME_IN_PRODUCTION")


@app.template_filter("ist")
def format_ist(timestamp_str: str) -> str:
    """Format a stored timestamp (naive, server-local/UTC) as a readable
    IST (UTC+5:30) datetime string, e.g. "07 Jul 2026, 05:07 PM IST"."""
    if not timestamp_str:
        return "Unknown"
    try:
        dt = datetime.fromisoformat(timestamp_str)
        dt_ist = dt + timedelta(hours=5, minutes=30)
        return dt_ist.strftime("%d %b %Y, %I:%M %p IST")
    except ValueError:
        return timestamp_str

# Security: limit upload size to 25 MB
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

# Block access to sensitive files
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def load_config() -> dict:
    """Load and validate config.json."""
    if not CONFIG_PATH.exists():
        log.error("config.json not found at %s", CONFIG_PATH)
        raise FileNotFoundError("Server configuration missing.")
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)
    return config


def load_latest_batch() -> dict[str, int]:
    """Load current-batch PDF processing counts from records.json.

    Returns:
        Dict with "processed", "success", "failed" keys (all 0 if the
        batch has never run yet or the file/key is missing).
    """
    return history_store.load_latest_batch(RECORDS_PATH)


def get_live_sheet_row_count() -> int:
    """Return the current live row count summed across every account's
    worksheet tab (there is no single master sheet — each account has
    its own tab, e.g. "YES BANK - 2477").

    Returns:
        Total data rows across all account tabs (0 if none have data yet).
    """
    credentials_path = SCRIPT_DIR / DEFAULT_CREDENTIALS

    client = get_gspread_client(credentials_path)
    spreadsheet = client.open_by_key(MASTER_SHEET_ID)

    total_rows = 0
    for worksheet in get_account_worksheets(spreadsheet):
        rows = worksheet.get_all_values()
        total_rows += max(len(rows) - 1, 0)

    return total_rows


# ---------------------------------------------------------------------------
# /sheet_rows in-memory cache
# ---------------------------------------------------------------------------
# Avoids hitting the Google Sheets API on every dashboard poll. Only
# /sheet_rows reads this cache — no other route is affected.
_SHEET_ROWS_CACHE_TTL_SECONDS = 30

_sheet_rows_cache: dict[str, Any] = {
    "row_count": None,   # last known row count (None until first successful read)
    "timestamp": 0.0,    # time.time() of that read
}
_sheet_rows_cache_lock = threading.Lock()


def _is_quota_error(exc: Exception) -> bool:
    """True if exc is a gspread APIError caused by a 429 quota response."""
    if not isinstance(exc, gspread.exceptions.APIError):
        return False
    try:
        return exc.response.status_code == 429
    except Exception:
        return False


def sanitize_filename(filename: str) -> str:
    """Sanitize filename to prevent path traversal attacks.

    Args:
        filename: Original filename from upload.

    Returns:
        Safe filename with timestamp prefix.
    """
    # Remove any path components
    filename = os.path.basename(filename)
    # Remove non-alphanumeric chars (keep dots, hyphens, underscores)
    import re
    filename = re.sub(r"[^a-zA-Z0-9._\-]", "_", filename)
    # Ensure it ends with .pdf
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    # Prefix with timestamp and UUID to prevent collisions
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    short_id = str(uuid.uuid4())[:8]
    return f"{timestamp}_{short_id}_{filename}"


def cleanup_directories() -> None:
    """Clean up processed/ and failed/ directories, keeping only recent files."""
    try:
        config = load_config()
        folders = config.get("folders", {})
        processed_dir = SCRIPT_DIR / folders.get("processed", "processed")
        failed_dir = SCRIPT_DIR / folders.get("failed", "failed")

        def clean_dir(d: Path, keep_count: int):
            if not d.exists():
                return
            files = [f for f in d.iterdir() if f.is_file()]
            files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            for f in files[keep_count:]:
                try:
                    f.unlink()
                    log.info("Cleanup: Deleted %s", f.name)
                except Exception as e:
                    log.warning("Cleanup: Could not delete %s: %s", f.name, e)

        clean_dir(processed_dir, 10)
        clean_dir(failed_dir, 5)
    except Exception as exc:
        log.warning("Error during directory cleanup: %s", exc)


def run_pipeline_in_thread(
    filename: str,
    pdf_path: Path,
    password: str,
    bank_name: str,
    account_number: str = "",
) -> None:
    """Run the pipeline in a background thread.

    Args:
        filename: Unique filename serving as the process/file key.
        pdf_path: Path to the input PDF.
        password: PDF password.
        bank_name: Name of the bank.
        account_number: Account number this PDF belongs to — determines
            which account's own worksheet tab the rows are uploaded into.
    """
    log.info("Background thread started for file: %s (bank: %s)", filename, bank_name)
    try:
        config = load_config()
        log.info("Successfully loaded configuration for file: %s", filename)
    except Exception as exc:
        log.error("Configuration error for file %s: %s", filename, exc)
        _set_status(filename, {
            "status": "failed",
            "error": f"Configuration error: {exc}",
            "progress": 0,
        })
        save_latest_batch({"processed": 1, "success": 0, "failed": 1})
        return

    # Update status: starting
    log.info("Updating status to 'processing' for file: %s", filename)
    _set_status(filename, {
        "status": "processing",
        "message": "Starting pipeline...",
        "progress": 10,
        "filename": pdf_path.name,
        "bank_name": bank_name,
        "timestamp": datetime.now().isoformat(),
        "total_rows": 0,
        "new_rows": 0,
        "duplicates_skipped": 0,
    })

    # Call run_pipeline() directly, in-process (no subprocess — unreliable
    # on serverless deployments such as Vercel, and avoids process-startup
    # overhead everywhere else too).
    try:
        log.info("Running pipeline in-process for file: %s", filename)
        success, result = run_pipeline_fn(
            password=password,
            input_pdf=pdf_path,
            config=config,
            bank_name=bank_name,
            account_number=account_number,
            logger=log,
        )
        log.info("Pipeline for file %s finished, success=%s", filename, success)

        total_rows = result.get("total_rows", 0)
        new_rows = result.get("new_rows", 0)
        duplicates_skipped = result.get("duplicates_skipped", 0)
        sheet_url = result.get("sheet_url", "")
        child_req_id = result.get("request_id", "")

        # Update history with source="Manual"
        try:
            history_store.update_history_source(HISTORY_PATH, child_req_id, pdf_path.name, "Manual")
        except Exception as e:
            log.warning("Could not update history source: %s", e)

        log.info(
            "Parsed results for file %s: total_rows=%d, new_rows=%d, duplicates_skipped=%d",
            filename, total_rows, new_rows, duplicates_skipped
        )

        if success:
            log.info("Pipeline executed successfully for file: %s", filename)
            _update_status(filename, {
                "status": "completed",
                "message": "Processing complete!",
                "progress": 100,
                "total_rows": total_rows,
                "new_rows": new_rows,
                "duplicates_skipped": duplicates_skipped,
            })
            save_latest_batch({"processed": 1, "success": 1, "failed": 0})
        else:
            error_msg = result.get("error") or "Pipeline failed (no error message captured)"
            log.error("Pipeline execution failed for file: %s. Error message: %s", filename, error_msg)
            _update_status(filename, {
                "status": "failed",
                "message": "Processing failed",
                "error": error_msg,
                "progress": 0,
            })
            save_latest_batch({"processed": 1, "success": 0, "failed": 1})

    except Exception as exc:
        log.exception("Unexpected exception in background thread for file %s", filename)
        _update_status(filename, {
            "status": "failed",
            "error": str(exc),
            "progress": 0,
        })
        save_latest_batch({"processed": 1, "success": 0, "failed": 1})
    finally:
        cleanup_directories()


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
def login_required(view):
    """Redirect to /login if the current session isn't authenticated."""
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("authenticated"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


@app.route("/login", methods=["GET", "POST"])
def login():
    """Password-only login backed by Supabase Auth (single shared account)."""
    if request.method == "GET":
        return render_template("login.html")

    password = request.form.get("password", "")
    tokens = auth.login(password)

    if tokens is None:
        flash("Incorrect password.", "error")
        return render_template("login.html"), 401

    session["authenticated"] = True
    session["refresh_token"] = tokens["refresh_token"]

    next_path = request.args.get("next")
    return redirect(next_path or url_for("index"))


@app.route("/logout", methods=["POST"])
def logout():
    """Clear the session (and best-effort sign out of Supabase)."""
    auth.logout(session.get("refresh_token"))
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route("/", methods=["GET"])
@login_required
def index():
    """Upload page."""
    try:
        config = load_config()
        sheet_url = config.get("sheet_url", "#")
    except Exception:
        sheet_url = "#"
    return render_template("index.html", sheet_url=sheet_url)


@app.route("/upload", methods=["POST"])
@login_required
def upload_file():
    """Handle file upload.

    Validates file type, sanitizes filename, saves to input/.
    Returns unique request ID.
    """
    try:
        # Check if file was included
        if "pdf_file" not in request.files:
            return jsonify({"error": "No file selected."}), 400

        file = request.files["pdf_file"]
        if file.filename == "":
            return jsonify({"error": "No file selected."}), 400

        # Validate file extension
        if not file.filename.lower().endswith(".pdf"):
            return jsonify({"error": "Invalid file type. Only PDF files are allowed."}), 400

        # Save with sanitized name
        safe_name = sanitize_filename(file.filename)
        filepath = INPUT_DIR / safe_name
        file.save(str(filepath))
        log.info("Uploaded file: %s (original: %s)", safe_name, file.filename)

        return jsonify({
            "filename": safe_name,
            "original_name": file.filename,
            "status": "uploaded",
        })

    except Exception as exc:
        log.exception("Upload error")
        return jsonify({"error": f"Upload failed: {exc}"}), 500


@app.route("/process", methods=["POST"])
@login_required
def process_file():
    """Start pipeline processing in background."""
    try:
        data = request.get_json()
        if not data:
            log.warning("Process endpoint received empty JSON body.")
            return jsonify({"error": "Invalid request."}), 400

        filename = data.get("filename")
        password = data.get("password", "").strip()
        bank_name = data.get("bank_name", "YES BANK")
        account_number = data.get("account_number", "").strip()

        if not filename:
            log.warning("Process endpoint called without filename.")
            return jsonify({"error": "Filename is required."}), 400
        if not password:
            log.warning("Process endpoint called without password for file: %s", filename)
            return jsonify({"error": "Password is required."}), 400
        if not account_number:
            log.warning("Process endpoint called without account_number for file: %s", filename)
            return jsonify({"error": "Account Number is required."}), 400

        pdf_path = INPUT_DIR / filename
        if not pdf_path.exists():
            log.error("Process file not found: %s", pdf_path)
            return jsonify({"error": "Uploaded file not found. Please re-upload."}), 404

        # Generate unique request ID (keep for logging/audit, but index status by filename)
        request_id = str(uuid.uuid4())[:12]
        log.info("Initializing status store for file: %s (request ID: %s)", filename, request_id)

        # Initialize persisted status, keyed by filename
        _set_status(filename, {
            "status": "processing",
            "message": "Initializing...",
            "progress": 5,
            "filename": filename,
            "bank_name": bank_name,
            "timestamp": datetime.now().isoformat(),
            "total_rows": 0,
            "new_rows": 0,
            "duplicates_skipped": 0,
        })

        # Start background thread
        log.info("Spawning background thread to process file: %s", filename)
        thread = threading.Thread(
            target=run_pipeline_in_thread,
            args=(filename, pdf_path, password, bank_name, account_number),
            daemon=True,
        )
        thread.start()

        log.info("Process endpoint returning success response for file: %s", filename)
        return jsonify({
            "status": "processing",
            "request_id": request_id,
            "message": "Processing started. Please wait...",
        })

    except Exception as exc:
        log.exception("Error in process_file endpoint")
        return jsonify({"error": str(exc)}), 500


@app.route("/status/<filename>", methods=["GET"])
@login_required
def check_status(filename: str):
    """Check processing status by filename."""
    status = _get_status(filename)

    if not status:
        return jsonify({"status": "unknown", "message": "Status not found."})

    return jsonify(status)


@app.route("/success/<filename>", methods=["GET"])
@login_required
def success_page(filename: str):
    """Success page after processing."""
    status = _get_status(filename)

    if not status:
        return redirect(url_for("index"))

    data = {
        "filename": status.get("filename", filename),
        "bank_name": status.get("bank_name", "YES BANK"),
        "total_rows": status.get("total_rows", 0),
        "new_rows": status.get("new_rows", 0),
        "duplicates_skipped": status.get("duplicates_skipped", 0),
    }

    return render_template("success.html", data=data)


@app.route("/error/<filename>", methods=["GET"])
@login_required
def error_page(filename: str):
    """Error page after processing failure."""
    status = _get_status(filename)

    if not status:
        return redirect(url_for("index"))

    data = {
        "filename": status.get("filename", filename),
        "bank_name": status.get("bank_name", "YES BANK"),
        "timestamp": status.get("timestamp"),
    }

    return render_template("error.html", error=status.get("error", "Unknown error"), data=data)


@app.route("/history", methods=["GET"])
@login_required
def history():
    """Display processing history."""
    try:
        config = load_config()
        history_entries = history_store.load_history(HISTORY_PATH)

        # Inject default source if missing
        for entry in history_entries:
            if "source" not in entry:
                entry["source"] = "Email"

        # Sort by timestamp descending (most recent first)
        history_entries.sort(key=lambda x: x.get("timestamp", ""), reverse=True)
        
        sheet_url = config.get("sheet_url", "#")

        return render_template("history.html", history=history_entries, sheet_url=sheet_url)

    except Exception as exc:
        log.exception("History page error")
        return render_template("history.html", history=[])


@app.route("/history/<request_id>/delete", methods=["POST"])
@login_required
def history_delete(request_id: str):
    """Delete a single processing-history entry by its request_id."""
    try:
        history_store.delete_history_entry(request_id, HISTORY_PATH)
        flash("History entry deleted.", "success")
    except Exception as exc:
        log.warning("Could not delete history entry %s: %s", request_id, exc)
        flash(f"Could not delete entry: {exc}", "error")

    return redirect(url_for("history"))


@app.route("/latest_batch", methods=["GET"])
@login_required
def latest_batch():
    """Return current-batch PDF processing counts (not lifetime totals)."""
    return jsonify(load_latest_batch())


@app.route("/sheet_rows", methods=["GET"])
@login_required
def sheet_rows():
    """Return the live current row count in the master Google Sheet.

    Cached in-memory for _SHEET_ROWS_CACHE_TTL_SECONDS to avoid calling
    the Google Sheets API on every dashboard poll. On a 429 quota error,
    falls back to the last cached value instead of failing the request
    (only returns an error/500 if no cached value exists yet at all).
    """
    now = time.time()

    with _sheet_rows_cache_lock:
        cache_age = now - _sheet_rows_cache["timestamp"]
        if _sheet_rows_cache["row_count"] is not None and cache_age < _SHEET_ROWS_CACHE_TTL_SECONDS:
            return jsonify({"total_rows": _sheet_rows_cache["row_count"], "cached": True})

    try:
        total_rows = get_live_sheet_row_count()
        with _sheet_rows_cache_lock:
            _sheet_rows_cache["row_count"] = total_rows
            _sheet_rows_cache["timestamp"] = time.time()
        return jsonify({"total_rows": total_rows, "cached": False})
    except Exception as exc:
        if _is_quota_error(exc) and _sheet_rows_cache["row_count"] is not None:
            log.warning(
                "Google Sheets quota exceeded (429) on /sheet_rows — "
                "returning last cached row count (%s): %s",
                _sheet_rows_cache["row_count"], exc,
            )
            return jsonify({
                "total_rows": _sheet_rows_cache["row_count"],
                "cached": True,
                "warning": "quota_exceeded",
            })
        log.error("Could not fetch live sheet row count: %s", exc)
        return jsonify({"total_rows": 0, "error": str(exc)}), 500


@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint."""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "version": "1.0.0",
    })


@app.route("/check_emails", methods=["POST"])
@login_required
def check_emails():
    """Trigger email checking manually.

    Calls process_emails() directly, in-process (no subprocess —
    unreliable on serverless deployments such as Vercel), and uses its
    returned batch_stats dict for a precise status instead of scraping
    subprocess stdout text for phrases.
    """
    try:
        batch_stats = process_emails()
        cleanup_directories()

        if batch_stats.get("processed", 0) == 0:
            return jsonify({"status": "no_emails", "message": "No unread emails found"})
        if batch_stats.get("failed", 0) > 0:
            return jsonify({"status": "failed", "message": "Failed to process some emails", "batch": batch_stats})
        return jsonify({"status": "success", "message": "Successfully processed emails", "batch": batch_stats})
    except Exception as e:
        log.exception("Error checking emails")
        return jsonify({"status": "failed", "error": str(e)}), 500


@app.route("/accounts_list", methods=["GET"])
@login_required
def accounts_list():
    """Return configured accounts (account_number, password, bank_name)
    for the manual upload form's Account Number dropdown/autofill."""
    accounts = credentials_store.list_credentials(RECORDS_PATH)
    return jsonify([
        {
            "account_number": acc.get("account_number"),
            "password": acc.get("password"),
            "bank_name": acc.get("bank_name"),
        }
        for acc in accounts
    ])


@app.route("/admin/passwords", methods=["GET"])
@login_required
def admin_passwords():
    """Admin page listing/managing bank account -> PDF-password mappings."""
    accounts = credentials_store.list_credentials(RECORDS_PATH)

    # Bank Name dropdown options: the pipeline's supported banks, plus any
    # additional bank names already saved in the accounts list.
    try:
        config = load_config()
        supported_bank_names = [
            b.get("display_name") for b in config.get("supported_banks", {}).values()
        ]
    except Exception:
        supported_bank_names = []

    existing_bank_names = [acc.get("bank_name") for acc in accounts]
    bank_names = sorted({name for name in supported_bank_names + existing_bank_names if name})

    return render_template("admin_passwords.html", accounts=accounts, bank_names=bank_names)


@app.route("/admin/passwords/add", methods=["POST"])
@login_required
def admin_passwords_add():
    """Add a new bank account credential (requires DATABASE_URL)."""
    bank_name = request.form.get("bank_name", "").strip()
    account_number = request.form.get("account_number", "").strip()
    password = request.form.get("password", "").strip()

    if not bank_name or not account_number or not password:
        flash("Bank name, account number, and password are all required.", "error")
        return redirect(url_for("admin_passwords"))

    try:
        credentials_store.add_credential(bank_name, account_number, password)
        flash(f"Added account {account_number}.", "success")
    except Exception as exc:
        log.warning("Could not add account credential: %s", exc)
        flash(f"Could not add account: {exc}", "error")

    return redirect(url_for("admin_passwords"))


@app.route("/admin/passwords/<int:credential_id>/delete", methods=["POST"])
@login_required
def admin_passwords_delete(credential_id: int):
    """Delete a bank account credential by id (requires DATABASE_URL)."""
    try:
        credentials_store.delete_credential(credential_id)
        flash("Account deleted.", "success")
    except Exception as exc:
        log.warning("Could not delete account credential %s: %s", credential_id, exc)
        flash(f"Could not delete account: {exc}", "error")

    return redirect(url_for("admin_passwords"))


# ---------------------------------------------------------------------------
# Error handlers
# ---------------------------------------------------------------------------
@app.errorhandler(413)
def request_entity_too_large(error):
    """Handle file too large error."""
    return jsonify({"error": "File too large. Maximum size is 25 MB."}), 413


@app.errorhandler(500)
def internal_error(error):
    """Handle internal server error."""
    log.exception("Internal server error")
    return jsonify({"error": "An internal error occurred."}), 500


@app.errorhandler(404)
def not_found(error):
    """Handle not found error."""
    return jsonify({"error": "Page not found."}), 404


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> int:
    """Start the Flask web server."""
    try:
        config = load_config()
    except Exception as exc:
        log.error("Failed to load config: %s", exc)
        return 1

    web_config = config.get("web_app", {})
    host = web_config.get("host", "0.0.0.0")
    port = web_config.get("port", 5000)
    debug = web_config.get("debug", False)

    log.info("=" * 50)
    log.info("Bank Statement Processor — Web App")
    log.info("=" * 50)
    log.info("Server starting on http://%s:%d", host, port)
    log.info("Open in browser: http://localhost:%d", port)
    log.info("=" * 50)

    app.run(
        host=host,
        port=port,
        debug=debug,
        use_reloader=False,  # Disable reloader for stability
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
