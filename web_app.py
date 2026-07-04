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

import json
import logging
import os
import shutil
import subprocess
import sys
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    url_for,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
LOG_PATH = SCRIPT_DIR / "logs" / "web_app.log"
INPUT_DIR = SCRIPT_DIR / "input"
PROCESSED_DIR = SCRIPT_DIR / "processed"
FAILED_DIR = SCRIPT_DIR / "failed"

# Processing status store (in-memory, production would use Redis/DB)
processing_status: dict[str, dict[str, Any]] = {}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    handlers=[
        logging.FileHandler(str(LOG_PATH), encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("web_app")

# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------
app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "CHANGE_ME_IN_PRODUCTION")

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
) -> None:
    """Run the pipeline in a background thread.

    Args:
        filename: Unique filename serving as the process/file key.
        pdf_path: Path to the input PDF.
        password: PDF password.
        bank_name: Name of the bank.
    """
    log.info("Background thread started for file: %s (bank: %s)", filename, bank_name)
    try:
        config = load_config()
        log.info("Successfully loaded configuration for file: %s", filename)
    except Exception as exc:
        log.error("Configuration error for file %s: %s", filename, exc)
        processing_status[filename] = {
            "status": "failed",
            "error": f"Configuration error: {exc}",
            "progress": 0,
        }
        return

    # Update status: starting
    log.info("Updating status to 'processing' for file: %s", filename)
    processing_status[filename] = {
        "status": "processing",
        "message": "Starting pipeline...",
        "progress": 10,
        "filename": pdf_path.name,
        "bank_name": bank_name,
        "timestamp": datetime.now().isoformat(),
        "total_rows": 0,
        "new_rows": 0,
        "duplicates_skipped": 0,
    }

    # Call run_pipeline.py as subprocess
    try:
        log.info("Launching run_pipeline.py subprocess for file: %s", filename)
        result = subprocess.run(
            [
                sys.executable,
                str(SCRIPT_DIR / "run_pipeline.py"),
                "--password", password,
                "--input", str(pdf_path),
                "--config", str(CONFIG_PATH),
            ],
            capture_output=True,
            text=True,
            timeout=300,  # 5 minute timeout
        )
        log.info("Pipeline subprocess for file %s finished with exit code: %d", filename, result.returncode)

        # Parse output for metrics JSON
        stdout = result.stdout or ""
        stderr = result.stderr or ""

        total_rows = 0
        new_rows = 0
        duplicates_skipped = 0
        sheet_url = ""
        child_req_id = ""

        lines = [line.strip() for line in stdout.split("\n") if line.strip()]
        if lines:
            last_line = lines[-1]
            if last_line.startswith("{") and last_line.endswith("}"):
                try:
                    metrics = json.loads(last_line)
                    total_rows = metrics.get("total_rows", 0)
                    new_rows = metrics.get("new_rows", 0)
                    duplicates_skipped = metrics.get("duplicates_skipped", 0)
                    sheet_url = metrics.get("sheet_url", "")
                    child_req_id = metrics.get("request_id", "")
                except json.JSONDecodeError:
                    pass

        # Update history with source="Manual"
        try:
            history_path = SCRIPT_DIR / config.get("processing", {}).get("history_file", "logs/processing_history.json")
            if history_path.exists():
                with open(history_path, "r", encoding="utf-8") as f:
                    hist = json.load(f)
                for entry in reversed(hist):
                    if (child_req_id and entry.get("request_id") == child_req_id) or entry.get("file") == pdf_path.name:
                        entry["source"] = "Manual"
                        break
                with open(history_path, "w", encoding="utf-8") as f:
                    json.dump(hist, f, indent=2, ensure_ascii=False)
        except Exception as e:
            log.warning("Could not update history source: %s", e)

        log.info(
            "Parsed results for file %s: total_rows=%d, new_rows=%d, duplicates_skipped=%d",
            filename, total_rows, new_rows, duplicates_skipped
        )

        if result.returncode == 0:
            log.info("Pipeline executed successfully for file: %s", filename)
            processing_status[filename].update({
                "status": "completed",
                "message": "Processing complete!",
                "progress": 100,
                "total_rows": total_rows,
                "new_rows": new_rows,
                "duplicates_skipped": duplicates_skipped,
            })
        else:
            # Log full stderr so no detail is lost
            if stderr.strip():
                log.error("Full stderr for file %s:\n%s", filename, stderr.strip())
            if stdout.strip():
                log.error("Full stdout for file %s:\n%s", filename, stdout.strip())
            # Extract the last non-empty line as the display error
            non_empty_lines = [l.strip() for l in stderr.strip().splitlines() if l.strip()]
            error_msg = non_empty_lines[-1] if non_empty_lines else "Pipeline failed (no error output captured)"
            log.error("Pipeline execution failed for file: %s. Error message: %s", filename, error_msg)
            processing_status[filename].update({
                "status": "failed",
                "message": "Processing failed",
                "error": error_msg,
                "progress": 0,
            })

    except subprocess.TimeoutExpired:
        log.error("Pipeline execution timed out (5 minutes) for file: %s", filename)
        processing_status[filename].update({
            "status": "failed",
            "error": "Processing timed out (5 minutes).",
            "progress": 0,
        })
    except Exception as exc:
        log.exception("Unexpected exception in background thread for file %s", filename)
        processing_status[filename].update({
            "status": "failed",
            "error": str(exc),
            "progress": 0,
        })
    finally:
        cleanup_directories()


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    """Upload page."""
    try:
        config = load_config()
        sheet_url = config.get("sheet_url", "#")
    except Exception:
        sheet_url = "#"
    return render_template("index.html", sheet_url=sheet_url)


@app.route("/upload", methods=["POST"])
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

        if not filename:
            log.warning("Process endpoint called without filename.")
            return jsonify({"error": "Filename is required."}), 400
        if not password:
            log.warning("Process endpoint called without password for file: %s", filename)
            return jsonify({"error": "Password is required."}), 400

        pdf_path = INPUT_DIR / filename
        if not pdf_path.exists():
            log.error("Process file not found: %s", pdf_path)
            return jsonify({"error": "Uploaded file not found. Please re-upload."}), 404

        # Generate unique request ID (keep for logging/audit, but index status by filename)
        request_id = str(uuid.uuid4())[:12]
        log.info("Initializing status store for file: %s (request ID: %s)", filename, request_id)

        # Initialize status in processing_status dictionary using filename
        processing_status[filename] = {
            "status": "processing",
            "message": "Initializing...",
            "progress": 5,
            "filename": filename,
            "bank_name": bank_name,
            "timestamp": datetime.now().isoformat(),
            "total_rows": 0,
            "new_rows": 0,
            "duplicates_skipped": 0,
        }

        # Start background thread
        log.info("Spawning background thread to process file: %s", filename)
        thread = threading.Thread(
            target=run_pipeline_in_thread,
            args=(filename, pdf_path, password, bank_name),
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
def check_status(filename: str):
    """Check processing status by filename."""
    status = processing_status.get(filename)

    if not status:
        return jsonify({"status": "unknown", "message": "Status not found."})

    return jsonify(status)


@app.route("/success/<filename>", methods=["GET"])
def success_page(filename: str):
    """Success page after processing."""
    status = processing_status.get(filename)

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
def error_page(filename: str):
    """Error page after processing failure."""
    status = processing_status.get(filename)

    if not status:
        return redirect(url_for("index"))

    data = {
        "filename": status.get("filename", filename),
        "bank_name": status.get("bank_name", "YES BANK"),
        "timestamp": status.get("timestamp"),
    }

    return render_template("error.html", error=status.get("error", "Unknown error"), data=data)


@app.route("/history", methods=["GET"])
def history():
    """Display processing history."""
    try:
        config = load_config()
        history_path = SCRIPT_DIR / config.get("processing", {}).get(
            "history_file", "logs/processing_history.json"
        )

        history_entries = []
        if history_path.exists():
            with open(history_path, "r", encoding="utf-8") as f:
                history_entries = json.load(f)

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


@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint."""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "version": "1.0.0",
    })


@app.route("/check_emails", methods=["POST"])
def check_emails():
    """Trigger email checking manually."""
    try:
        result = subprocess.run(
            [sys.executable, str(SCRIPT_DIR / "email_reader.py")],
            capture_output=True, text=True, timeout=120
        )
        cleanup_directories()
        
        out = result.stdout + result.stderr
        if "No unread emails found" in out or "No unread emails" in out:
            return jsonify({"status": "no_emails", "message": "No unread emails found"})
        elif "Pipeline completed" in out or "Google Sheets upload success" in out:
            return jsonify({"status": "success", "message": "Successfully processed emails"})
        elif "Pipeline failed" in out or "Unlock failed" in out:
            return jsonify({"status": "failed", "message": "Failed to process some emails"})
        else:
            if result.returncode == 0:
                return jsonify({"status": "success", "message": "Email check completed"})
            return jsonify({"status": "failed", "message": "Error running email check"})
    except subprocess.TimeoutExpired:
        return jsonify({"status": "failed", "message": "Email check timed out."}), 504
    except Exception as e:
        log.exception("Error checking emails")
        return jsonify({"status": "failed", "error": str(e)}), 500


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
