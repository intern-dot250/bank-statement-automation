"""Production pipeline: unlock PDF → extract → upload to Google Sheets.

Each run is fully isolated:
  - Unique request_id per run (timestamp + uuid)
  - Unique output file names
  - Data appended to that account's own worksheet tab (e.g. "YES BANK - 2477")
  - Duplicate rows are detected and skipped automatically
  - History entry written on every run

Usage:
    py run_pipeline.py --password "MySecret123" --input input/statement.pdf --account-number 045563400002477
"""

from __future__ import annotations

import argparse
import json
import logging
import shutil
import sys
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from unlock_pdf import decrypt_pdf
from extract_statement import extract_statement
from upload_to_sheets import upload_to_sheets, build_account_worksheet_name
from classify_transactions import classify_transactions
from runtime_paths import base_data_dir
import credentials_store
import history_store

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = base_data_dir(SCRIPT_DIR)
CONFIG_PATH = SCRIPT_DIR / "config.json"  # config.json ships with the code; read-only is fine
LOG_PATH = DATA_DIR / "logs" / "app.log"
HISTORY_PATH = DATA_DIR / "logs" / "processing_history.json"
LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
def setup_logging(log_path: Path) -> logging.Logger:
    """Set up logging to both file and console."""
    log_path.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("pipeline")
    logger.setLevel(logging.DEBUG)

    fh = logging.FileHandler(str(log_path), encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(LOG_FORMAT))

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter(LOG_FORMAT))
    try:
        ch.stream.reconfigure(encoding="utf-8")
    except Exception:
        pass

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
def load_config(config_path: Path) -> dict:
    """Load and validate config.json."""
    if not config_path.exists():
        raise FileNotFoundError(
            f"Config file not found: {config_path}\n"
            "Create a config.json file in the project root."
        )

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    required = ["credentials_path"]
    for key in required:
        if key not in config:
            raise ValueError(f"Missing required config key: '{key}'")

    return config


# ---------------------------------------------------------------------------
# Pipeline steps
# ---------------------------------------------------------------------------
# unlock/extract/upload are called directly, in-process (no subprocess) —
# spawning subprocesses is unreliable/unsupported in serverless deployments
# such as Vercel. This also removes the extra process-startup overhead the
# subprocess approach paid on every step.

def step_unlock(
    input_pdf: Path,
    output_pdf: Path,
    password: str,
    logger: logging.Logger,
) -> bool:
    """Step 1: Decrypt the PDF."""
    logger.info("--- Step 1: Unlocking PDF ---")
    try:
        decrypt_pdf(input_pdf, output_pdf, password)
        return True
    except Exception as exc:
        logger.error("Unlock failed: %s", exc)
        return False


def step_extract(
    unlocked_pdf: Path,
    excel_file: Path,
    logger: logging.Logger,
) -> bool:
    """Step 2: Extract transactions from PDF to Excel."""
    logger.info("--- Step 2: Extracting statement ---")
    try:
        extract_statement(unlocked_pdf, excel_file)
        return True
    except Exception as exc:
        logger.error("Extraction failed: %s", exc, exc_info=True)
        return False


def step_upload(
    excel_file: Path,
    credentials_path: Path,
    logger: logging.Logger,
    account_number: str,
    bank_name: str,
    source_pdf_name: str = "unknown.pdf",
) -> tuple[bool, str, dict]:
    """Step 3: Append extracted data to that account's own Google Sheet tab.

    Returns:
        Tuple of (success, sheet_url, metrics_dict).
    """
    logger.info("--- Step 3: Uploading to Google Sheets ---")
    metrics: dict = {"total_rows": 0, "new_rows": 0, "duplicates_skipped": 0, "sheet_url": ""}

    try:
        metrics = upload_to_sheets(
            input_path=excel_file,
            credentials_path=credentials_path,
            source_pdf_name=source_pdf_name,
            account_number=account_number,
            bank_name=bank_name,
        )
        logger.info("Metrics: %s", metrics)
        return True, metrics.get("sheet_url", ""), metrics
    except Exception as exc:
        logger.error("Upload failed: %s", exc)
        return False, "", metrics


def step_classify(
    credentials_path: Path,
    worksheet_name: str,
    logger: logging.Logger,
    spreadsheet=None,
) -> bool:
    """Phase 2 step: classify transactions in this account's worksheet tab."""
    logger.info("--- Step 4: Classifying transactions (Phase 2) ---")
    classify_transactions(
        credentials_path=credentials_path,
        worksheet_name=worksheet_name,
        spreadsheet=spreadsheet,
    )
    return True


def step_rag_classify(
    credentials_path: Path,
    logger: logging.Logger,
    spreadsheet=None,
) -> None:
    """Stage 9C: RAG AI classifier for any rows still showing '?' after rules.

    Uses Groq (free tier) with TF-IDF retrieval. Non-critical — a failure
    here does not stop the pipeline. Skipped silently if GROQ_API_KEY is
    not set.
    """
    import os
    if not os.environ.get("GROQ_API_KEY"):
        logger.info("[STAGE 9C SKIPPED] GROQ_API_KEY not set — RAG classifier disabled.")
        return
    try:
        from rag_classifier import run_rag_classifier
        resolved, unknown = run_rag_classifier(
            credentials_path=credentials_path,
            spreadsheet=spreadsheet,
        )
        logger.info("[STAGE 9C SUCCESS] RAG AI: %d resolved, %d still unknown.", resolved, unknown)
    except Exception as exc:
        logger.error("[STAGE 9C FAILED] RAG classifier: %s", exc)




# ---------------------------------------------------------------------------
# File routing
# ---------------------------------------------------------------------------
def move_file(src: Path, dest_dir: Path, logger: logging.Logger) -> Path:
    """Move a file to dest_dir with a timestamped suffix."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest_name = f"{src.stem}_{timestamp}{src.suffix}"
    dest_path = dest_dir / dest_name

    shutil.move(str(src), str(dest_path))
    logger.info("Moved %s → %s", src.name, dest_path)
    return dest_path


# ---------------------------------------------------------------------------
# History tracking
# ---------------------------------------------------------------------------
def load_history(history_path: Path, logger: logging.Logger) -> list[dict[str, Any]]:
    """Load processing history (Postgres via DATABASE_URL if set, else JSON file)."""
    return history_store.load_history(history_path)


def save_history_entry(
    history_path: Path,
    entry: dict[str, Any],
    logger: logging.Logger,
) -> None:
    """Save a processing-history entry (Postgres via DATABASE_URL if set, else JSON file)."""
    history_store.save_history_entry(entry, history_path)
    logger.debug("History entry saved.")


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------
def run_pipeline(
    password: str,
    input_pdf: Path,
    config: dict,
    account_number: str,
    bank_name: str = "YES BANK",
    logger: logging.Logger | None = None,
) -> tuple[bool, dict[str, Any]]:
    """Run the full isolated pipeline for one PDF.

    account_number is stamped onto every uploaded row and determines
    which account's worksheet tab (e.g. "YES BANK - 2477") the statement
    is uploaded and classified into — there is no shared master sheet.

    Returns:
        Tuple of (success, result_dict).
        result_dict always contains: total_rows, new_rows, duplicates_skipped, sheet_url.
    """
    if logger is None:
        logger = logging.getLogger("pipeline")

    # Resolve the canonical bank name for this account (from
    # account_credentials), so the same account always maps to the same
    # worksheet tab regardless of whether this run came from the manual
    # upload form (which sends a bank CODE like "YESBANK") or the email
    # flow (which sends the display name from account_credentials, e.g.
    # "YES BANK") — using whichever string the caller happened to pass
    # would otherwise split one account across two differently-named tabs.
    if account_number:
        records_path = DATA_DIR / "data" / "records.json"
        for account in credentials_store.list_credentials(records_path):
            if account.get("account_number") == account_number and account.get("bank_name"):
                if account["bank_name"] != bank_name:
                    logger.info(
                        "Using canonical bank name %r for account %s (was %r).",
                        account["bank_name"], account_number, bank_name,
                    )
                bank_name = account["bank_name"]
                break

    # Unique request ID for this run
    request_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"

    result: dict[str, Any] = {
        "timestamp": datetime.now().isoformat(),
        "file": input_pdf.name,
        "bank": bank_name,
        "account_number": account_number,
        "request_id": request_id,
        "status": "started",
        "total_rows": 0,
        "new_rows": 0,
        "duplicates_skipped": 0,
        "total_rows_in_pdf": 0,
        "sheet_url": "",
        "error": None,
        "failed_stage": None,
    }

    logger.info("=" * 60)
    logger.info("PIPELINE START — processing: %s (Bank: %s)", input_pdf.name, bank_name)
    logger.info("Request ID: %s", request_id)
    logger.info("=" * 60)

    # Resolve folders
    folders = config.get("folders", {})
    processed_dir = DATA_DIR / folders.get("processed", "processed")
    failed_dir = DATA_DIR / folders.get("failed", "failed")

    # Unique output paths — never overwrite between runs
    output_dir = DATA_DIR / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_pdf = output_dir / f"unlocked_{request_id}.pdf"
    excel_file = output_dir / f"bank_statement_{request_id}.xlsx"

    logger.info("Unlocked PDF  : %s", output_pdf)
    logger.info("Excel output  : %s", excel_file)

    creds_path = SCRIPT_DIR / config["credentials_path"]
    account_worksheet_name = build_account_worksheet_name(bank_name, account_number)

    def _fail(step_name: str, exc: Exception, failed_stage: int) -> tuple[bool, dict]:
        logger.error("Step '%s' error: %s", step_name, exc)
        result["status"] = "failed"
        result["error"] = str(exc)
        result["failed_stage"] = failed_stage
        
        logger.info("[STAGE 10 START] File Cleanup (Failed dir)")
        try:
            if input_pdf.exists():
                move_file(input_pdf, failed_dir, logger)
            logger.info("[STAGE 10 SUCCESS] File Cleanup")
        except Exception as move_exc:
            logger.error("[STAGE 10 FAILED] File Cleanup: %s", move_exc)
            
        logger.info("PIPELINE FAILED — file moved to: %s", failed_dir)
        # Persist history
        save_history_entry(HISTORY_PATH, result, logger)
        return False, result

    if not account_number:
        return _fail(
            "Account lookup",
            ValueError("No account number provided — cannot determine which account tab to use."),
            failed_stage=5,
        )

    # ── Step 1: Unlock ──────────────────────────────────────────────────────
    try:
        logger.info("[STAGE 6 START] PDF Unlock")
        ok = step_unlock(input_pdf, output_pdf, password, logger)
        if not ok:
            raise RuntimeError("step_unlock returned False (non-zero exit code).")
        logger.info("[STAGE 6 SUCCESS] PDF Unlock")
    except Exception as exc:
        logger.error("[STAGE 6 FAILED] PDF Unlock: %s", exc)
        return _fail("Unlock", exc, failed_stage=6)

    # ── Step 2: Extract ─────────────────────────────────────────────────────
    try:
        logger.info("[STAGE 7 START] Statement Extraction")
        ok = step_extract(output_pdf, excel_file, logger)
        if not ok:
            raise RuntimeError("step_extract returned False (non-zero exit code).")
        logger.info("[STAGE 7 SUCCESS] Statement Extraction")
    except Exception as exc:
        logger.error("[STAGE 7 FAILED] Statement Extraction: %s", exc)
        return _fail("Extract", exc, failed_stage=7)

    # ── Step 3: Upload ──────────────────────────────────────────────────────
    try:
        logger.info("[STAGE 8 START] Duplicate Validation")
        logger.info("[STAGE 9 START] Google Sheets Upload")
        ok, sheet_url, metrics = step_upload(
            excel_file, creds_path, logger,
            source_pdf_name=input_pdf.name,
            account_number=account_number,
            bank_name=bank_name,
        )
        if not ok:
            raise RuntimeError("step_upload returned False (non-zero exit code).")
        logger.info("[STAGE 8 SUCCESS] Duplicate Validation")
        logger.info("[STAGE 9 SUCCESS] Google Sheets Upload")
    except Exception as exc:
        logger.error("[STAGE 9 FAILED] Upload/Validation: %s", exc)
        return _fail("Upload", exc, failed_stage=9)

    # Open one shared spreadsheet for all remaining stages — avoids 4 separate
    # re-authentication round-trips (classify + summary + final report + validate).
    from upload_to_sheets import get_gspread_client, MASTER_SHEET_ID
    shared_spreadsheet = None
    try:
        shared_spreadsheet = get_gspread_client(creds_path).open_by_key(MASTER_SHEET_ID)
    except Exception as exc:
        logger.warning("Could not open shared spreadsheet: %s — stages will fall back to individual auth.", exc)

    # ── Step 4: Classify (Phase 2) ──────────────────────────────────────────
    # Non-critical: Phase 1 is already successful at this point, so a
    # classification failure is logged separately and does not fail the run.
    logger.info("[STAGE 9B START] Transaction Classification")
    try:
        step_classify(creds_path, account_worksheet_name, logger, spreadsheet=shared_spreadsheet)
        logger.info("[STAGE 9B SUCCESS] Transaction Classification")
    except Exception as exc:
        logger.error("[STAGE 9B FAILED] Transaction Classification: %s", exc)

    # ── Step 4C: RAG AI fallback (Phase 3) ─────────────────────────────────
    # Runs only if GROQ_API_KEY is set. Classifies any rows still showing '?'
    # after rule-based classification. Non-critical.
    logger.info("[STAGE 9C START] RAG AI Classification")
    step_rag_classify(creds_path, logger, spreadsheet=shared_spreadsheet)

    # ── Success ─────────────────────────────────────────────────────────────
    rows_added = metrics.get("new_rows", 0)
    duplicates_skipped = metrics.get("duplicates_skipped", 0)

    result.update({
        "status": "success",
        "total_rows": metrics.get("total_rows", 0),
        "new_rows": rows_added,
        "duplicates_skipped": duplicates_skipped,
        "total_rows_in_pdf": rows_added + duplicates_skipped,
        "sheet_url": sheet_url,
        "error": None,
    })

    # Move original locked PDF to processed/
    logger.info("[STAGE 10 START] File Cleanup")
    try:
        if input_pdf.exists():
            move_file(input_pdf, processed_dir, logger)
        logger.info("[STAGE 10 SUCCESS] File Cleanup")
    except Exception as exc:
        logger.error("[STAGE 10 FAILED] File Cleanup: %s", exc)

    # Keep output files (PDF + Excel) in output/ for debugging — do NOT delete

    logger.info("PIPELINE SUCCESS — file moved to: %s", processed_dir)
    logger.info("Sheet: %s", sheet_url)
    logger.info("Rows: total=%d  new=%d  dupes=%d",
                result["total_rows"], result["new_rows"], result["duplicates_skipped"])
    logger.info("=" * 60)

    # Persist history
    save_history_entry(HISTORY_PATH, result, logger)

    # Emit final JSON metrics line — web_app.py parses the LAST JSON line of stdout
    final_metrics = {
        "total_rows": result["total_rows"],
        "new_rows": result["new_rows"],
        "duplicates_skipped": result["duplicates_skipped"],
        "sheet_url": sheet_url,
        "request_id": request_id,
    }
    print(json.dumps(final_metrics), flush=True)

    return True, result


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Bank statement automation pipeline.",
    )
    parser.add_argument("-p", "--password", required=True,
                        help="Password for the encrypted PDF.")
    parser.add_argument("-i", "--input", type=Path,
                        default=SCRIPT_DIR / "input" / "current.pdf",
                        help="Path to the input PDF.")
    parser.add_argument("-c", "--config", type=Path, default=CONFIG_PATH,
                        help=f"Path to config.json (default: {CONFIG_PATH}).")
    parser.add_argument("-a", "--account-number", required=True,
                        help="Account number this statement belongs to.")
    parser.add_argument("-b", "--bank-name", default="YES BANK",
                        help="Bank name (used in the account tab's name).")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    logger = setup_logging(LOG_PATH)

    try:
        config = load_config(args.config)
    except (FileNotFoundError, ValueError) as exc:
        logger.error("Configuration error: %s", exc)
        return 1

    if not args.input.exists():
        logger.error("Input PDF not found: %s", args.input)
        return 1

    success, _ = run_pipeline(args.password, args.input, config, args.account_number, args.bank_name)
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
