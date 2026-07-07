"""Production pipeline: unlock PDF → extract → upload to Google Sheets.

Each run is fully isolated:
  - Unique request_id per run (timestamp + uuid)
  - Unique output file names
  - All data appended to a single master worksheet (Bank_Statement_Master)
  - Duplicate rows are detected and skipped automatically
  - History entry written on every run

Usage:
    py run_pipeline.py --password "MySecret123" --input input/statement.pdf
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
from upload_to_sheets import upload_to_sheets
from classify_transactions import classify_transactions
from generate_summary import generate_summary
from generate_final_report import generate_final_report
from validate_report import validate_report
from runtime_paths import base_data_dir
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
        logger.error("Extraction failed: %s", exc)
        return False


def step_upload(
    excel_file: Path,
    sheet_title: str,
    credentials_path: Path,
    logger: logging.Logger,
    source_pdf_name: str = "unknown.pdf",
) -> tuple[bool, str, dict]:
    """Step 3: Append extracted data to the master Google Sheet.

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
        )
        logger.info("Metrics: %s", metrics)
        return True, metrics.get("sheet_url", ""), metrics
    except Exception as exc:
        logger.error("Upload failed: %s", exc)
        return False, "", metrics


def step_classify(
    credentials_path: Path,
    logger: logging.Logger,
) -> bool:
    """Phase 2 step: classify transactions in the master Google Sheet.

    Assigns Head + Narration to any unclassified rows. This step is
    non-critical: Phase 1 (unlock → extract → upload) is already complete
    by the time this runs, so a failure here is logged and does not affect
    the overall pipeline result.
    """
    logger.info("--- Step 4: Classifying transactions (Phase 2) ---")
    classify_transactions(credentials_path=credentials_path)
    return True


def step_generate_summary(
    credentials_path: Path,
    logger: logging.Logger,
) -> None:
    """Phase 2B step: regenerate the per-Head Summary worksheet.

    Reuses generate_summary.generate_summary() directly (no subprocess).
    Raises on failure — callers must treat this as a critical reporting
    stage that stops the pipeline.
    """
    logger.info("Starting Summary Generation...")
    generate_summary(credentials_path=credentials_path)
    logger.info("Summary Generation Completed.")


def step_generate_final_report(
    credentials_path: Path,
    logger: logging.Logger,
) -> None:
    """Phase 2C step: regenerate the Final Report worksheet from Summary.

    Reuses generate_final_report.generate_final_report() directly (no
    subprocess). Raises on failure — callers must treat this as a
    critical reporting stage that stops the pipeline.
    """
    logger.info("Starting Final Report Generation...")
    generate_final_report(credentials_path=credentials_path)
    logger.info("Final Report Generation Completed.")


def step_validate_report(
    credentials_path: Path,
    logger: logging.Logger,
) -> bool:
    """Phase 2D step: validate Master -> Summary -> Final Report consistency.

    Reuses validate_report.validate_report() directly (no subprocess).

    Returns:
        True if every validation check passed, False otherwise. Does not
        raise on a validation failure (that is an expected outcome, not
        an error) — callers must check the return value and stop the
        pipeline if it is False.
    """
    logger.info("Starting Validation...")
    passed = validate_report(credentials_path=credentials_path)
    if passed:
        logger.info("Validation Passed.")
    else:
        logger.error("Validation Failed.")
    return passed


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
    bank_name: str = "YES BANK",
    logger: logging.Logger | None = None,
) -> tuple[bool, dict[str, Any]]:
    """Run the full isolated pipeline for one PDF.

    Returns:
        Tuple of (success, result_dict).
        result_dict always contains: total_rows, new_rows, duplicates_skipped, sheet_url.
    """
    if logger is None:
        logger = logging.getLogger("pipeline")

    # Unique request ID for this run
    request_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"

    result: dict[str, Any] = {
        "timestamp": datetime.now().isoformat(),
        "file": input_pdf.name,
        "bank": bank_name,
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
    sheet_title = "Bank_Statement_Master"

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

    def _fail_reporting(step_name: str, exc: Exception, failed_stage: int) -> tuple[bool, dict]:
        """Fail path for the reporting stages (Summary/Final Report/Validation).

        By this point PDF extraction and the Google Sheet upload already
        succeeded, so — unlike _fail() above — the original PDF is moved
        to processed_dir (its data is genuinely in the sheet), not
        failed_dir. Only the overall pipeline result/exit code reflects
        the reporting failure, per the transactional requirement that a
        reporting-stage failure must stop the pipeline and return non-zero.
        """
        logger.error("Reporting step '%s' error: %s", step_name, exc)
        result["status"] = "failed"
        result["error"] = str(exc)
        result["failed_stage"] = failed_stage

        logger.info("[STAGE 10 START] File Cleanup")
        try:
            if input_pdf.exists():
                move_file(input_pdf, processed_dir, logger)
            logger.info("[STAGE 10 SUCCESS] File Cleanup")
        except Exception as move_exc:
            logger.error("[STAGE 10 FAILED] File Cleanup: %s", move_exc)

        logger.info("PIPELINE FAILED (reporting stage) — file moved to: %s", processed_dir)
        save_history_entry(HISTORY_PATH, result, logger)
        return False, result

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
            excel_file, sheet_title, creds_path, logger,
            source_pdf_name=input_pdf.name,
        )
        if not ok:
            raise RuntimeError("step_upload returned False (non-zero exit code).")
        logger.info("[STAGE 8 SUCCESS] Duplicate Validation")
        logger.info("[STAGE 9 SUCCESS] Google Sheets Upload")
    except Exception as exc:
        logger.error("[STAGE 9 FAILED] Upload/Validation: %s", exc)
        return _fail("Upload", exc, failed_stage=9)

    # ── Step 4: Classify (Phase 2) ──────────────────────────────────────────
    # Non-critical: Phase 1 is already successful at this point, so a
    # classification failure is logged separately and does not fail the run.
    logger.info("[STAGE 9B START] Transaction Classification")
    try:
        step_classify(creds_path, logger)
        logger.info("[STAGE 9B SUCCESS] Transaction Classification")
    except Exception as exc:
        logger.error("[STAGE 9B FAILED] Transaction Classification: %s", exc)

    # ── Step 5: Reporting pipeline (Summary → Final Report → Validation) ────
    # Runs after a successful Google Sheet upload (Step 3). Unlike Step 4
    # (Classification), each of these three stages is critical: a failure
    # here stops the pipeline immediately and the run is reported as
    # failed, per the transactional requirement for the reporting chain.
    logger.info("[STAGE 11 START] Summary Generation")
    try:
        step_generate_summary(creds_path, logger)
        logger.info("[STAGE 11 SUCCESS] Summary Generation")
    except Exception as exc:
        logger.error("[STAGE 11 FAILED] Summary Generation: %s", exc)
        return _fail_reporting("Summary Generation", exc, failed_stage=11)

    logger.info("[STAGE 12 START] Final Report Generation")
    try:
        step_generate_final_report(creds_path, logger)
        logger.info("[STAGE 12 SUCCESS] Final Report Generation")
    except Exception as exc:
        logger.error("[STAGE 12 FAILED] Final Report Generation: %s", exc)
        return _fail_reporting("Final Report Generation", exc, failed_stage=12)

    logger.info("[STAGE 13 START] Validation")
    try:
        validation_passed = step_validate_report(creds_path, logger)
    except Exception as exc:
        logger.error("[STAGE 13 FAILED] Validation: %s", exc)
        return _fail_reporting("Validation", exc, failed_stage=13)

    if not validation_passed:
        logger.error("[STAGE 13 FAILED] Validation reported failure — see validation output above.")
        return _fail_reporting(
            "Validation",
            RuntimeError("validate_report reported one or more failed checks."),
            failed_stage=13,
        )

    logger.info("[STAGE 13 SUCCESS] Validation")

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

    success, _ = run_pipeline(args.password, args.input, config)
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
