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
import subprocess
import sys
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
LOG_PATH = SCRIPT_DIR / "logs" / "app.log"
HISTORY_PATH = SCRIPT_DIR / "logs" / "processing_history.json"
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
# Subprocess runner
# ---------------------------------------------------------------------------
def run_command(cmd: list[str], logger: logging.Logger) -> tuple[int, str, str]:
    """Run a subprocess command.

    Returns:
        Tuple of (exit_code, stdout, stderr).
    """
    logger.info("Running: %s", " ".join(cmd))
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.stdout:
        for line in result.stdout.strip().split("\n"):
            if line.strip():
                logger.info("[child] %s", line)

    if result.stderr:
        for line in result.stderr.strip().split("\n"):
            if line.strip():
                logger.error("[child] %s", line)

    return result.returncode, result.stdout or "", result.stderr or ""


# ---------------------------------------------------------------------------
# Pipeline steps
# ---------------------------------------------------------------------------
def step_unlock(
    input_pdf: Path,
    output_pdf: Path,
    password: str,
    logger: logging.Logger,
) -> bool:
    """Step 1: Decrypt the PDF."""
    logger.info("--- Step 1: Unlocking PDF ---")
    cmd = [
        sys.executable, str(SCRIPT_DIR / "unlock_pdf.py"),
        "--input", str(input_pdf),
        "--output", str(output_pdf),
        "--password", password,
    ]
    rc, _, _ = run_command(cmd, logger)
    return rc == 0


def step_extract(
    unlocked_pdf: Path,

    excel_file: Path,
    logger: logging.Logger,
) -> bool:
    """Step 2: Extract transactions from PDF to Excel."""
    logger.info("--- Step 2: Extracting statement ---")
    cmd = [
        sys.executable, str(SCRIPT_DIR / "extract_statement.py"),
        "--input", str(unlocked_pdf),
        "--output", str(excel_file),
    ]
    rc, stdout, _ = run_command(cmd, logger)
    if rc == 0:
        # Forward child stdout so parent (web_app) can parse log lines from it
        if stdout:
            sys.stdout.write(stdout)
            sys.stdout.flush()
        return True
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
    cmd = [
        sys.executable, str(SCRIPT_DIR / "upload_to_sheets.py"),
        "--input", str(excel_file),
        "--credentials", str(credentials_path),
        "--sheet-title", sheet_title,
        "--source-pdf", source_pdf_name,
    ]
    rc, stdout, _ = run_command(cmd, logger)

    metrics: dict = {"total_rows": 0, "new_rows": 0, "duplicates_skipped": 0, "sheet_url": ""}

    if rc == 0:
        # Parse the last JSON line from stdout
        for line in reversed(stdout.strip().split("\n")):
            line = line.strip()
            if line.startswith("{"):
                try:
                    metrics = json.loads(line)
                    logger.info("Metrics parsed: %s", metrics)
                except json.JSONDecodeError:
                    logger.warning("Could not parse metrics JSON: %s", line)
                break
        return True, metrics.get("sheet_url", ""), metrics

    return False, "", metrics


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
    """Load processing history from JSON log file."""
    if not history_path.exists():
        return []
    try:
        with open(history_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning("Could not load history file: %s", exc)
        return []


def save_history_entry(
    history_path: Path,
    entry: dict[str, Any],
    logger: logging.Logger,
) -> None:
    """Append a new entry to the processing history JSON file."""
    history = load_history(history_path, logger)
    history.append(entry)

    if len(history) > 500:
        history = history[-500:]

    history_path.parent.mkdir(parents=True, exist_ok=True)
    with open(history_path, "w", encoding="utf-8") as f:
        json.dump(history, f, indent=2, default=str, ensure_ascii=False)

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
    processed_dir = SCRIPT_DIR / folders.get("processed", "processed")
    failed_dir = SCRIPT_DIR / folders.get("failed", "failed")

    # Unique output paths — never overwrite between runs
    output_dir = SCRIPT_DIR / "output"
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

    # ── Success ─────────────────────────────────────────────────────────────
    result.update({
        "status": "success",
        "total_rows": metrics.get("total_rows", 0),
        "new_rows": metrics.get("new_rows", 0),
        "duplicates_skipped": metrics.get("duplicates_skipped", 0),
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
