import base64
import json
import logging
import os
import shutil
import sys
import re
from pathlib import Path
from typing import Callable, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from runtime_paths import base_data_dir, is_serverless
import credentials_store
import gmail_accounts_store
import history_store
from unlock_pdf import decrypt_pdf
from run_pipeline import (
    run_pipeline as _run_pipeline_fn,
    load_config as _load_pipeline_config,
    categorize_pipeline_failure,
    CONFIG_PATH as _PIPELINE_CONFIG_PATH,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = base_data_dir(SCRIPT_DIR)
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
CREDENTIALS_FILE = SCRIPT_DIR / "config" / "gmail_credentials.json"
TOKEN_FILE = SCRIPT_DIR / "config" / "token.json"

# Env var fallbacks for serverless deployments, where gmail_credentials.json
# and token.json can't be committed to the repo or read from a local file.
GOOGLE_TOKEN_ENV_VAR = "GOOGLE_TOKEN_JSON"
GMAIL_CREDENTIALS_ENV_VAR = "GMAIL_CREDENTIALS_JSON"
INPUT_DIR = DATA_DIR / "input"
FAILED_DIR = DATA_DIR / "failed"
OUTPUT_DIR = DATA_DIR / "output"
PROCESSED_DIR = DATA_DIR / "processed"
LOG_DIR = DATA_DIR / "logs"

# Directory creation and file logging below can still fail even at the
# (possibly serverless-redirected) DATA_DIR in unexpected environments.
# Falling back gracefully here prevents the whole module from crashing on
# import in that case.
for _dir in (LOG_DIR, INPUT_DIR, FAILED_DIR, OUTPUT_DIR, PROCESSED_DIR):
    try:
        _dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass

# Set up logging
log_file = LOG_DIR / "email_reader.log"
_log_handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
try:
    _log_handlers.append(logging.FileHandler(str(log_file), encoding="utf-8"))
except OSError:
    pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    handlers=_log_handlers,
)
logger = logging.getLogger("email_reader")

# ---------------------------------------------------------------------------
# Config Loading
# ---------------------------------------------------------------------------
def load_accounts():
    """Load configured bank account -> password mappings.

    Reads from the account_credentials table (Postgres via DATABASE_URL)
    when configured, falling back to records.json's "accounts" list
    otherwise — same pattern as history_store.py.
    """
    records_path = SCRIPT_DIR / "data" / "records.json"
    return credentials_store.list_credentials(records_path)

# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------
def authenticate_gmail(token_json: Optional[str] = None, account_id: Optional[int] = None):
    """Authenticate with Gmail and return an authorized API client.

    Token resolution order: the active Gmail account connected via the
    Admin page (gmail_accounts_store — lets the accounts team switch which
    inbox "Check Bank Emails" reads from without a code change/redeploy),
    then (unchanged, for backwards compatibility during transition before
    any account has been connected this way) local token.json file, then
    the GOOGLE_TOKEN_JSON environment variable (needed on a serverless
    deployment where a local token file can't be read/written). If the
    token is missing/invalid and can be refreshed (has a refresh_token),
    that happens with no browser interaction — and if it came from the
    active connected account, the refreshed token is written back to
    gmail_accounts_store too, since Vercel's local filesystem is ephemeral.
    Only if there's no usable token at all does this fall back to the
    interactive OAuth consent flow (gmail_credentials.json /
    GMAIL_CREDENTIALS_JSON) — which opens a local browser and therefore
    cannot run in a serverless request; that case raises a clear error
    there instead of hanging/crashing.
    """
    creds = None

    # 1) Explicit per-account token (preferred when called from
    #    process_emails() looping over active accounts).
    if token_json:
        try:
            info = json.loads(token_json)
            creds = Credentials.from_authorized_user_info(info, SCOPES)
        except (json.JSONDecodeError, ValueError) as exc:
            logger.error("Explicit Gmail token is not valid: %s", exc)

    # 2) Active connected account from gmail_accounts_store (only if no
    #    explicit token was given, to avoid silently reading the wrong inbox).
    if not creds:
        active_token_json = gmail_accounts_store.get_active_token()
        if active_token_json:
            try:
                info = json.loads(active_token_json)
                creds = Credentials.from_authorized_user_info(info, SCOPES)
                account_id = gmail_accounts_store.get_active_account_id()
            except (json.JSONDecodeError, ValueError) as exc:
                logger.error("Active Gmail account's stored token is not valid: %s", exc)

    if not creds and TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    elif not creds:
        token_env_value = os.environ.get(GOOGLE_TOKEN_ENV_VAR)
        if token_env_value:
            try:
                info = json.loads(token_env_value)
                creds = Credentials.from_authorized_user_info(info, SCOPES)
            except (json.JSONDecodeError, ValueError) as exc:
                logger.error(
                    "%s environment variable is not valid: %s", GOOGLE_TOKEN_ENV_VAR, exc,
                )

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            if account_id is not None:
                gmail_accounts_store.update_token(account_id, creds.to_json())
        elif is_serverless():
            logger.error(
                "No valid Gmail token available, and the interactive OAuth "
                "consent flow cannot run in this environment."
            )
            raise RuntimeError(
                f"Gmail authentication unavailable: no valid token, and the "
                f"interactive consent flow requires a local browser. Set "
                f"{GOOGLE_TOKEN_ENV_VAR} to a valid, refreshable token JSON "
                "(generate it once locally first)."
            )
        else:
            credentials_file = CREDENTIALS_FILE
            if not credentials_file.exists():
                creds_env_value = os.environ.get(GMAIL_CREDENTIALS_ENV_VAR)
                if creds_env_value:
                    try:
                        client_config = json.loads(creds_env_value)
                        flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
                        creds = flow.run_local_server(port=5000)
                    except (json.JSONDecodeError, ValueError) as exc:
                        logger.error(
                            "%s environment variable is not valid: %s",
                            GMAIL_CREDENTIALS_ENV_VAR, exc,
                        )
                        raise
                else:
                    logger.error("Gmail credentials file not found: %s", credentials_file)
                    raise FileNotFoundError(
                        f"Gmail credentials file not found: {credentials_file}"
                    )
            else:
                flow = InstalledAppFlow.from_client_secrets_file(str(credentials_file), SCOPES)
                creds = flow.run_local_server(port=5000)

        # Persist the refreshed/obtained token for reuse. Best-effort only —
        # a failure here (e.g. read-only filesystem) must not crash auth.
        token_path = DATA_DIR / "config" / "token.json"
        try:
            token_path.parent.mkdir(parents=True, exist_ok=True)
            with open(token_path, 'w') as token:
                token.write(creds.to_json())
        except OSError as exc:
            logger.warning("Could not persist refreshed Gmail token: %s", exc)

    return build('gmail', 'v1', credentials=creds)

# ---------------------------------------------------------------------------
# Email Body Parsing
# ---------------------------------------------------------------------------
def get_email_body(payload: dict) -> str:
    body_data = ""
    
    if 'parts' in payload:
        for part in payload['parts']:
            if part['mimeType'] == 'text/plain':
                body_data = part.get('body', {}).get('data', '')
                break
            elif part['mimeType'] == 'text/html':
                body_data = part.get('body', {}).get('data', '')
            elif 'parts' in part:
                res = get_email_body(part)
                if res:
                    return res
    else:
        body_data = payload.get('body', {}).get('data', '')
        
    if body_data:
        try:
            return base64.urlsafe_b64decode(body_data).decode('utf-8', errors='ignore')
        except Exception:
            pass
            
    return ""

def extract_last_4_digits(body: str) -> str | None:
    pattern = r"(?i)(?:account\s*no\.?|a/c\s*no\.?|account\s*number)\s*[:\-]?\s*[X*A-Z\d]*(\d{4})\b"
    match = re.search(pattern, body)
    if match:
        return match.group(1)
    
    pattern_fallback = r"(?i)(?:a/c|account).{0,15}[X*A-Z\d]*(\d{4})\b"
    match = re.search(pattern_fallback, body)
    if match:
        return match.group(1)
        
    return None

def get_pdf_attachments(payload: dict):
    parts = payload.get('parts', [payload])
    for part in parts:
        if part.get('filename') and part['filename'].lower().endswith('.pdf'):
            yield part
        elif 'parts' in part:
            yield from get_pdf_attachments(part)

# ---------------------------------------------------------------------------
# Pipeline Execution
# ---------------------------------------------------------------------------
def run_pipeline_for_pdf(
    pdf_path: Path,
    password: str,
    account_number: str = "",
    bank_name: str = "YES BANK",
) -> tuple[bool, dict]:
    """Run the full pipeline for one PDF, in-process (no subprocess —
    unreliable/unsupported on serverless deployments such as Vercel, and
    avoids the process-startup overhead a subprocess pays every call).

    Returns:
        Tuple of (success, result). `result` is run_pipeline()'s own result
        dict on a normal success/failure, or a synthetic
        {"error": ..., "failed_stage": None} dict if config-loading or the
        pipeline call itself raised — always populated so the caller can
        pass it straight to categorize_pipeline_failure().
    """
    try:
        config = _load_pipeline_config(_PIPELINE_CONFIG_PATH)
    except Exception as exc:
        logger.error("[pipeline] Could not load config.json: %s", exc)
        return False, {"error": str(exc), "failed_stage": None}

    try:
        success, result = _run_pipeline_fn(
            password=password,
            input_pdf=pdf_path,
            config=config,
            bank_name=bank_name,
            account_number=account_number,
            logger=logger,
        )
    except Exception as exc:
        logger.error("[pipeline] %s", exc)
        return False, {"error": str(exc), "failed_stage": None}

    if success:
        return True, result

    logger.error(
        "[pipeline] %s",
        result.get("error") or "Pipeline failed (no error message captured).",
    )
    return False, result

# ---------------------------------------------------------------------------
# Main Logic
# ---------------------------------------------------------------------------
def log_failure_to_history(filename: str, stage: int, error_msg: str, account_number: str = ""):
    history_file = LOG_DIR / "processing_history.json"
    import uuid
    from datetime import datetime
    entry = {
        "timestamp": datetime.now().isoformat(),
        "file": filename or "Unknown",
        "bank": "Unknown",
        "account_number": account_number,
        "request_id": f"email_{uuid.uuid4().hex[:8]}",
        "status": "failed",
        "total_rows": 0,
        "new_rows": 0,
        "duplicates_skipped": 0,
        "sheet_url": "",
        "error": error_msg,
        "failed_stage": stage,
        "source": "Email"
    }
    try:
        history_store.save_history_entry(entry, history_file)
    except Exception as e:
        logger.error("Could not write failure to history: %s", e)


# An email whose PDF keeps failing at the same stage never gets marked read
# (see the mark-as-read check at the end of the main loop below), so it's
# picked up again by every future "Check Bank Emails" run — retrying a PDF
# that can't succeed is pointless and clutters Processing History with the
# same failure repeated indefinitely. This threshold stops that: once a
# filename has failed this many times before, further attempts are skipped
# (logged as a distinct "skipped" entry) and the email is allowed to be
# marked read, so it stops resurfacing. It still needs a human to actually
# fix/reprocess it (e.g. via Manual Upload) — this just stops the auto-loop.
_REPEATED_FAILURE_SKIP_THRESHOLD = 2


def _count_prior_extraction_failures(filename: str) -> int:
    """Count how many times this exact filename has already failed
    processing (any stage), per Processing History. Best-effort — returns
    0 (never skip) if history can't be read for any reason."""
    if not filename:
        return 0
    try:
        history_file = LOG_DIR / "processing_history.json"
        entries = history_store.load_history(history_file)
        return sum(
            1 for e in entries
            if e.get("file") == filename and e.get("status") == "failed"
        )
    except Exception as exc:
        logger.warning("Could not check prior failure count for %s: %s", filename, exc)
        return 0


def save_latest_batch(batch_stats: dict) -> None:
    """Save current-batch PDF processing counts into records.json.

    Adds/updates the "latest_batch" key only. All other keys already
    present in records.json (e.g. "accounts") and all history files are
    left untouched.

    Written to DATA_DIR (not necessarily SCRIPT_DIR) since this file must
    be writable at runtime — on a normal deployment DATA_DIR == SCRIPT_DIR
    so this is the same file as always; on a read-only serverless
    deployment it's redirected to a writable location. load_accounts()
    intentionally still reads the bundled, read-only SCRIPT_DIR copy for
    the "accounts" list, since that's config shipped with the code, not
    runtime state.
    """
    records_path = DATA_DIR / "data" / "records.json"
    history_store.save_latest_batch(batch_stats, records_path)

def _mask_account_number(account_number: str | None) -> str | None:
    """"045563400002477" -> "XXXX2477", for display only."""
    if not account_number or len(account_number) < 4:
        return None
    return "XXXX" + account_number[-4:]


def _now_time_str() -> str:
    from datetime import datetime
    return datetime.now().strftime("%I:%M %p")


# Same abbreviations shown on Admin -> Account Passwords' "Account Type"
# column (see web_app.py's _PROJECT_ABBREVIATIONS) — kept as a small local
# copy rather than a shared import to avoid a circular import (web_app.py
# already imports this module for process_emails()).
_PROJECT_ABBREVIATIONS = {"Casa Romana": "CR", "Aravali Heights": "AH"}


def _abbreviate_project(text: str | None) -> str:
    """"Casa Romana" -> "CR"; anything not in the table falls back to
    initials (e.g. "Some New Project" -> "SNP")."""
    if not text:
        return ""
    return _PROJECT_ABBREVIATIONS.get(text) or "".join(w[0] for w in text.split()).upper()


def _abbreviate_bank(bank_name: str | None) -> str:
    """"YES BANK" -> "YES"; "Bank of Maharashtra" -> "BOM" (first word is
    the generic word "Bank", so use initials instead)."""
    if not bank_name:
        return ""
    words = bank_name.split()
    if not words:
        return ""
    if words[0].lower() == "bank":
        return "".join(w[0] for w in words).upper()
    return words[0].upper()


def process_emails(
    on_progress: Optional[Callable[[str, int], None]] = None,
    on_pdf_update: Optional[Callable[[list[dict], dict], None]] = None,
) -> tuple[dict, list[dict]]:
    """Check unread Gmail messages for bank statement PDFs and process them.

    Iterates over all active Gmail accounts (is_active = TRUE in the
    gmail_accounts table). Each account is authenticated independently;
    results from all accounts are aggregated into a single batch_stats
    and processed_pdfs list.

    Args:
        on_progress: Optional callback invoked as on_progress(message, percent)
            at each stage, letting a caller (e.g. web_app.py) surface a live
            progress bar. Errors raised by the callback itself are swallowed —
            a progress-reporting failure must never abort email processing.
        on_pdf_update: Optional callback invoked as
            on_pdf_update(processed_pdfs, batch_stats) every time
            processed_pdfs changes (a new PDF starts processing, or one
            resolves to its final status) — lets a caller persist live,
            per-PDF progress instead of only seeing results once the whole
            batch finishes. Same swallow-errors contract as on_progress.

    Returns:
        (batch_stats, processed_pdfs) — batch_stats is {"processed",
        "success", "failed", "total_emails"} for this run; processed_pdfs
        is a list of one dict per PDF actually attempted ({"index", "label",
        "filename", "bank", "account_number", "date", "status", "message",
        "time"}), purely for a live "what did this run just do" view on
        the Dashboard — not written to any persistent history store.
    """
    def _report(message: str, percent: int) -> None:
        if on_progress is None:
            return
        try:
            on_progress(message, percent)
        except Exception:
            logger.debug("on_progress callback raised — ignoring", exc_info=True)

    def _notify_pdf_update() -> None:
        if on_pdf_update is None:
            return
        try:
            on_pdf_update(processed_pdfs, batch_stats)
        except Exception:
            logger.debug("on_pdf_update callback raised — ignoring", exc_info=True)

    accounts_config = load_accounts()

    batch_stats = {
        "processed": 0,
        "success": 0,
        "failed": 0,
        "total_emails": 0,
    }
    processed_pdfs: list[dict] = []

    # --- NEW: fetch all active Gmail accounts ---
    active_accounts = gmail_accounts_store.list_active_tokens()

    if not active_accounts:
        logger.info("No active Gmail accounts configured.")
        _report("No active Gmail accounts configured", 100)
        save_latest_batch(batch_stats)
        return batch_stats, processed_pdfs

    num_accounts = len(active_accounts)
    logger.info("Processing emails for %d active Gmail account(s)", num_accounts)

    # Track which emails have been marked as read across all accounts
    # so we can do the batch modify at the end (mirrors original behavior).
    emails_to_mark_read: list[tuple] = []  # (service, msg_id)

    for acc_idx, gmail_acc in enumerate(active_accounts):
        account_email = gmail_acc.get("email", f"account-{gmail_acc['id']}")
        account_label = f"[{account_email}]"

        # Progress: split 5–95% range evenly across accounts
        account_start_pct = 5 + int(85 * acc_idx / num_accounts)
        account_end_pct = 5 + int(85 * (acc_idx + 1) / num_accounts)

        logger.info("Authenticating Gmail account: %s", account_email)
        _report(f"Authenticating {account_email}...", account_start_pct)

        try:
            service = authenticate_gmail(
                token_json=gmail_acc.get("token_json"),
                account_id=gmail_acc.get("id"),
            )
        except Exception as exc:
            logger.error("Failed to authenticate %s: %s", account_email, exc)
            _report(f"Auth failed for {account_email}: {exc}", account_end_pct)
            continue

        # Stage 1: Fetch unread emails for this account
        logger.info("[%s] Fetching unread emails...", account_label)
        _report(f"Fetching emails from {account_email}...", account_start_pct + 3)
        try:
            query = "is:unread has:attachment filename:pdf"
            results = service.users().messages().list(userId='me', q=query).execute()
            messages = results.get('messages', [])
            logger.info("[%s] Unread emails fetched: %d", account_label, len(messages))
        except Exception as e:
            logger.error("[%s] Error fetching emails: %s", account_label, e)
            log_failure_to_history("Unknown", 1, f"Email fetch failed ({account_email}): {e}")
            continue

        if not messages:
            logger.info("[%s] No unread emails with PDF attachments found.", account_label)
            _report(f"No emails in {account_email}", account_end_pct)
            continue

        total_messages = len(messages)
        batch_stats["total_emails"] += total_messages
        _notify_pdf_update()
        _report(f"[{account_email}] Found {total_messages} email(s) with PDFs", account_start_pct + 7)

        success_processing_all = True

        for msg_index, msg in enumerate(messages):
            msg_progress = account_start_pct + 10 + int(
                (account_end_pct - account_start_pct - 15) * msg_index / max(total_messages, 1)
            )
            msg_id = msg['id']
            message = service.users().messages().get(userId='me', id=msg_id).execute()

            logger.info("[%s][STAGE 2 START] Parsing email body", account_label)
            try:
                payload = message['payload']
                body = get_email_body(payload)
                logger.info("[%s][STAGE 2 SUCCESS] Email body parsed", account_label)
            except Exception as e:
                logger.error("[%s][STAGE 2 FAILED] Could not parse email body: %s", account_label, e)
                log_failure_to_history("Unknown", 2, "Email body parse failed")
                continue

            email_date = next(
                (h.get("value") for h in payload.get("headers", []) if h.get("name") == "Date"),
                "",
            )

            logger.info("[%s][STAGE 3 START] Extracting account number", account_label)
            last_4_digits = extract_last_4_digits(body)

            password = None
            matched_account = None
            matched_bank_name = "YES BANK"
            matched_acc_record: dict = {}

            if last_4_digits:
                logger.info("[%s][STAGE 3 SUCCESS] Last 4 digits extracted: %s", account_label, last_4_digits)
                logger.info("[%s][STAGE 4 START] Looking up password", account_label)
                for acc in accounts_config:
                    if acc.get("account_number", "").endswith(last_4_digits):
                        matched_account = acc.get("account_number")
                        password = acc.get("password")
                        matched_bank_name = acc.get("bank_name") or "YES BANK"
                        matched_acc_record = acc
                        break

                if matched_account and password:
                    logger.info("[%s][STAGE 4 SUCCESS] Account matched: %s. Password found", account_label, matched_account)
                else:
                    logger.error("[%s][STAGE 4 FAILED] Password not found for account ending in %s", account_label, last_4_digits)
                    log_failure_to_history("Unknown", 4, f"Password missing for account ending in {last_4_digits}")
                    continue
            else:
                logger.warning("[%s][STAGE 3 SKIPPED] Account number not found in email body — skipping.", account_label)
                continue

            # Live, this-run-only label for the Dashboard's processed-PDFs —
            # abbreviated (bank -> first word, project -> initials) so it
            # fits on one line, e.g. "YES-DPL-CR-Free-2477".
            pdf_label = "-".join(
                part for part in (
                    _abbreviate_bank(matched_acc_record.get("bank_name")),
                    matched_acc_record.get("company"),
                    _abbreviate_project(matched_acc_record.get("business_unit")),
                    matched_acc_record.get("account_stage"),
                ) if part
            ) or _abbreviate_bank(matched_bank_name) or matched_bank_name

            acc_num = matched_acc_record.get("account_number", "")
            if acc_num and len(acc_num) >= 4:
                pdf_label += "-" + acc_num[-4:]

            attachments_found = False

            for part in get_pdf_attachments(payload):
                attachments_found = True
                filename = str(part.get('filename', '')).strip()
                batch_stats["processed"] += 1

                pdf_entry = {
                    "index": len(processed_pdfs) + 1,
                    "label": pdf_label,
                    "filename": filename,
                    "bank": matched_bank_name,
                    "account_number": _mask_account_number(matched_account),
                    "date": email_date,
                    "status": "processing",
                    "message": "Processing...",
                    "time": _now_time_str(),
                }
                processed_pdfs.append(pdf_entry)
                _notify_pdf_update()

                def _resolve(status: str, message: str) -> None:
                    pdf_entry.update(status=status, message=message, time=_now_time_str())
                    _notify_pdf_update()

                prior_failures = _count_prior_extraction_failures(filename)
                if prior_failures >= _REPEATED_FAILURE_SKIP_THRESHOLD:
                    logger.error(
                        "[%s][SKIPPED] %s has already failed %d time(s) — skipping.",
                        account_label, filename, prior_failures,
                    )
                    log_failure_to_history(
                        filename, 7,
                        f"SKIPPED after {prior_failures} repeated failures — needs manual "
                        "review (e.g. reprocess via Manual Upload once the underlying "
                        "issue is fixed). Not auto-retried again.",
                        account_number=matched_account or "",
                    )
                    batch_stats["failed"] += 1
                    _resolve("skipped", "Skipped after repeated failures")
                    continue

                logger.info("[%s][STAGE 5 START] Downloading PDF: %s", account_label, filename)
                _report(f"[{account_email}] Downloading {filename} ({msg_index + 1}/{total_messages})...", msg_progress)
                attachment_id = part['body'].get('attachmentId')
                if not attachment_id:
                    logger.error("[%s][STAGE 5 FAILED] Missing attachmentId for %s", account_label, filename)
                    log_failure_to_history(filename, 5, "Missing attachment ID", account_number=matched_account or "")
                    success_processing_all = False
                    batch_stats["failed"] += 1
                    _resolve("failed", "Missing attachment ID")
                    continue

                try:
                    attachment = service.users().messages().attachments().get(
                        userId='me', messageId=msg_id, id=attachment_id).execute()
                    file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))

                    pdf_path = INPUT_DIR / filename
                    with open(pdf_path, 'wb') as f:
                        f.write(file_data)
                    logger.info("[%s][STAGE 5 SUCCESS] PDF downloaded: %s", account_label, filename)
                except Exception as e:
                    logger.error("[%s][STAGE 5 FAILED] Could not download/save PDF %s: %s", account_label, filename, e)
                    log_failure_to_history(filename, 5, "PDF download/save failed", account_number=matched_account or "")
                    success_processing_all = False
                    batch_stats["failed"] += 1
                    _resolve("failed", f"PDF download/save failed: {e}")
                    continue

                logger.info("[%s][STAGE 6 START] PDF unlock test", account_label)
                unlocked_pdf_path = OUTPUT_DIR / f"temp_unlocked_{filename}"
                try:
                    decrypt_pdf(pdf_path, unlocked_pdf_path, password)
                    unlock_ok = unlocked_pdf_path.exists()
                except Exception as exc:
                    logger.error("[%s][STAGE 6 FAILED] Unlock test error for %s: %s", account_label, filename, exc)
                    unlock_ok = False

                if not unlock_ok:
                    logger.error("[%s][STAGE 6 FAILED] Unlock test failed for %s", account_label, filename)
                    log_failure_to_history(filename, 6, "Unlock test failed (incorrect password or corrupted PDF)", account_number=matched_account or "")
                    success_processing_all = False
                    batch_stats["failed"] += 1
                    _resolve("failed", "Invalid PDF password or corrupted file")
                    continue

                logger.info("[%s][STAGE 6 SUCCESS] PDF unlock test success", account_label)
                logger.info("[%s] Pipeline started", account_label)
                _report(f"[{account_email}] Classifying {filename}...", min(90, msg_progress + 5))

                ok, result = run_pipeline_for_pdf(
                    pdf_path, password,
                    account_number=matched_account,
                    bank_name=matched_bank_name,
                )
                request_id = result.get("request_id")

                if unlocked_pdf_path.exists():
                    unlocked_pdf_path.unlink()

                if ok:
                    logger.info("[%s] Google Sheets upload success", account_label)
                    logging.info("[%s] Pipeline completed. File lifecycle handled by run_pipeline.py", account_label)
                    batch_stats["success"] += 1
                    _resolve("success", "Successfully uploaded")

                    if request_id:
                        unlocked_out = OUTPUT_DIR / f"unlocked_{request_id}.pdf"
                        excel_file = OUTPUT_DIR / f"bank_statement_{request_id}.xlsx"
                        try:
                            if unlocked_out.exists(): unlocked_out.unlink()
                            if excel_file.exists(): excel_file.unlink()
                        except Exception:
                            pass
                else:
                    logger.error("[%s] Pipeline failed", account_label)
                    success_processing_all = False
                    batch_stats["failed"] += 1
                    _resolve("failed", categorize_pipeline_failure(result))

            if attachments_found and success_processing_all:
                emails_to_mark_read.append((service, msg_id))

    # Mark all processed emails as read across all accounts
    for svc, msg_id in emails_to_mark_read:
        try:
            svc.users().messages().modify(
                userId='me',
                id=msg_id,
                body={'removeLabelIds': ['UNREAD']}
            ).execute()
        except Exception as e:
            logger.error("Failed to mark email %s as read: %s", msg_id, e)

    save_latest_batch(batch_stats)
    _report("Done", 100)
    return batch_stats, processed_pdfs

if __name__ == "__main__":
    process_emails()
