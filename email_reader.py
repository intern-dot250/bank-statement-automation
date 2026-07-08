import base64
import json
import logging
import os
import shutil
import sys
import re
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from runtime_paths import base_data_dir, is_serverless
import credentials_store
import history_store
from unlock_pdf import decrypt_pdf
from run_pipeline import (
    run_pipeline as _run_pipeline_fn,
    load_config as _load_pipeline_config,
    CONFIG_PATH as _PIPELINE_CONFIG_PATH,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = base_data_dir(SCRIPT_DIR)
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
CREDENTIALS_FILE = SCRIPT_DIR / "gmail_credentials.json"
TOKEN_FILE = SCRIPT_DIR / "token.json"

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
    records_path = SCRIPT_DIR / "records.json"
    return credentials_store.list_credentials(records_path)

# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------
def authenticate_gmail():
    """Authenticate with Gmail and return an authorized API client.

    Token resolution order: local token.json file, then the
    GOOGLE_TOKEN_JSON environment variable (needed on a serverless
    deployment where a local token file can't be read/written). If the
    token is missing/invalid and can be refreshed (has a refresh_token),
    that happens with no browser interaction. Only if there's no usable
    token at all does this fall back to the interactive OAuth consent
    flow (gmail_credentials.json / GMAIL_CREDENTIALS_JSON) — which opens
    a local browser and therefore cannot run in a serverless request;
    that case raises a clear error there instead of hanging/crashing.
    """
    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    else:
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
        token_path = DATA_DIR / "token.json"
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
) -> tuple[bool, str | None]:
    """Run the full pipeline for one PDF, in-process (no subprocess —
    unreliable/unsupported on serverless deployments such as Vercel, and
    avoids the process-startup overhead a subprocess pays every call).

    Returns:
        Tuple of (success, request_id).
    """
    try:
        config = _load_pipeline_config(_PIPELINE_CONFIG_PATH)
    except Exception as exc:
        logger.error("[pipeline] Could not load config.json: %s", exc)
        return False, None

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
        return False, None

    if success:
        return True, result.get("request_id")

    logger.error(
        "[pipeline] %s",
        result.get("error") or "Pipeline failed (no error message captured).",
    )
    return False, None

# ---------------------------------------------------------------------------
# Main Logic
# ---------------------------------------------------------------------------
def log_failure_to_history(filename: str, stage: int, error_msg: str):
    history_file = LOG_DIR / "processing_history.json"
    import uuid
    from datetime import datetime
    entry = {
        "timestamp": datetime.now().isoformat(),
        "file": filename or "Unknown",
        "bank": "Unknown",
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
    records_path = DATA_DIR / "records.json"
    history_store.save_latest_batch(batch_stats, records_path)

def process_emails() -> dict:
    """Check unread Gmail messages for bank statement PDFs and process them.

    Returns:
        The batch_stats dict ({"processed", "success", "failed"}) for this
        run — callers that import this function directly (e.g. web_app.py)
        can use the return value instead of parsing subprocess stdout text.
    """
    accounts_config = load_accounts()

    batch_stats = {
        "processed": 0,
        "success": 0,
        "failed": 0
    }

    logger.info("Authenticating with Gmail...")
    service = authenticate_gmail()

    logger.info("[STAGE 1 START] Fetching unread emails...")
    try:
        query = "is:unread has:attachment filename:pdf"
        results = service.users().messages().list(userId='me', q=query).execute()
        messages = results.get('messages', [])
        logger.info("[STAGE 1 SUCCESS] Unread emails fetched")
    except Exception as e:
        logger.error("[STAGE 1 FAILED] Error fetching emails: %s", e)
        log_failure_to_history("Unknown", 1, f"Email fetch failed: {e}")
        return batch_stats

    if not messages:
        logger.info("No unread emails with PDF attachments found.")
        save_latest_batch(batch_stats)
        return batch_stats

    for msg in messages:
        msg_id = msg['id']
        message = service.users().messages().get(userId='me', id=msg_id).execute()
        
        logger.info("[STAGE 2 START] Parsing email body")
        try:
            payload = message['payload']
            body = get_email_body(payload)
            logger.info("[STAGE 2 SUCCESS] Email body parsed")
        except Exception as e:
            logger.error("[STAGE 2 FAILED] Could not parse email body: %s", e)
            log_failure_to_history("Unknown", 2, "Email body parse failed")
            continue
        
        logger.info("[STAGE 3 START] Extracting account number")
        last_4_digits = extract_last_4_digits(body)
        
        password = None
        matched_account = None
        matched_bank_name = "YES BANK"

        if last_4_digits:
            logger.info("[STAGE 3 SUCCESS] Last 4 digits extracted: %s", last_4_digits)
            logger.info("[STAGE 4 START] Looking up password")
            for acc in accounts_config:
                if acc.get("account_number", "").endswith(last_4_digits):
                    matched_account = acc.get("account_number")
                    password = acc.get("password")
                    matched_bank_name = acc.get("bank_name") or "YES BANK"
                    break
                    
            if matched_account and password:
                logger.info("[STAGE 4 SUCCESS] Account matched: %s. Password found", matched_account)
            else:
                logger.error("[STAGE 4 FAILED] Password not found for account ending in %s", last_4_digits)
                log_failure_to_history("Unknown", 4, f"Password missing for account ending in {last_4_digits}")
                continue
        else:
            logger.error("[STAGE 3 FAILED] Account not found in body")
            log_failure_to_history("Unknown", 3, "Account number not found in email body")
            continue

        success_processing_all = True
        attachments_found = False

        for part in get_pdf_attachments(payload):
            attachments_found = True
            filename = str(part.get('filename', '')).strip()
            batch_stats["processed"] += 1

            logger.info("[STAGE 5 START] Downloading PDF: %s", filename)
            attachment_id = part['body'].get('attachmentId')
            if not attachment_id:
                logger.error("[STAGE 5 FAILED] Missing attachmentId for %s", filename)
                log_failure_to_history(filename, 5, "Missing attachment ID")
                success_processing_all = False
                batch_stats["failed"] += 1
                continue

            try:
                attachment = service.users().messages().attachments().get(
                    userId='me', messageId=msg_id, id=attachment_id).execute()
                file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))

                pdf_path = INPUT_DIR / filename
                with open(pdf_path, 'wb') as f:
                    f.write(file_data)
                logger.info("[STAGE 5 SUCCESS] PDF downloaded: %s", filename)
            except Exception as e:
                logger.error("[STAGE 5 FAILED] Could not download/save PDF %s: %s", filename, e)
                log_failure_to_history(filename, 5, "PDF download/save failed")
                success_processing_all = False
                batch_stats["failed"] += 1
                continue

            logger.info("[STAGE 6 START] PDF unlock test")
            unlocked_pdf_path = OUTPUT_DIR / f"temp_unlocked_{filename}"
            # Call decrypt_pdf() directly in-process (no subprocess —
            # unreliable/unsupported on serverless deployments such as
            # Vercel). Any failure here (wrong password, corrupted file,
            # filesystem error) is treated as an unlock-test failure,
            # matching the previous subprocess's non-zero-exit behavior.
            try:
                decrypt_pdf(pdf_path, unlocked_pdf_path, password)
                unlock_ok = unlocked_pdf_path.exists()
            except Exception as exc:
                logger.error("[STAGE 6 FAILED] Unlock test error for %s: %s", filename, exc)
                unlock_ok = False

            if not unlock_ok:
                logger.error("[STAGE 6 FAILED] Unlock test failed for %s", filename)
                log_failure_to_history(filename, 6, "Unlock test failed (incorrect password or corrupted PDF)")
                success_processing_all = False
                batch_stats["failed"] += 1
                continue

            logger.info("[STAGE 6 SUCCESS] PDF unlock test success")
            logger.info("Pipeline started")

            ok, request_id = run_pipeline_for_pdf(
                pdf_path, password,
                account_number=matched_account,
                bank_name=matched_bank_name,
            )

            if unlocked_pdf_path.exists():
                unlocked_pdf_path.unlink()

            if ok:
                logger.info("Google Sheets upload success")
                logging.info("Pipeline completed. File lifecycle handled by run_pipeline.py")
                batch_stats["success"] += 1

                if request_id:
                    unlocked_out = OUTPUT_DIR / f"unlocked_{request_id}.pdf"
                    excel_file = OUTPUT_DIR / f"bank_statement_{request_id}.xlsx"
                    try:
                        if unlocked_out.exists(): unlocked_out.unlink()
                        if excel_file.exists(): excel_file.unlink()
                    except Exception:
                        pass
            else:
                logger.error("Pipeline failed")
                success_processing_all = False
                batch_stats["failed"] += 1

        if attachments_found and success_processing_all:
            logger.info("[STAGE 11 START] Marking email as read")
            try:
                service.users().messages().modify(
                    userId='me', 
                    id=msg_id, 
                    body={'removeLabelIds': ['UNREAD']}
                ).execute()
                logger.info("[STAGE 11 SUCCESS] Email marked as read")
            except Exception as e:
                logger.error("[STAGE 11 FAILED] Failed to mark email as read: %s", e)

    save_latest_batch(batch_stats)
    return batch_stats

if __name__ == "__main__":
    process_emails()
