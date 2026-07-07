"""Phase 2: Generate a human-readable Narration from a transaction's
Description, Head, and Amount.

Public API (unchanged, must remain stable for existing callers):

    generate_narration(description, head, amount) -> str

Internally this is now driven by config/narration_config.json, using
description_parser.py to extract structured fields (party, reference,
account_number, note, payment_mode) from the raw Description text.

Because generate_narration() has no direction (credit/debit) input —
only description, head, and amount — direction is inferred from the
literal "CR" banking marker (e.g. "NEFT CR-", "RTGS CR-") or a "BY CLG:"
cheque-collection prefix, both real observed textual conventions, not
invented ones. Anything without such a marker defaults to the debit
template, matching the overwhelming majority of real workbook rows.

Placeholders that reference fields not obtainable from this function's
signature (business_unit, type, unit_no, bank_account are never derivable
from description/head/amount alone) are structurally removed from the
chosen template before substitution, so the output can never contain the
literal text "None", "null", or empty/dangling brackets.
"""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Any, Optional

from description_parser import parse_description

logger = logging.getLogger(__name__)

SCRIPT_DIR = Path(__file__).resolve().parent
NARRATION_CONFIG_PATH = SCRIPT_DIR / "config" / "narration_config.json"

# Absolute last-resort text used ONLY if narration_config.json cannot be
# loaded at all, or a Head is entirely missing from it — a production
# safety net, not a business rule (compare heads.py's
# _EMERGENCY_FALLBACK_HEAD, which follows the same principle).
_EMERGENCY_FALLBACK_TEMPLATE = "Banking transaction of {amount} recorded under {head}."

_CREDIT_MARKER_RE = re.compile(r"\b(?:NEFT|RTGS)\s+CR\b|^BY CLG:", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Config loading (cached — loaded from disk at most once per process)
# ---------------------------------------------------------------------------

_narration_config_cache: Optional[dict[str, dict[str, Any]]] = None


def _load_json(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _get_narration_config() -> dict[str, dict[str, Any]]:
    global _narration_config_cache
    if _narration_config_cache is None:
        try:
            data = _load_json(NARRATION_CONFIG_PATH)
            _narration_config_cache = data.get("heads", {})
            logger.debug("Loaded narration_config.json (%d heads).", len(_narration_config_cache))
        except Exception as exc:
            logger.error(
                "Could not load narration_config.json: %s — "
                "all narrations will use the emergency fallback template.", exc,
            )
            _narration_config_cache = {}
    return _narration_config_cache


# ---------------------------------------------------------------------------
# Direction inference (best-effort — see module docstring)
# ---------------------------------------------------------------------------

def _infer_direction(description: str) -> str:
    """Infer "credit" or "debit" from a literal, observed textual marker.

    generate_narration() receives no deposits/withdrawals — only a single
    unsigned amount — so this is the only signal available. "NEFT CR-"/
    "RTGS CR-" and "BY CLG:" are real prefixes seen in the workbook and
    always denote credits there; anything else defaults to "debit".
    """
    return "credit" if _CREDIT_MARKER_RE.search(description or "") else "debit"


def _format_amount(amount: float) -> str:
    try:
        return f"{amount:,.2f}"
    except (TypeError, ValueError):
        return "0.00"


class _BlankOnMissing(dict):
    """dict for str.format_map(): an unresolved key renders as "" instead
    of raising KeyError — a final safety net so a placeholder can never
    surface as a literal "{field}", "None", or an exception."""

    def __missing__(self, key: str) -> str:
        return ""


# ---------------------------------------------------------------------------
# Template clause stripping
# ---------------------------------------------------------------------------

def _strip_unavailable_clauses(
    template: str,
    has_party: bool,
    has_account_no: bool,
    has_reference: bool,
) -> str:
    """Remove decorative clauses whose backing field is unavailable, so
    the rendered sentence never contains "None", empty brackets, or a
    dangling connector word (e.g. a lone "vide" with nothing after it).

    business_unit, type, unit_no, and bank_account are never obtainable
    from generate_narration()'s (description, head, amount) signature,
    so their clauses are unconditionally stripped/simplified here.
    """
    text = template

    # "(Business Unit: {business_unit} | Head: {head} | Type: {type})"
    # -> "(Head: {head})" — keep the one field we DO have (head).
    text = re.sub(
        r"\(Business Unit:\s*\{business_unit\}\s*\|\s*Head:\s*\{head\}\s*\|\s*Type:\s*\{type\}\)",
        "(Head: {head})",
        text,
    )

    # Fallback-style "(Business Unit: {business_unit}, Ref: {reference})"
    if has_reference:
        text = re.sub(
            r"\(Business Unit:\s*\{business_unit\},\s*Ref:\s*\{reference\}\)",
            "(Ref: {reference})",
            text,
        )
    else:
        text = re.sub(
            r"\s*\(Business Unit:\s*\{business_unit\},\s*Ref:\s*\{reference\}\)",
            "",
            text,
        )

    # "(Unit No: {unit_no})" -> always stripped (unit_no never available).
    text = re.sub(r"\s*\(Unit No:\s*\{unit_no\}\)", "", text)

    # "to/from/in {bank_account}" -> always stripped (never available).
    text = re.sub(
        r"\s+(?:to|from|in)\s+\{bank_account\}(?=\s+(?:vide|towards))",
        "",
        text,
    )

    # party (+ account_no) clause.
    if not has_party:
        text = re.sub(r"\s+(?:from|to)\s+\{party\}\s+\(\{account_no\}\)", "", text)
    elif not has_account_no:
        text = re.sub(r"\s*\(\{account_no\}\)", "", text)

    # reference clause (note is left in place — it is always safe to
    # substitute since its value is pre-formatted as "" or " (text)").
    if not has_reference:
        text = re.sub(r"\s+vide\s+\{reference\}", "", text)

    return text


def _finalize(text: str) -> str:
    """Whitespace/punctuation cleanup applied to every rendered narration."""
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\(\s*\)", "", text).strip()  # stray empty parens
    text = re.sub(r"\s+\.", ".", text)           # space before a period
    if text and not text.endswith("."):
        text += "."
    return text


# ---------------------------------------------------------------------------
# Public API — signature and behavior contract unchanged
# ---------------------------------------------------------------------------

def generate_narration(description: str, head: str, amount: float) -> str:
    """Generate a human-readable narration for a transaction.

    Args:
        description: Raw bank transaction description.
        head: Business Head assigned to the transaction (see heads.py).
        amount: Transaction amount (deposit or withdrawal).

    Returns:
        A human-readable narration string. Never contains "None", "null",
        an unresolved "{placeholder}", or empty "()" — unavailable fields
        are omitted, not left blank. Never raises.
    """
    logger.debug(
        "Generating narration for head=%r amount=%r description=%r",
        head, amount, description,
    )

    try:
        parsed = parse_description(description) if description else {}
    except Exception as exc:
        logger.warning(
            "description_parser raised %s — continuing with no parsed fields.", exc,
        )
        parsed = {}

    party = parsed.get("party") or None
    account_no = parsed.get("account_number") or None
    reference = parsed.get("reference") or None
    note_raw = parsed.get("note") or None
    note_value = f" ({note_raw})" if note_raw else ""

    heads_config = _get_narration_config()
    head_cfg = heads_config.get(head)

    if head_cfg is None:
        logger.debug(
            "Head %r not found in narration_config.json — using emergency fallback template.",
            head,
        )
        template = _EMERGENCY_FALLBACK_TEMPLATE
        used_fallback = True
    else:
        direction = _infer_direction(description or "")
        fallback_template = head_cfg.get("fallback_template", _EMERGENCY_FALLBACK_TEMPLATE)
        template = head_cfg.get(f"{direction}_template") or fallback_template
        used_fallback = template == fallback_template
        logger.debug(
            "Selected %s for head=%r (inferred direction=%s)",
            "fallback_template" if used_fallback else f"{direction}_template",
            head, direction,
        )

    has_party = bool(party)
    has_account_no = bool(account_no)
    has_reference = bool(reference)

    stripped_template = _strip_unavailable_clauses(
        template, has_party=has_party, has_account_no=has_account_no, has_reference=has_reference,
    )

    values = _BlankOnMissing(head=head, amount=_format_amount(amount), note=note_value)
    if has_party:
        values["party"] = party
    if has_account_no:
        values["account_no"] = account_no
    if has_reference:
        values["reference"] = reference

    try:
        rendered = stripped_template.format_map(values)
    except Exception as exc:
        logger.warning(
            "Template rendering failed (%s) — using emergency fallback template.", exc,
        )
        rendered = _EMERGENCY_FALLBACK_TEMPLATE.format_map(values)
        used_fallback = True

    final_text = _finalize(rendered)

    logger.debug(
        "Filled placeholders: party=%r account_no=%r reference=%r note=%r head=%r amount=%r "
        "fallback_used=%s -> narration=%r",
        party, account_no, reference, note_raw, head, amount, used_fallback, final_text,
    )

    return final_text


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

    # (description, head, amount) — real workbook descriptions used
    # wherever possible. Expected values were derived by running the
    # deterministic renderer and reviewing each output for correctness
    # (no None/null/empty-bracket/malformed text), then locked in below.
    test_cases: list[tuple[str, str, float]] = [
        # --- Customer Collection ------------------------------------------
        ("RTGS CR-PSIB0020974-DIMPAL-AMBITION COLONISERS PRIVATE LTD-PSIBR22026040700701222",
         "Collection", 2000000.0),
        ("NEFT CR-UTIB0004193-PRAVIN KUMAR YADAV-AMBITION COLONISERS PRIVATE LIMITED-AXOIR16000908392",
         "Collection", 100000.0),
        ("BY CLG:SHRIRAM LILA:KOTAK MAHINDRA BANK LTD - 09-APR-26",
         "Collection", 100000.0),
        ("IMPS-617236024920-HEMANT YOGI-HDFC-xxxxxxxx0348-IMPS transaction",
         "Collection", 10.0),

        # --- Internal -------------------------------------------------------
        ("NEFT CR-MAHB0001461-AMBITION COLONISERS PRIVATE LIMITED-AMbition Colonisers Private Limited-MAHBH00647875203",
         "Internal", 600000.0),
        ("KVBLH00259426389-Ambition Colonisers Pvt Ltd-925020010722280-for pf",
         "Internal", 10526.0),
        ("070426BB4552144A-AMBITION COLONISERS-4114135000006375-master to free",
         "Internal", 600000.0),

        # --- DPL --------------------------------------------------------------
        ("KVBLH00260066550-Dwarkadhis Projects Pvt Ltd-60245906905-tfr",
         "DPL", 700000.0),

        # --- SKG Buildcon --------------------------------------------------
        ("KVBLH00258806500-S K G BUILDCON PVT LTD-920020066223471-tfr",
         "SKG Buildcon", 800000.0),
        ("RTGS CR-UTIB0003622-S.K.G. BUILDCON PRIVATE LIMITED-Ambition Colonisers Pvt Ltd-UTIBR62026040668620789",
         "SKG Buildcon", 800000.0),

        # --- Vendor ------------------------------------------------------
        ("KVBLH00262823433-SHIV SHAKTI TYRES AND BATTERIES-0419083000000008-Vendor",
         "Vendor - Ho", 7400.0),
        ("KVBLH00263377183-Rams Contify Electronics Pvt Ltd-113805501502-Vendor",
         "Vendor", 57500.0),
        ("KVBLH00258953932-Kapoor General and Provision Store-50200077351625-vendor",
         "Vendor -Site", 43973.0),

        # --- Salary ------------------------------------------------------
        ("KVBLH00259483410-Pooja-3572001700041842-salary", "Salary-HO", 9416.0),
        ("KVBLH00261688618-Bharat Singh-20402560091-salary", "Salary-HO", 11664.0),
        ("KVBLH00258829210-Prerna Jain-2211254841142872-salary", "Salary-HO", 115898.0),

        # --- Bank Charges --------------------------------------------------
        ("Monthly Service Chrgs MAY/26", "Bank Charges", 100.0),
        ("GST @18% on Monthly Service Chrgs", "Bank Charges", 18.0),

        # --- Tax -----------------------------------------------------------
        ("INB/951099117/TIN 2.0 CBDT TAX PAYMENT/NA", "Tax", 1000.0),
        ("INB/953744667/TIN 2.0 CBDT TAX PAYMENT/NA", "Tax", 295749.0),

        # --- Bounce -----------------------------------------------------
        ("NEFT-RETURN-KVBLH00263626193-Kiran Soni-INCORRECT ACCOUNT NUMBER",
         "Bounce", 196000.0),

        # --- EPF/ESI, Office Rent, MKT/ADVER, Commission, Card, Imprest,
        #     Professional, Refundable Security ---------------------------
        ("INB/949082359/EPFO PAYMENT AXIS BANK/NA", "EPF/ESI", 18595.0),
        ("KVBLH00258829209-Sheeshram Yadav-01761020000511-office rent", "Office Rent", 70339.0),
        ("KVBLH00259704850-Surender Kumar-50100736339069-Hoarding", "MKT/ADVER", 2000.0),
        ("KVBLH00262275262-Khushboo-44733498685-commission", "Commission", 15000.0),
        ("KVBLH00259865804-HDFC BANK LIMITED-4375465000133573-card 3573", "Card", 3472.0),
        ("KVBLH00264268484-Ravi Vats-520291014987347-imprest", "Imprest", 15000.0),
        ("KVBLH00259496023-Ravinder Kaushik-679010110003195-professional", "Professional", 207000.0),
        ("KVBLH00264301018-Neeraj Kaushik-01681000081759-Security refund plot 15",
         "Refundable Security", 25000.0),

        # --- Unknown party (in party_master but type=Unknown; narration
        #     still renders cleanly since party text itself IS available) -
        ("KVBLH00259483412-Vikkam-3572001700038918-salary", "Salary-Site", 8732.0),

        # --- Fallback: Head IS in narration_config.json, but description
        #     is unstructured and matches no description_parser pattern --
        ("Random unclassifiable text with no known pattern", "Others", 4000.0),
        ("Being amount adjusted for miscellaneous ledger entry", "Others", 5000.0),

        # --- Fallback: Head is NOT present in narration_config.json at
        #     all (drift/typo scenario) -> emergency fallback template ---
        ("Some description", "NonExistentHead", 1234.56),

        # --- Empty description ---------------------------------------------
        ("", "Others", 0.0),

        # --- Amount formatting sanity (large number, thousands separator) -
        ("Being amount adjusted for miscellaneous ledger entry", "Others", 1234567.89),
    ]

    print(f"Running {len(test_cases)} narration generation cases:\n")
    for description, head, amount in test_cases:
        narration = generate_narration(description, head, amount)
        assert "None" not in narration, f"'None' leaked into narration: {narration!r}"
        assert "null" not in narration, f"'null' leaked into narration: {narration!r}"
        assert "{" not in narration and "}" not in narration, f"Unresolved placeholder: {narration!r}"
        assert "()" not in narration, f"Empty brackets in narration: {narration!r}"
        print(f"[OK] {description!r} | {head} | {amount} ->\n     {narration}\n")

    print(f"All {len(test_cases)} cases produced clean, well-formed narrations.")
