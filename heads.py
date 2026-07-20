"""Phase 2: Classify a bank transaction into a business Head.

Public API (unchanged, must remain stable for existing callers):

    get_head(description, deposits, withdrawals) -> str

Internally this is now a configuration-driven decision engine built on
top of:

    description_parser.py    -> parse_description()
    config/party_master.json -> known party -> type mapping
    config/heads_config.json -> per-Head priority/direction/keywords/
                                 patterns/party_types/fallback

Decision order (see _decide_head for the full implementation):

    STEP 1  Party match (party_master.json) — if the party's type maps
            to exactly one Head, that Head is used directly.
    STEP 2  Direction filter (credit/debit/both) narrows candidates.
    STEP 3  Keyword rules (heads_config.json) — checked across all
            direction-eligible Heads in priority order.
    STEP 4  Description-pattern rules (heads_config.json) — same, if no
            keyword matched.
    STEP 5  If Step 1 found an ambiguous but small party-type pool
            (<=2 Heads) and nothing else matched, fall back to the
            lowest-priority Head in that pool (still grounded in the
            confirmed party type, not a guess). Otherwise return the
            single Head configured with "fallback": true in
            heads_config.json (currently "Others") — never hardcoded.

Nothing here is guessed: every decision traces back to a config entry,
a party_master entry, or the literal description text.
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Optional

from description_parser import parse_description

logger = logging.getLogger(__name__)

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_DIR = SCRIPT_DIR / "config"
PARTY_MASTER_PATH = CONFIG_DIR / "party_master.json"
HEADS_CONFIG_PATH = CONFIG_DIR / "heads_config.json"

# A pool of Heads sharing the same party_type is only treated as
# decisive-by-type (Step 5) when it is this small. Larger pools (e.g.
# "Vendor", which spans 9+ Heads) are too broad for a type-only decision
# to be anything but a guess, so they fall through to the configured
# fallback Head instead when keywords/patterns don't resolve them.
_AMBIGUOUS_TYPE_POOL_MAX = 2

# Absolute last-resort value used ONLY if the config files themselves
# cannot be loaded at all (corrupt/missing JSON) — a production-safety
# net, not a business rule. Normal "no match" cases always resolve via
# the "fallback": true entry in heads_config.json instead.
_EMERGENCY_FALLBACK_HEAD = "Others"


# ---------------------------------------------------------------------------
# Config loading (cached — loaded from disk at most once per process)
# ---------------------------------------------------------------------------

_party_lookup_cache: Optional[dict[str, dict[str, str]]] = None
_heads_config_cache: Optional[dict[str, dict[str, Any]]] = None
_heads_by_priority_cache: Optional[list[tuple[str, dict[str, Any]]]] = None
_fallback_head_cache: Optional[str] = None


def _load_json(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _build_party_lookup() -> dict[str, dict[str, str]]:
    """Build a case-insensitive lookup of every party name/alias ->
    {"canonical": ..., "type": ..., "confidence": ...}, loaded once."""
    data = _load_json(PARTY_MASTER_PATH)
    lookup: dict[str, dict[str, str]] = {}
    for canonical_name, entry in data.get("parties", {}).items():
        record = {
            "canonical": canonical_name,
            "type": entry.get("type", "Unknown"),
            "confidence": entry.get("confidence", "low"),
        }
        lookup[canonical_name.strip().upper()] = record
        for alias in entry.get("aliases", []):
            lookup[alias.strip().upper()] = record
    return lookup


def _get_party_lookup() -> dict[str, dict[str, str]]:
    global _party_lookup_cache
    if _party_lookup_cache is None:
        try:
            _party_lookup_cache = _build_party_lookup()
            logger.debug("Loaded party_master.json (%d entries).", len(_party_lookup_cache))
        except Exception as exc:
            logger.error("Could not load party_master.json: %s — party matching disabled.", exc)
            _party_lookup_cache = {}
    return _party_lookup_cache


def _get_heads_config() -> dict[str, dict[str, Any]]:
    global _heads_config_cache
    if _heads_config_cache is None:
        try:
            data = _load_json(HEADS_CONFIG_PATH)
            _heads_config_cache = data.get("heads", {})
            logger.debug("Loaded heads_config.json (%d heads).", len(_heads_config_cache))
        except Exception as exc:
            logger.error("Could not load heads_config.json: %s — classification disabled.", exc)
            _heads_config_cache = {}
    return _heads_config_cache


def _get_heads_by_priority() -> list[tuple[str, dict[str, Any]]]:
    global _heads_by_priority_cache
    if _heads_by_priority_cache is None:
        heads_config = _get_heads_config()
        _heads_by_priority_cache = sorted(
            (item for item in heads_config.items() if not item[1].get("fallback", False)),
            key=lambda item: item[1].get("priority", 999),
        )
    return _heads_by_priority_cache


def _get_fallback_head() -> str:
    global _fallback_head_cache
    if _fallback_head_cache is None:
        heads_config = _get_heads_config()
        for head_name, cfg in heads_config.items():
            if cfg.get("fallback", False):
                _fallback_head_cache = head_name
                break
        if _fallback_head_cache is None:
            logger.error(
                "No head in heads_config.json is marked \"fallback\": true — "
                "using emergency fallback %r.", _EMERGENCY_FALLBACK_HEAD,
            )
            _fallback_head_cache = _EMERGENCY_FALLBACK_HEAD
    return _fallback_head_cache


# ---------------------------------------------------------------------------
# Decision engine helpers
# ---------------------------------------------------------------------------

def _get_direction(deposits: float, withdrawals: float) -> Optional[str]:
    """Return "credit" if Deposits > 0, "debit" if Withdrawals > 0, else None."""
    if deposits > 0:
        return "credit"
    if withdrawals > 0:
        return "debit"
    return None


def _direction_matches(head_direction: str, direction: Optional[str]) -> bool:
    """A Head qualifies for a transaction's direction if the Head accepts
    "both", or the transaction's direction is unknown (no filtering
    possible), or the directions match exactly."""
    if head_direction == "both" or direction is None:
        return True
    return head_direction == direction


def _lookup_party(party_text: Optional[str]) -> Optional[dict[str, str]]:
    """Look up a parsed party string in party_master.json.

    Only high-confidence, non-Unknown matches are returned — a party
    tagged "Unknown"/"low" in party_master carries no decision-making
    weight here (using it would be guessing, which is disallowed).
    """
    if not party_text:
        return None
    record = _get_party_lookup().get(party_text.strip().upper())
    if record is None:
        return None
    if record["type"] == "Unknown" or record["confidence"] != "high":
        logger.debug(
            "Party %r matched party_master but type=%r/confidence=%r — "
            "not confident enough to use for classification.",
            party_text, record["type"], record["confidence"],
        )
        return None
    return record


def _heads_for_party_type(party_type: str) -> list[str]:
    heads_config = _get_heads_config()
    return [
        head_name
        for head_name, cfg in heads_config.items()
        if not cfg.get("fallback", False) and party_type in cfg.get("party_types", [])
    ]


# Below this length, a term's stripped form is too short to safely
# fallback-match against a whitespace-stripped description — some short
# heads_config.json keywords rely on being a distinct short word ("rent",
# "card") where stripping spaces adds no value but a coincidental
# substring match inside an unrelated word becomes more likely. See
# classify_transactions.py's identical guard for the concrete case
# ("ESI " as a word-boundary anchor) that motivated this.
_MIN_TERM_LEN_FOR_WHITESPACE_FALLBACK = 5


def _text_matches(desc_upper: str, term: str) -> bool:
    """Substring-match term against an already-uppercased description,
    tolerant of a stray space PDF extraction sometimes inserts mid-word
    (e.g. "CH RGS" instead of "CHRGS"). Checks the description as-is first
    (cheap, the common case), then a whitespace-stripped copy against a
    whitespace-stripped term — this only ever adds matches versus a plain
    substring check, never removes one. Short terms are excluded from the
    whitespace-stripped fallback (see _MIN_TERM_LEN_FOR_WHITESPACE_FALLBACK)."""
    term_upper = term.upper()
    if term_upper in desc_upper:
        return True
    term_nospace = term_upper.replace(" ", "")
    if len(term_nospace) < _MIN_TERM_LEN_FOR_WHITESPACE_FALLBACK:
        return False
    return term_nospace in desc_upper.replace(" ", "")


def _search_keywords(
    desc_upper: str, direction: Optional[str],
) -> Optional[str]:
    for head_name, cfg in _get_heads_by_priority():
        if not _direction_matches(cfg.get("transaction_direction", "both"), direction):
            continue
        for keyword in cfg.get("keywords", []):
            if _text_matches(desc_upper, keyword):
                logger.debug("Matched keyword: %r -> Head: %s", keyword, head_name)
                return head_name
    return None


def _search_description_patterns(
    desc_upper: str, direction: Optional[str],
) -> Optional[str]:
    for head_name, cfg in _get_heads_by_priority():
        if not _direction_matches(cfg.get("transaction_direction", "both"), direction):
            continue
        for pattern in cfg.get("description_patterns", []):
            if _text_matches(desc_upper, pattern):
                logger.debug("Matched description pattern: %r -> Head: %s", pattern, head_name)
                return head_name
    return None


def _decide_head(description: str, deposits: float, withdrawals: float) -> str:
    """Run the full configuration-driven decision engine for one transaction."""
    heads_config = _get_heads_config()
    if not heads_config:
        # Config failed to load entirely — production safety net only.
        return _get_fallback_head()

    desc_upper = description.upper()
    direction = _get_direction(deposits, withdrawals)

    try:
        parsed = parse_description(description)
    except Exception as exc:
        logger.warning("description_parser raised %s — continuing with raw description only.", exc)
        parsed = {"party": None}

    party_text = parsed.get("party") if parsed else None

    # ---- STEP 1: party match -------------------------------------------
    ambiguous_type_pool: list[str] = []
    party_record = _lookup_party(party_text)
    if party_record:
        candidate_heads = _heads_for_party_type(party_record["type"])
        if len(candidate_heads) == 1:
            head_name = candidate_heads[0]
            logger.debug(
                "Matched party: %s -> type %s -> Head: %s",
                party_record["canonical"], party_record["type"], head_name,
            )
            return head_name
        if candidate_heads:
            logger.debug(
                "Matched party: %s -> type %s -> %d candidate Heads (%s) — "
                "not decisive yet, checking keywords/patterns first.",
                party_record["canonical"], party_record["type"],
                len(candidate_heads), ", ".join(candidate_heads),
            )
            ambiguous_type_pool = candidate_heads

    # ---- STEP 2 + STEP 3: direction filter + keyword rules --------------
    head_name = _search_keywords(desc_upper, direction)
    if head_name:
        return head_name

    # ---- STEP 4: description-pattern rules -------------------------------
    head_name = _search_description_patterns(desc_upper, direction)
    if head_name:
        return head_name

    # ---- STEP 5: fallback -------------------------------------------------
    if ambiguous_type_pool and len(ambiguous_type_pool) <= _AMBIGUOUS_TYPE_POOL_MAX:
        direction_ok = [
            h for h in ambiguous_type_pool
            if _direction_matches(heads_config[h].get("transaction_direction", "both"), direction)
        ]
        pool = direction_ok or ambiguous_type_pool
        head_name = min(pool, key=lambda h: heads_config[h].get("priority", 999))
        logger.debug(
            "No keyword/pattern matched. Falling back to party-type match only "
            "(pool size %d) -> Head: %s", len(ambiguous_type_pool), head_name,
        )
        return head_name

    fallback_head = _get_fallback_head()
    logger.debug("No rule matched -> Head: %s", fallback_head)
    return fallback_head


# ---------------------------------------------------------------------------
# Public API — signature and behavior contract unchanged
# ---------------------------------------------------------------------------

def is_internal_type_head(head_name: str) -> bool:
    """Return True if this Head's config entry has "Internal" among its
    party_types (config/heads_config.json) — meaning it represents a
    transfer to/from one of our own internal-type companies, however
    get_head() actually arrived at it (a direct party_master match, a
    keyword like "dwarkadhis"/"for esi", or a description pattern).

    classify_transactions.py's resolve_business_fields() uses this so any
    current or future Head tagged this way gets its Business Unit/Type/
    TCP filled in with the standard internal-transfer defaults, instead
    of a hardcoded list of specific Head names (like the original
    {"Internal", "DPL"} set) that would need updating by hand every time
    party_master.json or heads_config.json grows a new internal-type
    entry.
    """
    heads_config = _get_heads_config()
    cfg = heads_config.get(head_name)
    return bool(cfg) and "Internal" in cfg.get("party_types", [])


def get_head(description: str, deposits: float, withdrawals: float) -> str:
    """Return the business Head for a transaction.

    Args:
        description: Raw bank transaction description.
        deposits: Deposit amount for the transaction (0 if none).
        withdrawals: Withdrawal amount for the transaction (0 if none).

    Returns:
        The Head name assigned to this transaction. Always a string —
        never None — falling back to the Head configured with
        "fallback": true in config/heads_config.json when no rule
        confidently matches.
    """
    if not description:
        fallback_head = _get_fallback_head()
        logger.debug("Empty description — defaulting to %s.", fallback_head)
        return fallback_head

    logger.debug(
        "Classifying description=%r deposits=%r withdrawals=%r",
        description, deposits, withdrawals,
    )

    return _decide_head(description, deposits, withdrawals)


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

    # (description, deposits, withdrawals, expected_head) — real workbook
    # descriptions used wherever possible; a few synthetic ones are
    # marked where no real example exists for that scenario.
    test_cases = [
        # --- Known Customer (party_master) -> Collection ---------------
        ("RTGS CR-PSIB0020974-DIMPAL-AMBITION COLONISERS PRIVATE LTD-PSIBR22026040700701222",
         2000000.0, 0.0, "Collection"),
        ("NEFT CR-UTIB0004193-PRAVIN KUMAR YADAV-AMBITION COLONISERS PRIVATE LIMITED-AXOIR16000908392",
         100000.0, 0.0, "Collection"),
        ("IMPS-617236024920-HEMANT YOGI-HDFC-xxxxxxxx0348-IMPS transaction",
         10.0, 0.0, "Collection"),
        ("BY CLG:SHRIRAM LILA:KOTAK MAHINDRA BANK LTD - 09-APR-26",
         100000.0, 0.0, "Collection"),

        # --- Known Internal (party_master, company's own name) ----------
        ("NEFT CR-MAHB0001461-AMBITION COLONISERS PRIVATE LIMITED-AMbition Colonisers Private Limited-MAHBH00647875203",
         600000.0, 0.0, "Internal"),
        ("KVBLH00259426389-Ambition Colonisers Pvt Ltd-925020010722280-for pf",
         0.0, 10526.0, "Internal"),
        ("070426BB4552144A-AMBITION COLONISERS-4114135000006375-master to free",
         600000.0, 0.0, "Internal"),
        ("KVBLH00260070448-AMBITIONCOLONISERSPRIVATELIM-KARUR VYSYA BANK-for tds",
         0.0, 263042.0, "Internal"),

        # --- Dwarkadhis Projects Pvt Ltd -> Internal (not a separate "DPL"
        #     Head — cross-checked against the accounts team's own reference
        #     file, which never uses "DPL" as a Head value; heads_config.json
        #     used to have a standalone "DPL" head here, confirmed wrong and
        #     merged into "Internal"'s own keyword list) --------------------
        ("KVBLH00260066550-Dwarkadhis Projects Pvt Ltd-60245906905-tfr",
         0.0, 700000.0, "Internal"),

        # --- Known Vendor (party_master), resolved via a party-specific
        #     keyword (SKG Buildcon) rather than the generic Vendor type --
        ("KVBLH00258806500-S K G BUILDCON PVT LTD-920020066223471-tfr",
         0.0, 800000.0, "SKG Buildcon"),
        # NOTE: heads_config.json's SKG Buildcon keywords ("skg buildcon",
        # "s k g buildcon") do not cover this dotted real-world variant
        # ("S.K.G. BUILDCON"), and party_master's exact-match lookup also
        # doesn't bridge it into a decisive single-Head type pool (Vendor
        # spans 9 Heads). This is a genuine, pre-existing config coverage
        # gap — not an engine bug — documented here rather than papered
        # over; fixing it means updating heads_config.json, which is out
        # of scope for this heads.py-only task.
        ("RTGS CR-UTIB0003622-S.K.G. BUILDCON PRIVATE LIMITED-Ambition Colonisers Pvt Ltd-UTIBR62026040668620789",
         800000.0, 0.0, "Others"),

        # --- Unknown party (party_master has it, but type=Unknown) ------
        # Employee salary rows: "Pooja"/"Vikkam" are in party_master as
        # type=Unknown, so party matching must NOT decide these — the
        # generic "salary" keyword should.
        ("KVBLH00259483410-Pooja-3572001700041842-salary",
         0.0, 9416.0, "Salary-HO"),
        ("KVBLH00258829209-Sheeshram Yadav-01761020000511-office rent",
         0.0, 70339.0, "Office Rent"),

        # --- Unrecognized / no party_master entry at all ----------------
        # "Vendor - Ho" wins over "Vendor" here because both share the
        # keyword "vendor" and Vendor - Ho has a lower priority number
        # (checked first) — get_head() has no account/Project info to
        # disambiguate Ho vs Site vs generic Vendor from text alone. This
        # is the same known, documented ambiguity as heads_config.json's
        # Vendor / Vendor - Ho / Vendor -Site limitation.
        ("KVBLH00262823433-SHIV SHAKTI TYRES AND BATTERIES-0419083000000008-Vendor",
         0.0, 7400.0, "Vendor - Ho"),

        # --- Credit direction --------------------------------------------
        ("RTGS CR-YESB0000037-DIMPAL-AMBITION COLONISERS PRIVATE-YESBR52026062257934896",
         700000.0, 0.0, "Collection"),

        # --- Debit direction -----------------------------------------------
        ("KVBLH00259483411-Raj Kumar Mahto-1649255516-salary",
         0.0, 10440.0, "Salary-HO"),

        # --- Bank Charges ---------------------------------------------------
        ("Monthly Service Chrgs MAY/26", 0.0, 100.0, "Bank Charges"),
        ("GST @18% on Monthly Service Chrgs", 0.0, 18.0, "Bank Charges"),

        # --- Salary ------------------------------------------------------
        ("KVBLH00261688618-Bharat Singh-20402560091-salary", 0.0, 11664.0, "Salary-HO"),
        ("KVBLH00258829210-Prerna Jain-2211254841142872-salary", 0.0, 115898.0, "Salary-HO"),

        # --- Bounce ------------------------------------------------------
        ("NEFT-RETURN-KVBLH00263626193-Kiran Soni-INCORRECT ACCOUNT NUMBER",
         196000.0, 0.0, "Bounce"),

        # --- Tax -----------------------------------------------------------
        ("INB/951099117/TIN 2.0 CBDT TAX PAYMENT/NA", 0.0, 1000.0, "Tax"),
        ("INB/953744667/TIN 2.0 CBDT TAX PAYMENT/NA", 0.0, 295749.0, "Tax"),

        # --- EPF/ESI ---------------------------------------------------
        ("INB/949082359/EPFO PAYMENT AXIS BANK/NA", 0.0, 18595.0, "EPF/ESI"),
        ("INB/949083764/ESIC PAYMENTS/NA", 0.0, 2583.0, "EPF/ESI"),

        # --- Refundable Security / a real "refund"-style transaction ----
        ("KVBLH00264301018-Neeraj Kaushik-01681000081759-Security refund plot 15",
         0.0, 25000.0, "Refundable Security"),

        # --- Vendor Payment ------------------------------------------------
        # Real workbook row, tagged "Vendor" manually with NO textual
        # keyword present in the description at all. "Raj Tyres" IS a
        # known Vendor in party_master, but Vendor spans 9 Heads (an
        # intentionally-too-broad pool per _AMBIGUOUS_TYPE_POOL_MAX), so
        # the engine correctly refuses to guess and returns the
        # configured fallback rather than picking one of the 9 blindly.
        ("INB/NEFT/AXODH13902403426/RAJ TYRES/IDFC FIRST BANK LTD//////",
         0.0, 23600.0, "Others"),
        # Same Vendor/Vendor-Ho keyword-ambiguity as SHIV SHAKTI above.
        ("KVBLH00263377183-Rams Contify Electronics Pvt Ltd-113805501502-Vendor",
         0.0, 57500.0, "Vendor - Ho"),

        # --- MKT/ADVER, Commission, Card, Imprest, Professional ----------
        ("KVBLH00259704850-Surender Kumar-50100736339069-Hoarding", 0.0, 2000.0, "MKT/ADVER"),
        ("KVBLH00262275262-Khushboo-44733498685-commission", 0.0, 15000.0, "Commission"),
        ("KVBLH00259865804-HDFC BANK LIMITED-4375465000133573-card 3573", 0.0, 3472.0, "Credit Card"),
        ("KVBLH00264268484-Ravi Vats-520291014987347-imprest", 0.0, 15000.0, "Imprest"),
        ("KVBLH00259496023-Ravinder Kaushik-679010110003195-professional", 0.0, 207000.0, "Professional"),

        # --- Others: no rule matches at all (no matching keyword/pattern,
        #     no party_master entry, unrecognized text) -------------------
        ("Being amount adjusted for miscellaneous ledger entry", 0.0, 5000.0, "Others"),
        ("Random unclassifiable text with no known pattern", 4000.0, 0.0, "Others"),

        # --- Empty description --------------------------------------------
        ("", 0.0, 0.0, "Others"),
    ]

    passed = 0
    for description, deposits, withdrawals, expected in test_cases:
        result = get_head(description, deposits, withdrawals)
        status = "PASS" if result == expected else "FAIL"
        if status == "PASS":
            passed += 1
        print(f"[{status}] {description!r} (D={deposits}, W={withdrawals}) -> {result} (expected {expected})")

    print(f"\n{passed}/{len(test_cases)} test cases passed.")
