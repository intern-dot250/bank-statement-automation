"""Standalone parser: raw bank Description text -> structured fields.

Patterns are reverse-engineered directly from real Description strings
found in 'Copy of AMB Bank Statements 2026-27.xlsx' (5 bank sheets).
Each handler below corresponds to one concretely observed format family.

The one exception is UPI: no UPI transaction exists anywhere in the
source workbook. Per the task requirement to support UPI, its pattern
is modeled structurally on the observed IMPS format (same bank/UPI
handle layout convention), since no real UPI example is available to
derive a pattern from. This is called out explicitly wherever relevant
so it is never mistaken for an observed pattern.

This module is standalone: it is not imported by, and does not import,
any other project file (heads.py, narration.py, classify_transactions.py,
run_pipeline.py, etc.).
"""

from __future__ import annotations

import logging
import re
from typing import Optional, TypedDict

logger = logging.getLogger(__name__)


class ParsedDescription(TypedDict):
    party: Optional[str]
    payment_mode: Optional[str]
    reference: Optional[str]
    account_number: Optional[str]
    note: Optional[str]


def _empty_result() -> ParsedDescription:
    return {
        "party": None,
        "payment_mode": None,
        "reference": None,
        "account_number": None,
        "note": None,
    }


def _clean(value: Optional[str]) -> Optional[str]:
    """Trim whitespace; treat an empty result as None rather than ''."""
    if value is None:
        return None
    value = value.strip()
    return value if value else None


# ---------------------------------------------------------------------------
# Pattern handlers, one per observed Description format.
# Each handler returns a ParsedDescription on match, or None if it doesn't
# apply. Handlers are tried in order in parse_description(); the first
# match wins. No handler fabricates a value it cannot confidently read.
# ---------------------------------------------------------------------------

# "NEFT CR-MAHB0001461-AMBITION COLONISERS PRIVATE LIMITED-AMbition Colonisers Private Limited-MAHBH00647875203"
# "RTGS CR-PSIB0020974-DIMPAL-AMBITION COLONISERS PRIVATE LTD-PSIBR22026040700701222"
_PAT_NEFT_RTGS_CR = re.compile(
    r"^(?P<mode>NEFT|RTGS)\s+CR-(?P<ifsc>[A-Z0-9]+)-(?P<party>[^-]+)-(?P<self>[^-]+)-(?P<ref>[A-Z0-9]+)$"
)


def _handle_neft_rtgs_cr(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_NEFT_RTGS_CR.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": m.group("mode"),
        "reference": _clean(m.group("ref")),
        "account_number": None,
        "note": None,
    }


# "IMPS-617236024920-HEMANT YOGI-HDFC-xxxxxxxx0348-IMPS transaction"
# "IMPS-612921863417-Shobha Jain-AUBL-xxxxxxxxxxxx9179-Loan Repayment"
_PAT_IMPS_DASH = re.compile(
    r"^IMPS-(?P<ref>\d+)-(?P<party>[^-]+)-(?P<bank>[A-Z]+)-(?P<acct>[xX0-9]+)-(?P<note>.+)$"
)


def _handle_imps_dash(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_IMPS_DASH.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "IMPS",
        "reference": _clean(m.group("ref")),
        "account_number": _clean(m.group("acct")),
        "note": _clean(m.group("note")),
    }


# UPI: NOT observed anywhere in the workbook. Modeled structurally on the
# IMPS dash format above (same bank/masked-account layout convention),
# since IMPS and UPI descriptions from Indian banks commonly share this
# shape. No live example backs this pattern — documented, not invented
# as a "real" rule.
_PAT_UPI_DASH = re.compile(
    r"^UPI-(?P<ref>\d+)-(?P<party>[^-]+)-(?P<bank>[A-Z]+)-(?P<acct>[xX0-9@.a-zA-Z]+)-(?P<note>.+)$"
)


def _handle_upi_dash(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_UPI_DASH.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "UPI",
        "reference": _clean(m.group("ref")),
        "account_number": _clean(m.group("acct")),
        "note": _clean(m.group("note")),
    }


# "NEFT/KVBLH00259426392/AMBITIONCOLONISERSPRIVATELIM/KARUR VYSYA BANK/NA/NAfor esi/NANA/NA"
_PAT_NEFT_SLASH = re.compile(
    r"^NEFT/(?P<ref>[A-Z0-9]+)/(?P<party>[^/]+)/(?P<bank>[^/]+)/NA/NA(?P<note>[^/]*)/NANA/NA$"
)


def _handle_neft_slash(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_NEFT_SLASH.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "NEFT",
        "reference": _clean(m.group("ref")),
        "account_number": None,
        "note": _clean(m.group("note")),
    }


# "INB/NEFT/AXODH13902403426/RAJ TYRES/IDFC FIRST BANK LTD//////"
_PAT_INB_NEFT = re.compile(
    r"^INB/NEFT/(?P<ref>[A-Z0-9]+)/(?P<party>[^/]+)/(?P<bank>[^/]+)/+$"
)


def _handle_inb_neft(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_INB_NEFT.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "NEFT",
        "reference": _clean(m.group("ref")),
        "account_number": None,
        "note": None,
    }


# "INB/949082359/EPFO PAYMENT AXIS BANK/NA"
# "INB/951099117/TIN 2.0 CBDT TAX PAYMENT/NA"
_PAT_INB_STATUTORY = re.compile(r"^INB/(?P<ref>\d+)/(?P<note>.+)/NA$")


def _handle_inb_statutory(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_INB_STATUTORY.match(desc)
    if not m:
        return None
    return {
        "party": None,
        "payment_mode": "INB",
        "reference": _clean(m.group("ref")),
        "account_number": None,
        "note": _clean(m.group("note")),
    }


# "NEFT-RETURN-KVBLH00263626193-Kiran Soni-INCORRECT ACCOUNT NUMBER"
_PAT_NEFT_RETURN = re.compile(
    r"^NEFT-RETURN-(?P<ref>[A-Z0-9]+)-(?P<party>[^-]+)-(?P<note>.+)$"
)


def _handle_neft_return(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_NEFT_RETURN.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "NEFT",
        "reference": _clean(m.group("ref")),
        "account_number": None,
        "note": _clean(m.group("note")),
    }


# "BY CLG:SHRIRAM LILA:KOTAK MAHINDRA BANK LTD - 09-APR-26"
_PAT_CLG = re.compile(r"^BY CLG:(?P<party>[^:]+):(?P<note>.+)$")


def _handle_clg(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_CLG.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "CLG",
        "reference": None,
        "account_number": None,
        "note": _clean(m.group("note")),
    }


# "DD ISSUED/SAK/SDO/OP DHBVN DHARUHERA/atPar"
_PAT_DD = re.compile(
    r"^DD ISSUED/(?P<ref1>[^/]+)/(?P<ref2>[^/]+)/(?P<party>[^/]+)/(?P<note>.+)$"
)


def _handle_dd(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_DD.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "DD",
        "reference": _clean(f"{m.group('ref1')}/{m.group('ref2')}"),
        "account_number": None,
        "note": _clean(m.group("note")),
    }


# "BILLDESK-YKVB2692225493-IGL-IGL Payment-1763102000000018 H.O"
_PAT_BILLDESK = re.compile(
    r"^BILLDESK-(?P<ref>[A-Z0-9]+)-(?P<party>[^-]+)-(?P<note>[^-]+)-(?P<acct>\d.*)$"
)


def _handle_billdesk(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_BILLDESK.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "BILLDESK",
        "reference": _clean(m.group("ref")),
        "account_number": _clean(m.group("acct")),
        "note": _clean(m.group("note")),
    }


# "YIB-TPT-DWARKADHIS PROJECTS PVT LT D-TFR F-045563400002477"
# "YIB-TPT-DWARKADHIS PROJECTS PRIVATE LIMITED-TFR R-04556320000 0377"
# Internal-transfer format actually seen in the live bank statements: a
# bank/channel code, "TPT", our own company name (often mid-word-wrapped
# by the PDF extractor), then "TFR <reference>". The reference is
# whatever follows "TFR" (kept as-is, including any embedded space from a
# PDF wrap artifact, since it's still usable for last-4-digit matching).
_PAT_YIB_TPT = re.compile(
    r"^(?P<code>[A-Z]+)-TPT-(?P<company>.+?)-TFR[\s-]+(?P<ref>.+)$",
    re.IGNORECASE,
)


def _handle_yib_tpt(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_YIB_TPT.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("company")),
        "payment_mode": "TPT",
        "reference": _clean(m.group("ref")),
        "account_number": None,
        "note": None,
    }


# "UPI/310650751520/FROM:9799400249 -2@AXL/TO:045563200000264@YESB00 00455.IFSC.NPCI/PAYMENT FROM PHONE PE"
# UPI format actually seen in the live bank statements (distinct from the
# structurally-modeled _PAT_UPI_DASH above, which uses hyphens — this one
# uses slashes and explicit FROM:/TO: VPA labels). The paying party is
# whatever appears before "@" in the FROM: VPA (often a phone number, not
# a human name — still the most specific identifier the bank gives us).
_PAT_UPI_SLASH = re.compile(
    r"^UPI/(?P<ref>\d+)/FROM:(?P<from_vpa>[^/]+?)@[^/]+/TO:(?P<to_vpa>[^/]+?)@[^/]+/(?P<note>.+)$",
    re.IGNORECASE,
)


def _handle_upi_slash(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_UPI_SLASH.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("from_vpa").replace(" ", "")),
        "payment_mode": "UPI",
        "reference": _clean(m.group("ref")),
        "account_number": _clean(m.group("to_vpa").replace(" ", "")),
        "note": _clean(m.group("note")),
    }


# "IMPS/AMITKUMAR/XXX3986/RR N:618712584961/AXIS BANK"
# "IMPS/JAYANT RAITANI/XXX8180/RRN :618614869331/"
# IMPS format actually seen in the live bank statements: IMPS/<party>/
# XXX<masked account digits>/RRN:<reference>/<bank name, sometimes
# blank>. "RRN" is matched tolerant of a stray embedded space (PDF wrap
# artifact, same issue as elsewhere in this module).
_PAT_IMPS_SLASH_RRN = re.compile(
    r"^IMPS/(?P<party>[^/]+)/\s*XXX\s*(?P<acct>[\d\s]+)/RR\s*N\s*:\s*(?P<ref>\d+)/(?P<bank>.*)$",
    re.IGNORECASE,
)


def _handle_imps_slash_rrn(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_IMPS_SLASH_RRN.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "IMPS",
        "reference": _clean(m.group("ref")),
        "account_number": _clean(f"XXX{m.group('acct').replace(' ', '')}"),
        "note": _clean(m.group("bank")),
    }


# "YIB-NEFT-YESME61850064653-LALAN YADAV-FDRL0002158-CONTRACTOR-FEDERAL BANK"
# "YIB-NEFT-YESME61850057030-CHOU DHARY ENTERPRISES-IDIB000D618-VEN DOR-INDIAN BANK"
# Vendor/Contractor/Professional payment format seen in the live
# statements: YIB-NEFT-<UTR ref>-<party>-<counterparty IFSC>-<role>-
# <bank name>. The role segment (Vendor/Contractor/Professional) is
# already detected separately by classify_transactions.py's role
# extraction; this handler just supplies party/reference for the
# narration, so it doesn't need to interpret the role itself.
_PAT_YIB_NEFT_ROLE = re.compile(
    r"^YIB-NEFT-(?P<ref>[A-Z0-9]+)-(?P<party>[^-]+)-(?P<ifsc>[A-Z0-9]+)-(?P<role>[^-]+)-(?P<bank>.+)$",
    re.IGNORECASE,
)


def _handle_yib_neft_role(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_YIB_NEFT_ROLE.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "NEFT",
        "reference": _clean(m.group("ref")),
        "account_number": None,
        "note": _clean(m.group("bank")),
    }


# "NET-TPT-ASHISH SHUKLA-071691 900001571"
# Internal-transfer-channel format seen in the live statements without an
# own-company name segment (just a channel code, party, then reference) —
# distinct from the YIB-TPT format above, which always includes a company
# name segment before "TFR".
_PAT_NET_TPT = re.compile(
    r"^NET-TPT-(?P<party>[^-]+)-(?P<ref>.+)$",
    re.IGNORECASE,
)


def _handle_net_tpt(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_NET_TPT.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": "TPT",
        "reference": _clean(m.group("ref")),
        "account_number": None,
        "note": None,
    }


# "KVBLH00258806500-S K G BUILDCON PVT LTD-920020066223471-tfr"
# "070426BB4552144A-AMBITION COLONISERS-4114135000006375-master to free"
# "C0000240-VANDANA KHULLAR-DWARKADHIS PROJECTS PVT L-HDFCR52026070478907371"
# "DSF0002-ROHITAS K UMAR-DWARKADHIS PROJECT S PRIVA-IN12618745392512"
# "NEFT CR-HDFC0000001-VIJAY YAD AV-DWARKADHIS PROJECTS PVT L-HDFC H01104907359"
# "RTGS CR-ICIC0000083-RAJEEV SAIN I-DWARKADHIS PROJECTS PRIVA- ICICR12026070411554969"
# Customer-collection format: an optional "NEFT CR-"/"RTGS CR-" channel
# prefix, a short code, the paying customer's name, our OWN company name
# (as the beneficiary, often mid-word-wrapped by the PDF extractor), then
# a bank reference (itself sometimes wrapped, hence the ref pattern
# allows an internal space). Distinguished from the generic UTR-dash
# format below by the 3rd segment containing our own company name
# instead of a counterparty account number.
_PAT_COLLECTION_DASH = re.compile(
    r"^(?:(?:NEFT|RTGS)\s+CR-)?(?P<code>[A-Z0-9]+)-(?P<party>[^-]+)-"
    r"(?P<company>(?:DWARKADHIS|AMBITION)[^-]*)-\s*(?P<ref>[A-Za-z0-9]+\s?[A-Za-z0-9]*)$",
    re.IGNORECASE,
)


def _handle_collection_dash(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_COLLECTION_DASH.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": None,
        "reference": _clean(m.group("ref")),
        "account_number": None,
        "note": None,
    }


# Generic 4-segment UTR-code dash format. No explicit IMPS/NEFT/RTGS/UPI
# keyword appears in these descriptions, so payment_mode is left null
# rather than guessed, even though these are typically NEFT-style UTRs.
_PAT_UTR_DASH = re.compile(
    r"^(?P<ref>[A-Z0-9]{10,})-(?P<party>[^-]+)-(?P<acct>\d{6,})-(?P<note>.+)$"
)


def _handle_utr_dash(desc: str) -> Optional[ParsedDescription]:
    m = _PAT_UTR_DASH.match(desc)
    if not m:
        return None
    return {
        "party": _clean(m.group("party")),
        "payment_mode": None,
        "reference": _clean(m.group("ref")),
        "account_number": _clean(m.group("acct")),
        "note": _clean(m.group("note")),
    }


# Ordered list of handlers. More specific patterns are tried before the
# generic UTR-dash fallback pattern, to avoid a specific format being
# incorrectly swallowed by a looser one.
_HANDLERS = (
    _handle_neft_rtgs_cr,
    _handle_neft_return,
    _handle_imps_dash,
    _handle_imps_slash_rrn,
    _handle_upi_slash,
    _handle_upi_dash,
    _handle_yib_neft_role,
    _handle_yib_tpt,
    _handle_net_tpt,
    _handle_neft_slash,
    _handle_inb_neft,
    _handle_inb_statutory,
    _handle_clg,
    _handle_dd,
    _handle_billdesk,
    _handle_collection_dash,
    _handle_utr_dash,
)


def parse_description(description: Optional[str]) -> ParsedDescription:
    """Convert a raw bank Description string into structured fields.

    Args:
        description: Raw bank transaction description text.

    Returns:
        A dict with keys "party", "payment_mode", "reference",
        "account_number", "note". Any field that cannot be confidently
        extracted is None — nothing is guessed.
    """
    if not description or not description.strip():
        logger.debug("Empty description — returning all-null result.")
        return _empty_result()

    desc = description.strip()

    for handler in _HANDLERS:
        result = handler(desc)
        if result is not None:
            logger.debug("Matched %s for description=%r", handler.__name__, desc)
            return result

    logger.debug("No known pattern matched description=%r — returning all-null result.", desc)
    return _empty_result()


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

    # Real Description strings taken directly from the workbook
    # ('Copy of AMB Bank Statements 2026-27.xlsx'), plus one synthetic
    # UPI example (marked) since no real UPI row exists in the source.
    test_cases = [
        (
            "NEFT CR-MAHB0001461-AMBITION COLONISERS PRIVATE LIMITED-AMbition Colonisers Private Limited-MAHBH00647875203",
            {"party": "AMBITION COLONISERS PRIVATE LIMITED", "payment_mode": "NEFT",
             "reference": "MAHBH00647875203", "account_number": None, "note": None},
        ),
        (
            "RTGS CR-PSIB0020974-DIMPAL-AMBITION COLONISERS PRIVATE LTD-PSIBR22026040700701222",
            {"party": "DIMPAL", "payment_mode": "RTGS",
             "reference": "PSIBR22026040700701222", "account_number": None, "note": None},
        ),
        (
            "RTGS CR-UTIB0003622-S.K.G. BUILDCON PRIVATE LIMITED-Ambition Colonisers Pvt Ltd-UTIBR62026040668620789",
            {"party": "S.K.G. BUILDCON PRIVATE LIMITED", "payment_mode": "RTGS",
             "reference": "UTIBR62026040668620789", "account_number": None, "note": None},
        ),
        (
            "RTGS CR-YESB0000037-DIMPAL-AMBITION COLONISERS PRIVATE LTD-YESBR52026050557667138",
            {"party": "DIMPAL", "payment_mode": "RTGS",
             "reference": "YESBR52026050557667138", "account_number": None, "note": None},
        ),
        (
            "NEFT CR-UTIB0004193-PRAVIN KUMAR YADAV-AMBITION COLONISERS PRIVATE LIMITED-AXOIR16000908392",
            {"party": "PRAVIN KUMAR YADAV", "payment_mode": "NEFT",
             "reference": "AXOIR16000908392", "account_number": None, "note": None},
        ),
        (
            "IMPS-617236024920-HEMANT YOGI-HDFC-xxxxxxxx0348-IMPS transaction",
            {"party": "HEMANT YOGI", "payment_mode": "IMPS",
             "reference": "617236024920", "account_number": "xxxxxxxx0348", "note": "IMPS transaction"},
        ),
        (
            "IMPS-612921863417-Shobha Jain-AUBL-xxxxxxxxxxxx9179-Loan Repayment",
            {"party": "Shobha Jain", "payment_mode": "IMPS",
             "reference": "612921863417", "account_number": "xxxxxxxxxxxx9179", "note": "Loan Repayment"},
        ),
        (
            "IMPS-613215015084-Shobha Jain-AUBL-xxxxxxxx9006-Money Transfer",
            {"party": "Shobha Jain", "payment_mode": "IMPS",
             "reference": "613215015084", "account_number": "xxxxxxxx9006", "note": "Money Transfer"},
        ),
        (
            "NEFT/KVBLH00259426392/AMBITIONCOLONISERSPRIVATELIM/KARUR VYSYA BANK/NA/NAfor esi/NANA/NA",
            {"party": "AMBITIONCOLONISERSPRIVATELIM", "payment_mode": "NEFT",
             "reference": "KVBLH00259426392", "account_number": None, "note": "for esi"},
        ),
        (
            "NEFT/KVBLH00260070448/AMBITIONCOLONISERSPRIVATELIM/KARUR VYSYA BANK/NA/NAfor tds/NANA/NA",
            {"party": "AMBITIONCOLONISERSPRIVATELIM", "payment_mode": "NEFT",
             "reference": "KVBLH00260070448", "account_number": None, "note": "for tds"},
        ),
        (
            "INB/NEFT/AXODH13902403426/RAJ TYRES/IDFC FIRST BANK LTD//////",
            {"party": "RAJ TYRES", "payment_mode": "NEFT",
             "reference": "AXODH13902403426", "account_number": None, "note": None},
        ),
        (
            "INB/949082359/EPFO PAYMENT AXIS BANK/NA",
            {"party": None, "payment_mode": "INB",
             "reference": "949082359", "account_number": None, "note": "EPFO PAYMENT AXIS BANK"},
        ),
        (
            "INB/951099117/TIN 2.0 CBDT TAX PAYMENT/NA",
            {"party": None, "payment_mode": "INB",
             "reference": "951099117", "account_number": None, "note": "TIN 2.0 CBDT TAX PAYMENT"},
        ),
        (
            "NEFT-RETURN-KVBLH00263626193-Kiran Soni-INCORRECT ACCOUNT NUMBER",
            {"party": "Kiran Soni", "payment_mode": "NEFT",
             "reference": "KVBLH00263626193", "account_number": None, "note": "INCORRECT ACCOUNT NUMBER"},
        ),
        (
            "BY CLG:SHRIRAM LILA:KOTAK MAHINDRA BANK LTD - 09-APR-26",
            {"party": "SHRIRAM LILA", "payment_mode": "CLG",
             "reference": None, "account_number": None, "note": "KOTAK MAHINDRA BANK LTD - 09-APR-26"},
        ),
        (
            "DD ISSUED/SAK/SDO/OP DHBVN DHARUHERA/atPar",
            {"party": "OP DHBVN DHARUHERA", "payment_mode": "DD",
             "reference": "SAK/SDO", "account_number": None, "note": "atPar"},
        ),
        (
            "BILLDESK-YKVB2692225493-IGL-IGL Payment-1763102000000018 H.O",
            {"party": "IGL", "payment_mode": "BILLDESK",
             "reference": "YKVB2692225493", "account_number": "1763102000000018 H.O", "note": "IGL Payment"},
        ),
        (
            "KVBLH00258806500-S K G BUILDCON PVT LTD-920020066223471-tfr",
            {"party": "S K G BUILDCON PVT LTD", "payment_mode": None,
             "reference": "KVBLH00258806500", "account_number": "920020066223471", "note": "tfr"},
        ),
        (
            "KVBLH00258953932-Kapoor General and Provision Store-50200077351625-vendor",
            {"party": "Kapoor General and Provision Store", "payment_mode": None,
             "reference": "KVBLH00258953932", "account_number": "50200077351625", "note": "vendor"},
        ),
        (
            "070426BB4552144A-AMBITION COLONISERS-4114135000006375-master to free",
            {"party": "AMBITION COLONISERS", "payment_mode": None,
             "reference": "070426BB4552144A", "account_number": "4114135000006375", "note": "master to free"},
        ),
        (
            "KVBLH00259483410-Pooja-3572001700041842-salary",
            {"party": "Pooja", "payment_mode": None,
             "reference": "KVBLH00259483410", "account_number": "3572001700041842", "note": "salary"},
        ),
        # No real UPI row exists in the workbook — this example is
        # synthetic, built on the IMPS-style structural convention, and
        # exists only to demonstrate the (documented, not-invented-as-real)
        # UPI handler.
        (
            "UPI-412300998877-RAVI KUMAR-SBIN-ravikumar@oksbi-UPI transaction",
            {"party": "RAVI KUMAR", "payment_mode": "UPI",
             "reference": "412300998877", "account_number": "ravikumar@oksbi", "note": "UPI transaction"},
        ),
        (
            "B/F...",
            {"party": None, "payment_mode": None, "reference": None, "account_number": None, "note": None},
        ),
        (
            "Monthly Service Chrgs MAY/26",
            {"party": None, "payment_mode": None, "reference": None, "account_number": None, "note": None},
        ),
        (
            "",
            {"party": None, "payment_mode": None, "reference": None, "account_number": None, "note": None},
        ),
    ]

    passed = 0
    for description, expected in test_cases:
        result = parse_description(description)
        status = "PASS" if result == expected else "FAIL"
        if status == "PASS":
            passed += 1
        print(f"[{status}] {description!r}")
        if status == "FAIL":
            print(f"    got:      {result}")
            print(f"    expected: {expected}")

    print(f"\n{passed}/{len(test_cases)} test cases passed.")
