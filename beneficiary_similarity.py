"""Shared near-duplicate name matching for Beneficiary Master entries.

The matching rule here is ported unchanged from
scripts/maintenance/build_amb_beneficiary_master.py's _fuzzy_merge() -
that script proved this rule (prefix match, or >=92% sequence
similarity, with corporate-suffix normalization) against real AMB
ledger data. This module exists so the SAME rule can also run live, in
classify_transactions.py's _update_beneficiary_master(), instead of
only ever running as a manual one-off seeding script for one company.

Deliberately conservative: this only ever answers "is there an
existing name close enough that a human should double-check?" - it
never merges/deletes/renames anything itself. Two genuinely different
people who happen to have similar names are exactly the failure mode
this must avoid causing silently, so callers must always treat a match
here as a prompt for human review, never as grounds to auto-merge.
"""

from __future__ import annotations

import difflib
import re

# Corporate-suffix abbreviations collapsed to one token before comparing,
# so "LIMITED" vs "LTD" / "PRIVATE LIMITED" vs "PVT LTD" don't block an
# otherwise-identical name from matching.
_SUFFIX_EQUIV = [
    (r"\bPRIVATE LIMITED\b", "PVTLTD"),
    (r"\bPVT\.?\s*LTD\.?\b", "PVTLTD"),
    (r"\bLIMITED\b", "LTD"),
]

_SIMILARITY_THRESHOLD = 0.92
_MIN_COMPARABLE_LENGTH = 5


def _fuzzy_key(name: str) -> str:
    text = name.upper()
    for pattern, repl in _SUFFIX_EQUIV:
        text = re.sub(pattern, repl, text)
    return re.sub(r"[^A-Z0-9]", "", text)


def is_fuzzy_duplicate(name_a: str, name_b: str) -> bool:
    """True if name_a/name_b are close enough to plausibly be the same
    beneficiary — one is a truncated prefix of the other, or their
    sequences are >=92% similar (checked both raw and with corporate
    suffixes normalized). Case-insensitive; whitespace/punctuation is
    ignored (so "A N FILLING STA TION" vs "A N FILLING STATION" and
    "A FILLING STATION" vs "A N FILLING STATION" both match)."""
    # Callers (e.g. classify_transactions.py's _update_beneficiary_master())
    # treat two names as "exact match" only after .strip().upper() - NOT
    # after stripping internal whitespace/punctuation. So "A N FILLING
    # STA TION" vs "A N FILLING STATION" differ only by an internal space
    # and are NOT an exact match by that caller's own rule - this function
    # must still flag them as a fuzzy duplicate, not silently skip them as
    # "already exact", or they'd fall through both checks entirely.
    if name_a.strip().upper() == name_b.strip().upper():
        return False  # truly identical strings - the caller's own exact-match check already handles this

    a_raw = re.sub(r"[^A-Z0-9]", "", name_a.upper())
    b_raw = re.sub(r"[^A-Z0-9]", "", name_b.upper())
    if len(a_raw) < _MIN_COMPARABLE_LENGTH or len(b_raw) < _MIN_COMPARABLE_LENGTH:
        return False

    a_norm, b_norm = _fuzzy_key(name_a), _fuzzy_key(name_b)
    is_prefix = (
        a_raw.startswith(b_raw) or b_raw.startswith(a_raw)
        or a_norm.startswith(b_norm) or b_norm.startswith(a_norm)
    )
    ratio = max(
        difflib.SequenceMatcher(None, a_raw, b_raw).ratio(),
        difflib.SequenceMatcher(None, a_norm, b_norm).ratio(),
    )
    return is_prefix or ratio >= _SIMILARITY_THRESHOLD


def find_similar_name(name: str, candidates: list[str]) -> str | None:
    """Return the first candidate in *candidates* that's a fuzzy
    duplicate of *name* (see is_fuzzy_duplicate()), or None if no
    candidate is close enough. *candidates* should exclude *name*
    itself (an exact match is a different, non-fuzzy case)."""
    for candidate in candidates:
        if is_fuzzy_duplicate(name, candidate):
            return candidate
    return None
