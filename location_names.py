"""Normalize and match location names (postal zones, city aliases, unicode variants)."""

from __future__ import annotations

import re
import unicodedata

# Canonical key -> accepted alternate spellings (all lowercased match keys).
LOCATION_ALIASES: dict[str, frozenset[str]] = {
    "xian": frozenset({"xian", "xi'an", "xi'an", "xi an"}),
    "viet yen": frozenset({"viet yen", "việt yên", "viet yen", "việt yên"}),
}

_APOSTROPHE_VARIANTS = re.compile(r"[''`´ʹʼ]")


def normalize_location_text(value: object) -> str:
    """Fold case, strip accents, unify apostrophes and whitespace."""
    text = str(value).strip() if value is not None else ""
    if not text or text.lower() == "nan":
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = _APOSTROPHE_VARIANTS.sub("'", text)
    text = re.sub(r"\s+", " ", text)
    return text.casefold().strip()


def location_match_key(value: object) -> str:
    """Single key used for zone dictionaries and deduplication."""
    normalized = normalize_location_text(value)
    if not normalized:
        return ""
    for _canonical, variants in LOCATION_ALIASES.items():
        if normalized in variants:
            return _canonical
    return normalized


def locations_equivalent(left: object, right: object) -> bool:
    return location_match_key(left) == location_match_key(right)


def register_alias(canonical: str, *aliases: str) -> None:
    """Runtime extension for tests or carrier-specific tables."""
    key = location_match_key(canonical)
    bucket = set(LOCATION_ALIASES.get(key, frozenset({key})))
    bucket.add(key)
    for alias in aliases:
        bucket.add(location_match_key(alias))
    LOCATION_ALIASES[key] = frozenset(bucket)
