"""Normalize and match location names (postal zones, city aliases, unicode variants)."""

from __future__ import annotations

import re
import unicodedata

# Canonical key -> accepted alternate spellings (all lowercased match keys).
LOCATION_ALIASES: dict[str, frozenset[str]] = {
    "xian": frozenset({"xian", "xi'an", "xi'an", "xi an"}),
    "viet yen": frozenset({"viet yen", "việt yên", "viet yen", "việt yên"}),
    "bengaluru": frozenset({"bengaluru", "bangalore"}),
}

# Suffixes that match both a US state code and an ISO country code.
FOREIGN_CITIES_BY_AMBIGUOUS_SUFFIX: dict[str, frozenset[str]] = {
    "IN": frozenset(
        {
            "bengaluru",
            "chennai",
            "mumbai",
            "kolkata",
            "delhi",
            "mahindraworld",
            "sunguvarchatram",
            "tamilnadu",
            "chengalpattu",
            "sriperumbudur",
            "kanchipuram",
            "thane",
            "kolar",
            "karnataka",
        }
    ),
    "ID": frozenset({"jakarta"}),
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


def foreign_country_for_ambiguous_suffix(city: str, suffix: str) -> str | None:
    """Return ISO country when a city suffix doubles as a US state code (e.g. Chennai, IN)."""
    suffix_upper = suffix.strip().upper()
    foreign_cities = FOREIGN_CITIES_BY_AMBIGUOUS_SUFFIX.get(suffix_upper)
    if not foreign_cities:
        return None
    if location_match_key(city) in foreign_cities:
        return suffix_upper
    return None


def is_misclassified_foreign_us_city(city: str, region: str, country: str) -> bool:
    """True when a city was parsed as US/CA but suffix indicates a foreign country."""
    if country not in {"US", "CA"} or not region:
        return False
    return foreign_country_for_ambiguous_suffix(city, region) is not None


BOILERPLATE_CLUSTER_MARKERS = (
    "the metro area and surrounding communities",
    "all other",
    "lanes (except",
    "only for destination",
    "all except destination",
    "to specify the city",
    "hub & direct ship",
    "low volume order only",
)

CROSS_BORDER_CLUSTER_EXCLUSIONS: dict[str, frozenset[str]] = {
    "ho chi minh city": frozenset({"hong kong", "kowloon"}),
}


def is_boilerplate_cluster_member(text: str) -> bool:
    lower = normalize_location_text(text)
    if not lower:
        return True
    return any(marker in lower for marker in BOILERPLATE_CLUSTER_MARKERS)


def is_excluded_cluster_member(cluster_name: str, member_text: str) -> bool:
    if is_boilerplate_cluster_member(member_text):
        return True
    member_key = location_match_key(member_text)
    blocked = CROSS_BORDER_CLUSTER_EXCLUSIONS.get(location_match_key(cluster_name), frozenset())
    return member_key in blocked


def register_alias(canonical: str, *aliases: str) -> None:
    """Runtime extension for tests or carrier-specific tables."""
    key = location_match_key(canonical)
    bucket = set(LOCATION_ALIASES.get(key, frozenset({key})))
    bucket.add(key)
    for alias in aliases:
        bucket.add(location_match_key(alias))
    LOCATION_ALIASES[key] = frozenset(bucket)
