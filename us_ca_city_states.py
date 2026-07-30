"""US/CA city to state/province mapping for postal code zone formatting."""

from __future__ import annotations

US_CITY_STATE: dict[str, str] = {
    "Clayton": "",
    "Cupertino": "CA",
    "Fontana": "CA",
    "Fremont": "CA",
    "Fullerton": "CA",
    "Houston": "TX",
    "Jeffersonville": "IN",
    "Los Angeles": "CA",
    "Memphis": "TN",
    "Northampton": "PA",
    "Northamption": "PA",
    "Oakland": "CA",
    "Ontario": "CA",
    "San Francisco": "CA",
    "Sparks": "NV",
}

COMMON_RATING_CITY_STATE: dict[str, str] = {
    "Austin": "TX",
    "Carlisle": "PA",
    "Elk Grove": "CA",
    "Indianapolis": "IN",
    "Lebanon": "TN",
    "Plainfield": "NJ",
    "Rialto": "CA",
}

CLAYTON_STATE_BY_ZONE: dict[str, str] = {
    "Indianapolis": "IN",
    "Plainfield": "NJ",
}


def state_code_for_city(city: str, country: str, common_rating_city: str) -> str:
    if country not in {"US", "CA"}:
        return ""

    city_text = city.strip()
    zone_text = common_rating_city.strip()
    mapped = US_CITY_STATE.get(city_text)
    if mapped:
        return mapped
    if city_text == "Clayton":
        return CLAYTON_STATE_BY_ZONE.get(zone_text, COMMON_RATING_CITY_STATE.get(zone_text, ""))
    return COMMON_RATING_CITY_STATE.get(zone_text, "")


def format_us_ca_postal_city(city: str, country: str, common_rating_city: str) -> str:
    if country not in {"US", "CA"}:
        return city.strip()

    state = state_code_for_city(city, country, common_rating_city)
    city_part = city.strip().replace(" ", "")
    if not state:
        return city.strip()
    return f"{state}_{city_part}"
