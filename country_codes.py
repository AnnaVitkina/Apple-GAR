"""Country name and alias to ISO 3166-1 alpha-2 conversion for GAR matrix output."""

from __future__ import annotations

COUNTRY_TO_ISO: dict[str, str] = {
    "AE": "AE",
    "AU": "AU",
    "AUSTRALIA": "AU",
    "Australia": "AU",
    "CN": "CN",
    "CHINA": "CN",
    "CZ": "CZ",
    "CZECH R.": "CZ",
    "Czech R.": "CZ",
    "CZECH REPUBLIC": "CZ",
    "GB": "GB",
    "UK": "GB",
    "UNITED KINGDOM": "GB",
    "IE": "IE",
    "IRELAND": "IE",
    "IN": "IN",
    "INDIA": "IN",
    "IT": "IT",
    "ITALY": "IT",
    "Italy": "IT",
    "JP": "JP",
    "JAPAN": "JP",
    "KR": "KR",
    "KOREA": "KR",
    "NL": "NL",
    "NETHERLANDS": "NL",
    "Netherlands": "NL",
    "SA": "SA",
    "SAUDI ARABIA": "SA",
    "SG": "SG",
    "SINGAPORE": "SG",
    "TW": "TW",
    "TAIWAN": "TW",
    "TR": "TR",
    "TURKEY": "TR",
    "Turkey": "TR",
    "VN": "VN",
    "VIETNAM": "VN",
    "US": "US",
    "USA": "US",
    "UNITED STATES": "US",
    "DE": "DE",
    "GERMANY": "DE",
    "FR": "FR",
    "FRANCE": "FR",
    "ES": "ES",
    "SPAIN": "ES",
    "BE": "BE",
    "BELGIUM": "BE",
    "HU": "HU",
    "HUNGARY": "HU",
    "PL": "PL",
    "POLAND": "PL",
    "AT": "AT",
    "AUSTRIA": "AT",
    "CH": "CH",
    "SWITZERLAND": "CH",
    "SE": "SE",
    "SWEDEN": "SE",
    "NO": "NO",
    "NORWAY": "NO",
    "DK": "DK",
    "DENMARK": "DK",
    "FI": "FI",
    "FINLAND": "FI",
    "PT": "PT",
    "PORTUGAL": "PT",
    "GR": "GR",
    "GREECE": "GR",
    "RO": "RO",
    "ROMANIA": "RO",
    "BG": "BG",
    "BULGARIA": "BG",
    "HR": "HR",
    "CROATIA": "HR",
    "SK": "SK",
    "SLOVAKIA": "SK",
    "SI": "SI",
    "SLOVENIA": "SI",
    "LT": "LT",
    "LITHUANIA": "LT",
    "LV": "LV",
    "LATVIA": "LV",
    "EE": "EE",
    "ESTONIA": "EE",
    "LU": "LU",
    "LUXEMBOURG": "LU",
    "MX": "MX",
    "MEXICO": "MX",
    "CO": "CO",
    "COLOMBIA": "CO",
    "PE": "PE",
    "PERU": "PE",
    "BR": "BR",
    "BRAZIL": "BR",
    "MY": "MY",
    "MALAYSIA": "MY",
    "TH": "TH",
    "THAILAND": "TH",
    "ID": "ID",
    "INDONESIA": "ID",
    "PH": "PH",
    "PHILIPPINES": "PH",
    "NZ": "NZ",
    "NEW ZEALAND": "NZ",
    "ZA": "ZA",
    "SOUTH AFRICA": "ZA",
    "IL": "IL",
    "ISRAEL": "IL",
    "EG": "EG",
    "EGYPT": "EG",
    "QA": "QA",
    "QATAR": "QA",
    "KW": "KW",
    "KUWAIT": "KW",
    "BH": "BH",
    "BAHRAIN": "BH",
    "OM": "OM",
    "OMAN": "OM",
    "JO": "JO",
    "JORDAN": "JO",
    "LB": "LB",
    "LEBANON": "LB",
    "RU": "RU",
    "RUSSIA": "RU",
    "UA": "UA",
    "UKRAINE": "UA",
    "CA": "CA",
    "CANADA": "CA",
}


def to_country_iso(value: object) -> str:
    if value is None or (isinstance(value, float) and str(value) == "nan"):
        return ""

    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""

    if len(text) == 2 and text.isalpha():
        return text.upper()

    if text in COUNTRY_TO_ISO:
        return COUNTRY_TO_ISO[text]

    upper = text.upper()
    if upper in COUNTRY_TO_ISO:
        return COUNTRY_TO_ISO[upper]

    return text


def iso_to_bulk_country_label(iso: str) -> str:
    """Preferred display name for country-as-zone triggers (bulk country lanes)."""
    if not iso:
        return ""
    code = iso.upper()
    best = ""
    for name, mapped in COUNTRY_TO_ISO.items():
        if mapped != code:
            continue
        if len(name) <= 2:
            continue
        if name.isalpha() and name[0].isupper() and len(name) > len(best):
            best = name
    if best:
        return best
    for name, mapped in COUNTRY_TO_ISO.items():
        if mapped == code and len(name) > 2:
            return name.title()
    return code
