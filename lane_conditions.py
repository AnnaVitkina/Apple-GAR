"""Lane trigger fields (Origin/Destination) and APAC CN/TW conditional duplication."""

from __future__ import annotations

import copy

from carrier_rules import ApacCnTwPolicy, CarrierProfile, _cell_text, apac_cn_tw_policy
from country_codes import to_country_iso
from postal_constants import BULK_COUNTRY_POSTAL_CODE

APAC_CARRIER_NAME = "APAC"
NOT_APAC_CARRIER_NAME = "not APAC"
CN_TW_CONDITIONAL_ORIGIN = "CN, TW"

ATA_ATD_LANE_TYPES = frozenset({"ATA", "ATD"})

# Americas ISO 3166-1 (partial list + common GAR lanes); excludes US and CA for LATAM bulk rule.
AMERICAS_ISO = frozenset(
    {
        "AG",
        "AR",
        "AW",
        "BB",
        "BL",
        "BM",
        "BO",
        "BQ",
        "BR",
        "BS",
        "BZ",
        "CL",
        "CO",
        "CR",
        "CU",
        "CW",
        "DM",
        "DO",
        "EC",
        "FK",
        "GD",
        "GF",
        "GP",
        "GT",
        "GY",
        "HN",
        "HT",
        "JM",
        "KN",
        "KY",
        "LC",
        "MF",
        "MQ",
        "MS",
        "MX",
        "NI",
        "PA",
        "PE",
        "PR",
        "PY",
        "SR",
        "SV",
        "SX",
        "TC",
        "TT",
        "UY",
        "VE",
        "VG",
        "VI",
    }
)

HAN_ORIGIN_TOKENS = frozenset({"hanoi", "han"})
SGN_ORIGIN_TOKENS = frozenset(
    {"ho chi minh city", "ho chi minh", "hcmc", "sgn", "saigon"}
)

TAIPEI_TOKENS = frozenset({"taipei", "tpe"})


def _bulk_country_zone_value() -> str:
    return BULK_COUNTRY_POSTAL_CODE


def _combined_origin_text(row: dict[str, object]) -> str:
    parts = (
        _cell_text(row.get("Origin City")),
        _cell_text(row.get("Origin City Label")),
        _cell_text(row.get("Uplift Airport")),
        _cell_text(row.get("Origin Airport")),
    )
    return " ".join(p for p in parts if p).casefold()


def is_taipei_origin(row: dict[str, object]) -> bool:
    text = _combined_origin_text(row)
    return any(token in text for token in TAIPEI_TOKENS)


def _vn_origin_airport(row: dict[str, object]) -> str:
    text = _combined_origin_text(row)
    if any(token in text for token in HAN_ORIGIN_TOKENS):
        return "HAN"
    if any(token in text for token in SGN_ORIGIN_TOKENS):
        return "SGN"
    return ""


def is_latam_origin_country(country: object) -> bool:
    iso = to_country_iso(country)
    return bool(iso) and iso in AMERICAS_ISO


def _lane_type(row: dict[str, object]) -> str:
    return _cell_text(row.get("Lane Type")).upper()


def _is_od_row(row: dict[str, object]) -> bool:
    return bool(_cell_text(row.get("Measurement Type")))


def apply_origin_and_conditional_fields(matrix_rows: list[dict[str, object]]) -> None:
    """
    Set Origin City trigger, Conditional Origin City, and related fields per GAR BP.
    """
    for row in matrix_rows:
        label = _cell_text(row.get("Origin City Label"))
        uplift = _cell_text(row.get("Uplift Airport")) or _cell_text(row.get("Origin Airport"))
        origin_iso = to_country_iso(row.get("Origin Country"))
        lane_type = _lane_type(row)

        row.setdefault("Conditional Origin Country", "")
        row.setdefault("Conditional Origin City", "")
        row.setdefault("Conditional Destination", "")

        row["Conditional Origin City"] = label

        if origin_iso == "VN":
            airport = _vn_origin_airport(row)
            if airport:
                row["Origin City"] = airport
                row["Conditional Origin City"] = airport if lane_type in ATA_ATD_LANE_TYPES else label
            continue

        if is_latam_origin_country(origin_iso):
            bulk = _bulk_country_zone_value()
            row["Origin City"] = bulk
            if lane_type in ATA_ATD_LANE_TYPES:
                row["Conditional Origin City"] = uplift or bulk
            continue

        if lane_type not in ATA_ATD_LANE_TYPES:
            if not _cell_text(row.get("Origin City")):
                row["Origin City"] = uplift or label
            continue

        if uplift:
            row["Origin City"] = uplift
            row["Conditional Origin City"] = uplift
            continue

        if origin_iso == "US" and _is_od_row(row):
            row["Origin City"] = label
            row["Conditional Origin City"] = label
            continue

        bulk = _bulk_country_zone_value()
        row["Origin City"] = bulk
        row["Conditional Origin City"] = bulk


def matches_cn_tw_ata_atd_lane(row: dict[str, object]) -> bool:
    origin_iso = to_country_iso(row.get("Origin Country"))
    if origin_iso not in {"CN", "TW"}:
        return False
    if is_taipei_origin(row):
        return False
    return _lane_type(row) in ATA_ATD_LANE_TYPES


def apply_apac_cn_tw_conditionals(
    matrix_rows: list[dict[str, object]],
    *,
    carrier: CarrierProfile,
    source_stem: str = "",
) -> int:
    """APAC-only vs mixed-carrier CN/TW ATA/ATD conditional handling."""
    policy = apac_cn_tw_policy(source_stem, carrier)
    if policy is ApacCnTwPolicy.NOT_APPLICABLE:
        return 0

    apac_only = policy is ApacCnTwPolicy.APAC_ONLY
    touched = 0
    additions: list[dict[str, object]] = []

    for row in matrix_rows:
        if not matches_cn_tw_ata_atd_lane(row):
            continue
        if _cell_text(row.get("Carrier name")) in {APAC_CARRIER_NAME, NOT_APAC_CARRIER_NAME}:
            continue

        row["Conditional Origin Country"] = CN_TW_CONDITIONAL_ORIGIN
        touched += 1

        if apac_only:
            continue

        row["Carrier name"] = APAC_CARRIER_NAME
        copy_row = copy.deepcopy(row)
        copy_row["Conditional Origin Country"] = ""
        copy_row["Carrier name"] = NOT_APAC_CARRIER_NAME
        additions.append(copy_row)
        touched += 1

    matrix_rows.extend(additions)
    return touched


def apply_conditional_destination_from_trigger(row: dict[str, object]) -> None:
    """Set Conditional Destination from destination trigger (incl. bulk alphanumeric zone)."""
    if _cell_text(row.get("Conditional Destination")):
        return
    dest = _cell_text(row.get("Destination"))
    label = _cell_text(row.get("Destination Label"))
    if dest == BULK_COUNTRY_POSTAL_CODE or dest.startswith("0, 1, 2"):
        row["Conditional Destination"] = BULK_COUNTRY_POSTAL_CODE
    elif dest:
        row["Conditional Destination"] = dest
    elif label:
        row["Conditional Destination"] = label
