"""Carrier-specific matrix rules (destination airport, SCH VN, zone copies, etc.)."""

from __future__ import annotations

import copy
import re
from datetime import date, datetime
from enum import Enum

import pandas as pd

from country_codes import to_country_iso
from postal_constants import BULK_COUNTRY_POSTAL_CODE

ADDITIONAL_LANE_PATTERN = re.compile(r"\(additional\s+lane\)", re.IGNORECASE)
ALTERNATIVE_GATEWAY_SUFFIX = " (Alternative Gateway)"

EMEA_CARRIERS_LABEL = "EMEA carriers"
NOT_EMEA_CARRIERS_LABEL = "not EMEA carriers"

SCH_ZONE_DESTINATION_REPLACEMENTS = (
    ("GAR-US Zone 5", "York Springs"),
    ("GAR-US Zone 6", "Miami Shores"),
)

CEVA_DGF_DESTINATION_AIRPORT: dict[str, str] = {
    "mumbai": "BOM",
    "kolkata": "CCU",
    "bengaluru": "BLR",
    "bangalore": "BLR",
    "delhi": "DEL",
    "bogota": "BOG",
    "bogotá": "BOG",
    "mexico city": "MEX",
    "toluca": "TLC",
    "lima": "LIM",
}

KWE_DESTINATION_AIRPORT: dict[str, str] = {
    "tokyo": "NRT",
    "osaka": "KIX",
}

HAN_CITY_TOKENS = frozenset({"hanoi", "han"})
SGN_CITY_TOKENS = frozenset({"ho chi minh city", "ho chi minh", "hcmc", "sgn"})


class CarrierProfile(str, Enum):
    CEVA = "CEVA"
    DGF = "DGF"
    SCH = "SCH"
    DSV = "DSV"
    FTN = "FTN"
    KN = "KN"
    KWE = "KWE"
    NIPPON = "NIPPON"
    APEX = "APEX"
    GEO = "GEO"
    MEC = "MEC"
    EXPD = "EXPD"
    GENERIC = "GENERIC"


class ApacCnTwPolicy(str, Enum):
    """CN/TW ATA/ATD conditional Origin Country rules."""

    NOT_APPLICABLE = "not_applicable"
    APAC_ONLY = "apac_only"
    MIXED = "mixed"


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _format_display_date(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, datetime):
        return value.date().strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    text = _cell_text(value)
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text[:10], fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return text


def detect_carrier(source_stem: str) -> CarrierProfile:
    upper = source_stem.upper().replace("&", " ")
    if "CEVA" in upper:
        return CarrierProfile.CEVA
    if "DGF" in upper:
        return CarrierProfile.DGF
    if "SCHENKER" in upper or re.search(r"\bSCH\b", upper):
        return CarrierProfile.SCH
    if "DSV" in upper:
        return CarrierProfile.DSV
    if "FTN" in upper or "FEDEX" in upper:
        return CarrierProfile.FTN
    if "K&N" in upper or "K N" in upper or " KN" in f" {upper}" or upper.startswith("KN "):
        return CarrierProfile.KN
    if "KWE" in upper:
        return CarrierProfile.KWE
    if "NIPPON" in upper:
        return CarrierProfile.NIPPON
    if "APEX" in upper:
        return CarrierProfile.APEX
    if "GEO" in upper and "GAR" in upper:
        return CarrierProfile.GEO
    if "MEC" in upper:
        return CarrierProfile.MEC
    if "EXPD" in upper or "EXPEDITORS" in upper:
        return CarrierProfile.EXPD
    return CarrierProfile.GENERIC


def apac_cn_tw_policy(_source_stem: str, carrier: CarrierProfile) -> ApacCnTwPolicy:
    if carrier in {
        CarrierProfile.KN,
        CarrierProfile.FTN,
        CarrierProfile.CEVA,
        CarrierProfile.GEO,
    }:
        return ApacCnTwPolicy.NOT_APPLICABLE
    if carrier in {CarrierProfile.KWE, CarrierProfile.NIPPON}:
        return ApacCnTwPolicy.APAC_ONLY
    if carrier in {
        CarrierProfile.DGF,
        CarrierProfile.APEX,
        CarrierProfile.MEC,
        CarrierProfile.DSV,
        CarrierProfile.EXPD,
    }:
        return ApacCnTwPolicy.MIXED
    return ApacCnTwPolicy.NOT_APPLICABLE


def carrier_allows_alternative_gateway(carrier: CarrierProfile) -> bool:
    return carrier is not CarrierProfile.CEVA


def carrier_includes_od_alternative_gateway(carrier: CarrierProfile) -> bool:
    return carrier is CarrierProfile.FTN


def is_additional_lane_tab(tab: object) -> bool:
    return bool(ADDITIONAL_LANE_PATTERN.search(_cell_text(tab)))


def remove_additional_lane_rows(matrix_rows: list[dict[str, object]]) -> int:
    before = len(matrix_rows)
    kept = [row for row in matrix_rows if not is_additional_lane_tab(row.get("Tab"))]
    removed = before - len(kept)
    matrix_rows[:] = kept
    return removed


def lookup_destination_airport(destination_city: str, mapping: dict[str, str]) -> str:
    key = destination_city.casefold().strip()
    if not key:
        return ""
    if key in mapping:
        return mapping[key]
    for name, iata in mapping.items():
        if name in key or key in name:
            return iata
    return ""


def lookup_ceva_dgf_airport(destination_city: str) -> str:
    return lookup_destination_airport(destination_city, CEVA_DGF_DESTINATION_AIRPORT)


def _bulk_country_destination(_row: dict[str, object]) -> str:
    return BULK_COUNTRY_POSTAL_CODE


def apply_destination_airports(matrix_rows: list[dict[str, object]], carrier: CarrierProfile) -> int:
    updated = 0
    for row in matrix_rows:
        label = _cell_text(row.get("Destination Label"))
        dest = _cell_text(row.get("Destination"))
        city = label or dest

        if carrier in {CarrierProfile.CEVA, CarrierProfile.DGF}:
            airport = lookup_ceva_dgf_airport(city)
            if airport:
                row["Destination Airport"] = airport
                bulk = _bulk_country_destination(row)
                row["Destination"] = bulk
                row["Conditional Destination"] = bulk
                updated += 1
            continue

        if carrier is CarrierProfile.KWE:
            airport = lookup_destination_airport(city, KWE_DESTINATION_AIRPORT)
            if airport:
                row["Destination Airport"] = airport
                row["Destination"] = airport
                row["Destination Label"] = city
                updated += 1
            continue

        if carrier is CarrierProfile.KN:
            dest_country = to_country_iso(row.get("Destination Country"))
            lane_type = _cell_text(row.get("Lane Type")).upper()
            if dest_country == "SA" and lane_type in {"DTA", "ATA"}:
                row["Destination Airport"] = "RUH"
                row["Destination"] = "RUH"
                row["Destination Label"] = city or "RUH"
                updated += 1

    return updated


def _as_sortable_date(value: object):
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def compute_lane_valid_to(
    source_row: pd.Series,
    *,
    fuel_row: pd.Series | None,
    transport_periods: list,
    fuel_periods: list,
    has_fuel_on_lane: bool,
) -> str:
    transport_dates: list[object] = []
    for period in transport_periods:
        col = f"Valid To{period.suffix}" if period.suffix else "Valid To"
        if col in source_row.index:
            value = source_row.get(col)
            if not pd.isna(value) and _cell_text(value):
                transport_dates.append(value)

    if has_fuel_on_lane and fuel_row is not None and fuel_periods:
        fuel_dates: list[object] = []
        for period in fuel_periods:
            col = f"Valid To{period.suffix}" if period.suffix else "Valid To"
            if col in fuel_row.index:
                value = fuel_row.get(col)
                if not pd.isna(value) and _cell_text(value):
                    fuel_dates.append(value)
        if fuel_dates:
            return _format_display_date(max(fuel_dates, key=_as_sortable_date))

    if transport_dates:
        return _format_display_date(max(transport_dates, key=_as_sortable_date))
    return _format_display_date(source_row.get("Valid To_5") or source_row.get("Valid To"))


def _vn_destination_airport_and_code(row: dict[str, object]) -> tuple[str, str] | None:
    dest_label = _cell_text(row.get("Destination Label")).casefold()
    dest_code = _cell_text(row.get("Destination")).casefold()
    combined = f"{dest_label} {dest_code}"
    if any(token in combined for token in HAN_CITY_TOKENS):
        return "HAN", "HAN"
    if any(token in combined for token in SGN_CITY_TOKENS):
        return "SGN", "SGN"
    return None


def append_sch_vn_carrier_rows(matrix_rows: list[dict[str, object]], carrier: CarrierProfile) -> int:
    if carrier is not CarrierProfile.SCH:
        return 0

    additions: list[dict[str, object]] = []
    for row in matrix_rows:
        if to_country_iso(row.get("Destination Country")) != "VN":
            continue
        if ALTERNATIVE_GATEWAY_SUFFIX in _cell_text(row.get("Tab")):
            continue
        if _cell_text(row.get("Carrier name")):
            continue

        row["Carrier name"] = EMEA_CARRIERS_LABEL

        not_emea = copy.deepcopy(row)
        not_emea["Carrier name"] = NOT_EMEA_CARRIERS_LABEL
        airport_code = _vn_destination_airport_and_code(row)
        if airport_code:
            iata, dest_code = airport_code
            not_emea["Destination Airport"] = iata
            not_emea["Destination"] = dest_code
            not_emea["Destination Label"] = dest_code
            not_emea["Conditional Destination"] = dest_code
        additions.append(not_emea)

    matrix_rows.extend(additions)
    return len(additions)


def _matches_zone_destination(row: dict[str, object], zone_name: str) -> bool:
    dest = _cell_text(row.get("Destination"))
    label = _cell_text(row.get("Destination Label"))
    zone_cf = zone_name.casefold()
    return dest.casefold() == zone_cf or label.casefold() == zone_cf


def append_schenker_us_zone_destination_rows(
    matrix_rows: list[dict[str, object]],
    carrier: CarrierProfile,
) -> int:
    if carrier is not CarrierProfile.SCH:
        return 0

    additions: list[dict[str, object]] = []
    for row in matrix_rows:
        if ALTERNATIVE_GATEWAY_SUFFIX in _cell_text(row.get("Tab")):
            continue
        for zone_name, new_dest in SCH_ZONE_DESTINATION_REPLACEMENTS:
            if not _matches_zone_destination(row, zone_name):
                continue
            copy_row = copy.deepcopy(row)
            copy_row["Destination"] = new_dest
            copy_row["Destination Label"] = new_dest
            copy_row["Conditional Destination"] = new_dest
            additions.append(copy_row)

    matrix_rows.extend(additions)
    return len(additions)
