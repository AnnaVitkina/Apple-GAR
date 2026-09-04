"""
Build a GAR matrix workbook (Rate card sheet) from extracted processing data.

Target layout follows Advanced Export:
  - Shipment columns (lane identity, service product, measurement type)
  - Transport cost groups (5 peak/non-peak validity periods)
  - Optional fuel surcharge columns when Fuel Rates sheet is present
  - OD handling rows (ODH / OD2 / OD3 / OD4) from carrier OD tabs
"""

from __future__ import annotations

import copy
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from convert_to_processing import prompt_selection
from country_codes import to_country_iso
from lane_conditions import (
    apply_apac_cn_tw_conditionals,
    apply_conditional_destination_from_trigger,
    apply_origin_and_conditional_fields,
)
from carrier_rules import (
    CarrierProfile,
    append_schenker_us_zone_destination_rows,
    append_sch_vn_carrier_rows,
    apply_destination_airports,
    carrier_allows_alternative_gateway,
    carrier_includes_od_alternative_gateway,
    compute_lane_valid_to,
    detect_carrier,
    rc_valid_to_from_dataframe,
    remove_additional_lane_rows,
)
from project_paths import OUTPUT_DIR, PROCESSING_DIR, ensure_workspace_dirs

EXTRACTED_GLOB = "*_extracted.xlsx"
BASE_FREIGHT_TAB = "base freight rates"
FUEL_RATES_TAB = "fuel rates"
VALUE_COLUMN_PATTERN = re.compile(r"^Value\s*\(USD\)(?:_\d+)?$", re.IGNORECASE)
OD_VALUE_COLUMN = "O/D Total Value (USD)"
ITEM_ID_COLUMN = "2026 Item_ID"

SERVICE_CODES = {
    "standard": "STD",
    "deferred": "DEF",
    "express": "EXP",
    "dg": "DGR",
}
RATE_TYPE_CODES = {
    "hub/direct": "HUB",
    "direct": "DIR",
    "hub": "HUB",
    "ac_hub": "HUB",
    "ac hub": "HUB",
}

FSC_PERIOD_TO_GROUP = {
    "apr part 1": "Fuel Surcharge (April 2026 1st part)",
    "apr part 2": "Fuel Surcharge (April 2026 2nd part)",
    "may part 1": "Fuel Surcharge (May 2026 1st part)",
    "may part 2": "Fuel Surcharge (May 2026 2nd part)",
    "june part 1": "Fuel Surcharge (June 2026 1st part)",
    "june part 2": "Fuel Surcharge (June 2026 2nd part)",
    "july part 1": "Fuel Surcharge (July 2026 1st part)",
    "july part 2": "Fuel Surcharge (July 2026 2nd part)",
    "jun part 1": "Fuel Surcharge (June 2026 1st part)",
    "jun part 2": "Fuel Surcharge (June 2026 2nd part)",
    "jul part 1": "Fuel Surcharge (July 2026 1st part)",
    "jul part 2": "Fuel Surcharge (July 2026 2nd part)",
    "aug part 1": "Fuel Surcharge (August 2026 1st part)",
}

OD_TAB_MEASUREMENT: dict[str, str] = {
    "od total rates": "ODH",
    "od rates total 2026": "ODH",
    "qr blr-ord od total rates": "ODH",
    "qr maa-ord od total rates": "OD2",
    "ca han-ord od total rates": "OD3",
    "fxl han-ord od total rates": "OD4",
}

OD_COST_GROUPS = {
    "ODH": "Origin and Destination Handling",
    "OD2": 'Additional Origin and Destination Handling, v2 (Tab "QR MAA-ORD OD Total Rates")',
    "OD3": 'Additional Origin and Destination Handling, v3 (Tab "CA HAN-ORD OD Total Rates")',
    "OD4": "Additional Origin and Destination Handling, v4 (Tab 'FXL HAN-ORD OD Total Rates')",
}

MATRIX_COLUMNS = (
    "Lane #",
    "Tab",
    "Origin Country",
    "Origin State",
    "Origin City",
    "Origin City Label",
    "Uplift Airport",
    "Origin Airport",
    "Rate Type",
    "Destination Country",
    "Destination State",
    "Destination",
    "Destination Label",
    "Destination Airport",
    "Country of Clearance",
    "Service",
    "Lane Type",
    "Lane Type Alt",
    "Service Product",
    "Measurement Type",
    "Business Segment",
    "Invoice type",
    "Valid to",
)

EXPORT_HEADERS = {
    "Origin City Label": "Origin City ",
    "Destination Label": "Destination ",
    "Lane Type Alt": "Lane Type",
    "Service Product": "Service",
}

# Shipment columns where RAM conditions apply (bold in Rate card export).
BOLD_CONDITION_COLUMNS = frozenset(
    {
        "Origin City",
        "Origin Airport",
        "Destination",
        "Lane Type",
        "Service Product",
        "Business Segment",
        "Measurement Type",
        "Invoice type",
        "Valid to",
    }
)

SHIPMENT_COLUMNS = MATRIX_COLUMNS
ORIGIN_CITY_LABEL_COLUMN = EXPORT_HEADERS["Origin City Label"]
DESTINATION_LABEL_COLUMN = EXPORT_HEADERS["Destination Label"]

TITLE_ROW = 1
AGREEMENT_ROW = 2
COST_GROUP_ROW = 3
VALIDITY_ROW = 4
RATE_BY_ROW = 5
CONDITIONAL_ROW = 6
COLUMN_HEADER_ROW = 7
DATA_START_ROW = 8

TRANSPORT_RATE_BY = "Rate by: Weight/chargeable kg\nRegular rule"
FUEL_RATE_BY = "Rate by: Weight/chargeable kg\nRegular rule"
OD_RATE_BY = "Rate by: Weight/chargeable kg\nRegular rule"
DEFAULT_APPLY_IF = "Applies if invoiced by Carrier"
DEFAULT_INVOICE_TYPE = "not TR"
ACCESSORIAL_SERVICE_SPECIAL = "SPECIAL"
DEFAULT_ACCESSORIAL_TAB = "Accessorial Tariff BOC26"
ALTERNATIVE_GATEWAY_SUFFIX = " (Alternative Gateway)"
ALTERNATIVE_GATEWAY_INVOICE_TYPE = "TR"

# APAC origin countries eligible for Alternative Gateway (ISO 3166-1 alpha-2).
APAC_ORIGIN_COUNTRIES = frozenset(
    {
        "AU",
        "CN",
        "HK",
        "ID",
        "IN",
        "JP",
        "KH",
        "KR",
        "LA",
        "MM",
        "MO",
        "MY",
        "NZ",
        "PH",
        "SG",
        "TH",
        "TW",
        "VN",
    }
)

# When uplift airport is empty, still duplicate and set Origin Airport manually.
MANUAL_ORIGIN_AIRPORT_BY_COUNTRY: dict[str, str] = {
    "JP": "NRT",
    "SG": "SIN",
    "ID": "BTH",
    "TW": "TPE",
}

DEFAULT_CURRENCY = "USD"

HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
TRANSPORT_GROUP_FILL = PatternFill("solid", fgColor="9BC2E6")
TRANSPORT_COST_FILL = PatternFill("solid", fgColor="BDD7EE")
FUEL_GROUP_FILL = PatternFill("solid", fgColor="C6E0B4")
OD_GROUP_FILL = PatternFill("solid", fgColor="FCE4D6")
COST_META_FILL = PatternFill("solid", fgColor="F2F2F2")
BOLD = Font(bold=True)
NORMAL = Font()
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RATE_NUMBER_FORMAT = "#,##0.00"
THIN_BORDER = Border(
    left=Side(style="thin", color="B4B4B4"),
    right=Side(style="thin", color="B4B4B4"),
    top=Side(style="thin", color="B4B4B4"),
    bottom=Side(style="thin", color="B4B4B4"),
)


@dataclass(frozen=True)
class PeriodColumn:
    period_label: str
    valid_from: object
    valid_to: object
    value_column: str
    suffix: str


@dataclass(frozen=True)
class CostSpec:
    group_title: str
    validity_text: str
    currency_column: str
    rate_column: str
    rate_by: str
    apply_if: str
    fill: PatternFill


def cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def rate_value(value: object) -> float | int | None:
    if pd.isna(value):
        return None
    text = cell_text(value)
    if not text:
        return None
    try:
        number = float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None
    if number.is_integer():
        return int(number)
    return number


def format_display_date(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, datetime):
        return value.date().strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    text = cell_text(value)
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text[:10], fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return text


def validity_text(valid_from: object, valid_to: object) -> str:
    start = format_display_date(valid_from)
    end = format_display_date(valid_to)
    if not start or not end:
        return DEFAULT_APPLY_IF
    return f"Validity period: from {start} to {end}"


def service_product(service: object, rate_type: object, lane_type: object) -> str:
    service_key = cell_text(service).lower()
    rate_key = cell_text(rate_type).lower().replace(" ", "/")
    lane_key = cell_text(lane_type).upper()
    service_code = SERVICE_CODES.get(service_key, service_key.upper()[:3] or "STD")
    rate_code = RATE_TYPE_CODES.get(rate_key, "HUB")
    return f"{service_code}_{rate_code}_{lane_key}"


def _normalize_rate_type_key(rate_type: object) -> str:
    return cell_text(rate_type).upper().replace(" ", "")


def is_ac_rate_type(rate_type: object) -> bool:
    return _normalize_rate_type_key(rate_type).startswith("AC_HUB")


def is_hub_style_rate_type(rate_type: object) -> bool:
    key = _normalize_rate_type_key(rate_type)
    if is_ac_rate_type(rate_type):
        return False
    return key in {"HUB", "HUB/DIRECT", "DIRECT"}


def _non_ac_segment_label(source_stem: str) -> str:
    """FTN Advanced Export uses 'not AC'; KN and BP default use 'FG'."""
    if "FTN" in source_stem.upper():
        return "not AC"
    return "FG"


def _segment_lane_key(row: dict[str, object]) -> tuple[str, ...]:
    return (
        cell_text(row.get("Origin Country")),
        cell_text(row.get("Origin City Label")),
        cell_text(row.get("Destination Country")),
        cell_text(row.get("Destination Label")),
        cell_text(row.get("Service")),
        cell_text(row.get("Lane Type")),
    )


def apply_business_segments(matrix_rows: list[dict[str, object]], *, source_stem: str) -> None:
    """
    Set Business Segment for base-freight rows (not OD):
      - AC_Hub* rate types -> AC
      - Hub/Direct (etc.) on the same lane as an AC row -> FG or 'not AC' (carrier)
    Most lanes stay empty, matching Advanced Export.
    """
    non_ac_label = _non_ac_segment_label(source_stem)
    ftn_style = "FTN" in source_stem.upper()
    freight_indices = [
        idx
        for idx, row in enumerate(matrix_rows)
        if not cell_text(row.get("Measurement Type"))
    ]

    groups: dict[tuple[str, ...], list[int]] = {}
    for idx in freight_indices:
        key = _segment_lane_key(matrix_rows[idx])
        groups.setdefault(key, []).append(idx)

    for idx in freight_indices:
        if is_ac_rate_type(matrix_rows[idx].get("Rate Type")):
            matrix_rows[idx]["Business Segment"] = "AC"

    for indices in groups.values():
        ac_indices = [idx for idx in indices if is_ac_rate_type(matrix_rows[idx].get("Rate Type"))]
        hub_indices = [idx for idx in indices if is_hub_style_rate_type(matrix_rows[idx].get("Rate Type"))]
        if not ac_indices or not hub_indices:
            continue
        for idx in ac_indices:
            matrix_rows[idx]["Business Segment"] = "AC"
        for idx in hub_indices:
            matrix_rows[idx]["Business Segment"] = non_ac_label

    if ftn_style:
        return

    for idx in freight_indices:
        if is_ac_rate_type(matrix_rows[idx].get("Rate Type")) and not cell_text(
            matrix_rows[idx].get("Business Segment")
        ):
            matrix_rows[idx]["Business Segment"] = "AC"


def _is_apac_origin_country(country: object) -> bool:
    iso = to_country_iso(country)
    return bool(iso) and iso in APAC_ORIGIN_COUNTRIES


def _alternative_gateway_airport(row: dict[str, object]) -> str:
    country = to_country_iso(row.get("Origin Country"))
    uplift = cell_text(row.get("Origin Airport")) or cell_text(row.get("Uplift Airport"))
    if uplift:
        return uplift
    return MANUAL_ORIGIN_AIRPORT_BY_COUNTRY.get(country, "")


def _ag_origin_country_ok(row: dict[str, object], carrier: CarrierProfile) -> bool:
    iso = to_country_iso(row.get("Origin Country"))
    if carrier is CarrierProfile.SCH and iso == "IN":
        return True
    return _is_apac_origin_country(row.get("Origin Country"))


def _row_eligible_for_alternative_gateway(
    row: dict[str, object],
    *,
    carrier: CarrierProfile,
) -> bool:
    if not _ag_origin_country_ok(row, carrier):
        return False
    country = to_country_iso(row.get("Origin Country"))
    if country in MANUAL_ORIGIN_AIRPORT_BY_COUNTRY:
        return True
    return bool(cell_text(row.get("Origin Airport")) or cell_text(row.get("Uplift Airport")))


def _alternative_gateway_tab_name(tab: str) -> str:
    text = cell_text(tab)
    if ALTERNATIVE_GATEWAY_SUFFIX in text:
        return text
    return f"{text}{ALTERNATIVE_GATEWAY_SUFFIX}"


def _is_base_freight_tab(tab: object) -> bool:
    return cell_text(tab).startswith("Base Freight Rates")


def append_alternative_gateway_rows(
    matrix_rows: list[dict[str, object]],
    *,
    source_stem: str = "",
    carrier: CarrierProfile | None = None,
) -> int:
    """Duplicate eligible APAC lanes with Alternative Gateway tab suffix and invoice TR."""
    carrier = carrier or detect_carrier(source_stem)
    if not carrier_allows_alternative_gateway(carrier):
        return 0

    lanes_with_fuel_tab = _lanes_with_fuel_base_tab(matrix_rows)

    additions: list[dict[str, object]] = []
    for row in matrix_rows:
        if ALTERNATIVE_GATEWAY_SUFFIX in cell_text(row.get("Tab")):
            continue
        if cell_text(row.get("Invoice type")).upper() == ALTERNATIVE_GATEWAY_INVOICE_TYPE:
            continue
        if not _should_duplicate_for_alternative_gateway(row, carrier=carrier):
            continue
        if not _row_eligible_for_alternative_gateway(row, carrier=carrier):
            continue
        if _skip_plain_base_when_fuel_tab_exists(row, lanes_with_fuel_tab):
            continue
        additions.append(_make_alternative_gateway_row(row))

    matrix_rows.extend(additions)
    return len(additions)


def _ag_source_lane_key(row: dict[str, object]) -> tuple[str, ...]:
    return (
        cell_text(row.get("Origin Country")),
        cell_text(row.get("Origin City Label")),
        cell_text(row.get("Destination Country")),
        cell_text(row.get("Destination Label")),
        cell_text(row.get("Service")),
        cell_text(row.get("Lane Type")),
        cell_text(row.get("Rate Type")),
    )


def _lanes_with_fuel_base_tab(matrix_rows: list[dict[str, object]]) -> set[tuple[str, ...]]:
    keys: set[tuple[str, ...]] = set()
    for row in matrix_rows:
        tab = cell_text(row.get("Tab"))
        if not _is_base_freight_tab(tab) or ALTERNATIVE_GATEWAY_SUFFIX in tab:
            continue
        if "+ Fuel Rates" not in tab:
            continue
        keys.add(_ag_source_lane_key(row))
    return keys


def _skip_plain_base_when_fuel_tab_exists(
    row: dict[str, object],
    lanes_with_fuel_tab: set[tuple[str, ...]],
) -> bool:
    tab = cell_text(row.get("Tab"))
    if not _is_base_freight_tab(tab) or ALTERNATIVE_GATEWAY_SUFFIX in tab:
        return False
    if "+ Fuel Rates" in tab:
        return False
    return _ag_source_lane_key(row) in lanes_with_fuel_tab


def _should_duplicate_for_alternative_gateway(
    row: dict[str, object],
    *,
    carrier: CarrierProfile,
) -> bool:
    if _is_base_freight_tab(row.get("Tab")):
        return _ag_origin_country_ok(row, carrier)
    if not carrier_includes_od_alternative_gateway(carrier):
        return False
    if not cell_text(row.get("Measurement Type")):
        return False
    iso = to_country_iso(row.get("Origin Country"))
    if carrier is CarrierProfile.FTN:
        return iso in APAC_ORIGIN_COUNTRIES or iso == "IN"
    return False


def _make_alternative_gateway_row(row: dict[str, object]) -> dict[str, object]:
    ag_row = copy.deepcopy(row)
    ag_row["Tab"] = _alternative_gateway_tab_name(cell_text(row.get("Tab")))
    ag_row["Invoice type"] = ALTERNATIVE_GATEWAY_INVOICE_TYPE
    airport = _alternative_gateway_airport(row)
    ag_row["Origin City"] = ""
    ag_row["Origin City Label"] = ""
    ag_row["Origin Airport"] = airport
    ag_row["Uplift Airport"] = airport
    return ag_row


def list_extracted_files() -> list[Path]:
    return sorted(PROCESSING_DIR.glob(EXTRACTED_GLOB), key=lambda path: path.stat().st_mtime, reverse=True)


def select_extracted_file(files: list[Path], *, auto: bool = False) -> Path:
    if not files:
        print(f"No extracted files found in: {PROCESSING_DIR}")
        sys.exit(1)

    labels = [path.name for path in files]
    if auto:
        print(f"\nAuto mode: using extracted file {labels[0]}")
        return files[0]

    indices = prompt_selection("Select extracted file to build matrix from:", labels)
    if len(indices) != 1:
        print("Please select exactly one file.")
        return select_extracted_file(files)
    return files[indices[0]]


def _suffix_from_column(column_name: str) -> str:
    match = re.search(r"_(\d+)$", column_name)
    return "" if match is None else f"_{match.group(1)}"


def discover_transport_periods(df: pd.DataFrame) -> list[PeriodColumn]:
    periods: list[PeriodColumn] = []
    for column in df.columns:
        if not VALUE_COLUMN_PATTERN.match(cell_text(column)):
            continue
        suffix = _suffix_from_column(cell_text(column))
        period_col = "Period" if not suffix else f"Period{suffix}"
        from_col = "Valid From" if not suffix else f"Valid From{suffix}"
        to_col = "Valid To" if not suffix else f"Valid To{suffix}"
        if period_col not in df.columns:
            continue
        sample = df[df[period_col].notna() & (df[period_col].astype(str).str.strip() != "")]
        if sample.empty:
            continue
        sample_row = sample.iloc[0]
        periods.append(
            PeriodColumn(
                period_label=cell_text(sample_row[period_col]),
                valid_from=sample_row.get(from_col),
                valid_to=sample_row.get(to_col),
                value_column=column,
                suffix=suffix,
            )
        )
    return periods


def discover_fuel_periods(df: pd.DataFrame) -> list[PeriodColumn]:
    periods: list[PeriodColumn] = []
    columns = list(df.columns)

    for idx, column in enumerate(columns):
        text = cell_text(column)
        lower = text.lower()
        if "value (usd)" not in lower or "part" not in lower:
            continue

        period_col = from_col = to_col = None
        for back_idx in range(idx - 1, max(idx - 4, -1), -1):
            header = cell_text(columns[back_idx])
            if header.startswith("Period") and period_col is None:
                period_col = columns[back_idx]
            elif header.startswith("Valid From") and from_col is None:
                from_col = columns[back_idx]
            elif header.startswith("Valid To") and to_col is None:
                to_col = columns[back_idx]

        if period_col is None or from_col is None or to_col is None:
            continue

        sample = df[df[period_col].notna() & (df[period_col].astype(str).str.strip() != "")]
        if sample.empty:
            continue
        sample_row = sample.iloc[0]
        periods.append(
            PeriodColumn(
                period_label=cell_text(sample_row[period_col]),
                valid_from=sample_row.get(from_col),
                valid_to=sample_row.get(to_col),
                value_column=column,
                suffix=_suffix_from_column(cell_text(column)),
            )
        )
    return periods


def build_transport_cost_specs(periods: list[PeriodColumn]) -> list[CostSpec]:
    specs: list[CostSpec] = []
    for idx, period in enumerate(periods, start=1):
        specs.append(
            CostSpec(
                group_title=f"Transport cost ({period.period_label})",
                validity_text=validity_text(period.valid_from, period.valid_to),
                currency_column=f"transport_currency_{idx}",
                rate_column=f"transport_rate_{idx}",
                rate_by=TRANSPORT_RATE_BY,
                apply_if=DEFAULT_APPLY_IF,
                fill=TRANSPORT_GROUP_FILL,
            )
        )
    return specs


def build_fuel_cost_specs(periods: list[PeriodColumn]) -> list[CostSpec]:
    specs: list[CostSpec] = []
    for idx, period in enumerate(periods, start=1):
        group_title = FSC_PERIOD_TO_GROUP.get(period.period_label.lower())
        if not group_title:
            group_title = f"Fuel Surcharge ({period.period_label})"
        specs.append(
            CostSpec(
                group_title=group_title,
                validity_text=validity_text(period.valid_from, period.valid_to),
                currency_column=f"fuel_currency_{idx}",
                rate_column=f"fuel_rate_{idx}",
                rate_by=FUEL_RATE_BY,
                apply_if=DEFAULT_APPLY_IF,
                fill=FUEL_GROUP_FILL,
            )
        )
    return specs


def build_od_cost_specs(measurements: list[str]) -> list[CostSpec]:
    specs: list[CostSpec] = []
    for idx, measurement in enumerate(measurements, start=1):
        specs.append(
            CostSpec(
                group_title=OD_COST_GROUPS[measurement],
                validity_text=DEFAULT_APPLY_IF,
                currency_column=f"od_currency_{measurement}",
                rate_column=f"od_rate_{measurement}",
                rate_by=OD_RATE_BY,
                apply_if=DEFAULT_APPLY_IF,
                fill=OD_GROUP_FILL,
            )
        )
    return specs


def _lane_key(row: pd.Series) -> tuple[str, ...]:
    return (
        cell_text(row.get(ITEM_ID_COLUMN)),
        cell_text(row.get("Origin Country")),
        cell_text(row.get("Origin City")),
        cell_text(row.get("Destination Country")),
        cell_text(row.get("Destination City")),
        cell_text(row.get("Service")),
        cell_text(row.get("Lane Type")),
    )


def _normalize_city(value: object) -> str:
    return re.sub(r"\s+", " ", cell_text(value)).casefold()


def _fsc_lane_key(row: pd.Series) -> tuple[str, ...]:
    return (
        to_country_iso(row.get("Origin Country")),
        _normalize_city(row.get("Origin City")),
        to_country_iso(row.get("Destination Country")),
        _normalize_city(row.get("Destination City")),
    )


def _shipment_values(
    row: pd.Series,
    *,
    tab_name: str,
    measurement: str = "",
    fuel_row: pd.Series | None = None,
    transport_periods: list[PeriodColumn] | None = None,
    fuel_periods: list[PeriodColumn] | None = None,
    has_fuel_on_lane: bool = False,
    rc_valid_to_override: object | None = None,
) -> dict[str, object]:
    uplift = cell_text(row.get("Uplift Airport"))
    origin_city = cell_text(row.get("Origin City"))
    destination_city = cell_text(row.get("Destination City"))
    if transport_periods:
        valid_to = compute_lane_valid_to(
            row,
            fuel_row=fuel_row,
            transport_periods=transport_periods,
            fuel_periods=fuel_periods or [],
            has_fuel_on_lane=has_fuel_on_lane,
            rc_valid_to_override=rc_valid_to_override,
        )
    else:
        valid_to = format_display_date(row.get("Valid To_5") or row.get("Valid To"))

    return {
        "Tab": tab_name,
        "Origin Country": to_country_iso(row.get("Origin Country")),
        "Origin State": cell_text(row.get("Origin State")),
        "Origin City": uplift or origin_city,
        "Origin City Label": origin_city,
        "Uplift Airport": uplift,
        "Origin Airport": uplift,
        "Rate Type": cell_text(row.get("Rate Type")),
        "Destination Country": to_country_iso(row.get("Destination Country")),
        "Destination State": cell_text(row.get("Destination State")),
        "Destination": destination_city,
        "Destination Label": destination_city,
        "Destination Airport": "",
        "Country of Clearance": cell_text(row.get("Country of Clearance")),
        "Service": cell_text(row.get("Service")),
        "Lane Type": cell_text(row.get("Lane Type")),
        "Lane Type Alt": "",
        "Service Product": service_product(
            row.get("Service"),
            row.get("Rate Type"),
            row.get("Lane Type"),
        ),
        "Measurement Type": measurement,
        "Business Segment": "",
        "Invoice type": DEFAULT_INVOICE_TYPE,
        "Valid to": valid_to,
    }


def build_base_freight_rows(
    df: pd.DataFrame,
    *,
    transport_specs: list[CostSpec],
    transport_periods: list[PeriodColumn],
    fuel_df: pd.DataFrame | None,
    fuel_specs: list[CostSpec],
    fuel_periods: list[PeriodColumn],
    rc_valid_to_override: object | None = None,
) -> list[dict[str, object]]:
    fuel_lookup: dict[tuple[str, ...], pd.Series] = {}
    if fuel_df is not None and not fuel_df.empty:
        for _, fuel_row in fuel_df.iterrows():
            fuel_lookup[_fsc_lane_key(fuel_row)] = fuel_row

    rows: list[dict[str, object]] = []
    for _, source_row in df.iterrows():
        fuel_row = fuel_lookup.get(_fsc_lane_key(source_row))
        has_fuel = fuel_row is not None and fuel_periods
        tab_name = "Base Freight Rates + Fuel Rates" if has_fuel else "Base Freight Rates"

        record = _shipment_values(
            source_row,
            tab_name=tab_name,
            fuel_row=fuel_row,
            transport_periods=transport_periods,
            fuel_periods=fuel_periods,
            has_fuel_on_lane=bool(has_fuel),
            rc_valid_to_override=rc_valid_to_override,
        )
        for spec in transport_specs:
            record[spec.currency_column] = ""
            record[spec.rate_column] = None

        for spec, period in zip(transport_specs, transport_periods, strict=True):
            amount = rate_value(source_row.get(period.value_column))
            if amount is not None:
                record[spec.currency_column] = DEFAULT_CURRENCY
                record[spec.rate_column] = amount

        for spec in fuel_specs:
            record[spec.currency_column] = ""
            record[spec.rate_column] = None

        if fuel_row is not None:
            for spec, period in zip(fuel_specs, fuel_periods, strict=True):
                amount = rate_value(fuel_row.get(period.value_column))
                if amount is not None:
                    record[spec.currency_column] = DEFAULT_CURRENCY
                    record[spec.rate_column] = amount

        for spec in build_od_cost_specs(list(OD_COST_GROUPS)):
            record[spec.currency_column] = ""
            record[spec.rate_column] = None

        rows.append(record)
    return rows


def find_accessorial_tab_name(processing_path: Path) -> str:
    """Name for the single accessorial matrix row (matches Accessorial Charges tab when present)."""
    workbook = pd.ExcelFile(processing_path)
    for name in workbook.sheet_names:
        if "accessorial" in name.strip().lower():
            return name.strip()
    return DEFAULT_ACCESSORIAL_TAB


def build_accessorial_tariff_row(
    *,
    tab_name: str,
    cost_specs: list[CostSpec],
) -> dict[str, object]:
    """One placeholder lane triggered only by Service = SPECIAL (ACC costs TBD)."""
    row: dict[str, object] = {column: "" for column in MATRIX_COLUMNS}
    row["Tab"] = tab_name
    row["Service"] = ACCESSORIAL_SERVICE_SPECIAL
    row["Service Product"] = ACCESSORIAL_SERVICE_SPECIAL
    for spec in cost_specs:
        row[spec.currency_column] = ""
        row[spec.rate_column] = None
    return row


def build_od_rows(
    df: pd.DataFrame,
    *,
    tab_name: str,
    measurement: str,
    od_spec: CostSpec,
    base_lookup: dict[tuple[str, ...], pd.Series],
    fuel_lookup: dict[tuple[str, ...], pd.Series],
    transport_periods: list[PeriodColumn],
    fuel_periods: list[PeriodColumn],
    rc_valid_to_override: object | None = None,
) -> list[dict[str, object]]:
    value_column = OD_VALUE_COLUMN if OD_VALUE_COLUMN in df.columns else None
    if value_column is None:
        for column in df.columns:
            if "value (usd)" in cell_text(column).lower():
                value_column = column
                break
    if value_column is None:
        return []

    best_by_lane: dict[tuple[str, ...], tuple[float | None, dict[str, object]]] = {}

    for _, source_row in df.iterrows():
        amount = rate_value(source_row.get(value_column))
        if amount is None:
            continue

        lane_key = _lane_key(source_row)
        base_row = base_lookup.get(lane_key)
        if base_row is None:
            base_row = source_row

        fuel_row = fuel_lookup.get(_fsc_lane_key(base_row))
        has_fuel = fuel_row is not None and bool(fuel_periods)
        record = _shipment_values(
            base_row,
            tab_name=tab_name,
            measurement=measurement,
            fuel_row=fuel_row,
            transport_periods=transport_periods,
            fuel_periods=fuel_periods,
            has_fuel_on_lane=has_fuel,
            rc_valid_to_override=rc_valid_to_override,
        )
        record["Service"] = ACCESSORIAL_SERVICE_SPECIAL
        record["Service Product"] = ACCESSORIAL_SERVICE_SPECIAL
        for spec in build_od_cost_specs(list(OD_COST_GROUPS)):
            record[spec.currency_column] = ""
            record[spec.rate_column] = None
        record[od_spec.currency_column] = DEFAULT_CURRENCY
        record[od_spec.rate_column] = amount

        existing = best_by_lane.get(lane_key)
        if existing is None:
            best_by_lane[lane_key] = (amount, record)
            continue
        prev_amount, _ = existing
        if prev_amount is None or (amount is not None and amount > prev_amount):
            best_by_lane[lane_key] = (amount, record)

    return [record for _, record in best_by_lane.values()]


def build_matrix_dataframe(processing_path: Path) -> tuple[pd.DataFrame, list[CostSpec], str]:
    workbook = pd.ExcelFile(processing_path)
    sheet_map = {name.strip().lower(): name for name in workbook.sheet_names}

    if BASE_FREIGHT_TAB not in sheet_map:
        raise RuntimeError(
            f"No 'Base Freight Rates' sheet found in {processing_path.name}. "
            "Run conversion with rate tabs first."
        )

    base_df = pd.read_excel(processing_path, sheet_name=sheet_map[BASE_FREIGHT_TAB])
    rc_valid_to_override = rc_valid_to_from_dataframe(base_df)
    transport_periods = discover_transport_periods(base_df)
    if not transport_periods:
        raise RuntimeError("No transport value periods found in Base Freight Rates.")

    transport_specs = build_transport_cost_specs(transport_periods)
    fuel_specs: list[CostSpec] = []
    fuel_periods: list[PeriodColumn] = []
    fuel_df = None
    if FUEL_RATES_TAB in sheet_map:
        fuel_df = pd.read_excel(processing_path, sheet_name=sheet_map[FUEL_RATES_TAB])
        fuel_periods = [
            period
            for period in discover_fuel_periods(fuel_df)
            if period.period_label.lower() in FSC_PERIOD_TO_GROUP
        ]
        fuel_specs = build_fuel_cost_specs(fuel_periods)

    base_lookup = {_lane_key(row): row for _, row in base_df.iterrows()}
    fuel_lookup: dict[tuple[str, ...], pd.Series] = {}
    if fuel_df is not None and not fuel_df.empty:
        for _, fuel_row in fuel_df.iterrows():
            fuel_lookup[_fsc_lane_key(fuel_row)] = fuel_row

    matrix_rows = build_base_freight_rows(
        base_df,
        transport_specs=transport_specs,
        transport_periods=transport_periods,
        fuel_df=fuel_df,
        fuel_specs=fuel_specs,
        fuel_periods=fuel_periods,
        rc_valid_to_override=rc_valid_to_override,
    )

    active_measurements: list[str] = []
    od_specs: list[CostSpec] = []
    seen_od_tabs: set[str] = set()
    for lower_name, sheet_name in sheet_map.items():
        measurement = OD_TAB_MEASUREMENT.get(lower_name.strip())
        if measurement is None:
            continue
        if lower_name.strip() in {"od total rates", "od rates total 2026"} and any(
            key in sheet_map for key in OD_TAB_MEASUREMENT if key not in {"od total rates", "od rates total 2026"}
        ):
            continue
        if measurement in seen_od_tabs:
            continue
        seen_od_tabs.add(measurement)
        od_df = pd.read_excel(processing_path, sheet_name=sheet_name)
        od_spec = CostSpec(
            group_title=OD_COST_GROUPS[measurement],
            validity_text=DEFAULT_APPLY_IF,
            currency_column=f"od_currency_{measurement}",
            rate_column=f"od_rate_{measurement}",
            rate_by=OD_RATE_BY,
            apply_if=DEFAULT_APPLY_IF,
            fill=OD_GROUP_FILL,
        )
        if measurement not in active_measurements:
            active_measurements.append(measurement)
            od_specs.append(od_spec)
        matrix_rows.extend(
            build_od_rows(
                od_df,
                tab_name=sheet_name,
                measurement=measurement,
                od_spec=od_spec,
                base_lookup=base_lookup,
                fuel_lookup=fuel_lookup,
                transport_periods=transport_periods,
                fuel_periods=fuel_periods,
                rc_valid_to_override=rc_valid_to_override,
            )
        )

    source_stem = processing_path.stem.replace("_extracted", "")
    carrier = detect_carrier(source_stem)

    removed_lanes = remove_additional_lane_rows(matrix_rows)
    if removed_lanes:
        print(f"  - Additional lane rows removed: {removed_lanes}")

    apply_origin_and_conditional_fields(matrix_rows)

    airport_updates = apply_destination_airports(matrix_rows, carrier)
    if airport_updates:
        print(f"  - Destination airport codes set: {airport_updates}")

    for row in matrix_rows:
        apply_conditional_destination_from_trigger(row)

    apply_business_segments(matrix_rows, source_stem=source_stem)

    cn_tw = apply_apac_cn_tw_conditionals(matrix_rows, carrier=carrier, source_stem=source_stem)
    if cn_tw:
        print(f"  - CN/TW ATA/ATD conditional lanes touched: {cn_tw}")
    ag_count = append_alternative_gateway_rows(matrix_rows, source_stem=source_stem, carrier=carrier)
    if ag_count:
        print(f"  - Alternative Gateway rows added: {ag_count}")

    zone_copies = append_schenker_us_zone_destination_rows(matrix_rows, carrier)
    if zone_copies:
        print(f"  - Schenker US zone destination copies: {zone_copies}")

    vn_carrier_rows = append_sch_vn_carrier_rows(matrix_rows, carrier)
    if vn_carrier_rows:
        print(f"  - SCH Vietnam carrier rows added: {vn_carrier_rows}")

    if rc_valid_to_override is not None:
        rc_display = format_display_date(rc_valid_to_override)
        if rc_display:
            for row in matrix_rows:
                row["Valid to"] = rc_display

    cost_specs = [*transport_specs, *fuel_specs, *od_specs]

    for lane_number, row in enumerate(matrix_rows, start=1):
        row["Lane #"] = lane_number

    columns = [
        *MATRIX_COLUMNS,
        *[column for spec in cost_specs for column in (spec.currency_column, spec.rate_column)],
    ]
    return pd.DataFrame(matrix_rows, columns=columns), cost_specs, source_stem


def _style_header_cell(
    cell,
    *,
    fill: PatternFill = HEADER_FILL,
    center: bool = False,
    bold: bool = True,
) -> None:
    cell.font = BOLD if bold else NORMAL
    cell.fill = fill
    cell.alignment = CENTER if center else LEFT
    cell.border = THIN_BORDER


def _style_meta_cell(cell, *, fill: PatternFill = COST_META_FILL) -> None:
    cell.font = NORMAL
    cell.fill = fill
    cell.alignment = LEFT
    cell.border = THIN_BORDER


def _merge_write(worksheet, row: int, start_col: int, end_col: int, value: str, *, fill: PatternFill) -> None:
    if end_col > start_col:
        worksheet.merge_cells(
            start_row=row,
            start_column=start_col,
            end_row=row,
            end_column=end_col,
        )
    cell = worksheet.cell(row, start_col, value)
    if row == COST_GROUP_ROW:
        _style_header_cell(cell, fill=fill, center=True)
    else:
        _style_meta_cell(cell, fill=fill)


def write_rate_card_sheet(
    workbook: Workbook,
    matrix_df: pd.DataFrame,
    cost_specs: list[CostSpec],
    *,
    agreement_title: str,
    sheet_name: str = "Rate card",
) -> None:
    worksheet = workbook.active
    worksheet.title = sheet_name

    worksheet.cell(TITLE_ROW, 1, "Rate Card (Active)")
    worksheet.cell(AGREEMENT_ROW, 1, agreement_title)

    shipment_count = len(MATRIX_COLUMNS)
    cost_layout: list[tuple[CostSpec, int, int]] = []
    next_col = shipment_count + 1
    for spec in cost_specs:
        cost_layout.append((spec, next_col, next_col + 1))
        next_col += 2

    for spec, currency_col, rate_col in cost_layout:
        _merge_write(worksheet, COST_GROUP_ROW, currency_col, rate_col, spec.group_title, fill=spec.fill)
        _merge_write(worksheet, VALIDITY_ROW, currency_col, rate_col, spec.validity_text, fill=COST_META_FILL)
        _merge_write(worksheet, RATE_BY_ROW, currency_col, rate_col, spec.rate_by, fill=COST_META_FILL)
        _merge_write(worksheet, CONDITIONAL_ROW, currency_col, rate_col, spec.apply_if, fill=COST_META_FILL)

    for col_index, header in enumerate(MATRIX_COLUMNS, start=1):
        export_header = EXPORT_HEADERS.get(header, header)
        cell = worksheet.cell(COLUMN_HEADER_ROW, col_index, export_header)
        header_bold = header == "Lane #" or header in BOLD_CONDITION_COLUMNS
        _style_header_cell(cell, bold=header_bold)

    for spec, currency_col, rate_col in cost_layout:
        currency_header = worksheet.cell(COLUMN_HEADER_ROW, currency_col, "Currency")
        _style_header_cell(currency_header, center=True, fill=TRANSPORT_COST_FILL)
        rate_header = worksheet.cell(COLUMN_HEADER_ROW, rate_col, "p/unit")
        _style_header_cell(rate_header, center=True, fill=TRANSPORT_COST_FILL)

    for row_offset, (_, row) in enumerate(matrix_df.iterrows()):
        excel_row = DATA_START_ROW + row_offset
        for col_index, header in enumerate(MATRIX_COLUMNS, start=1):
            cell = worksheet.cell(excel_row, col_index, row.get(header))
            cell.alignment = LEFT
            cell.border = THIN_BORDER
            if header in BOLD_CONDITION_COLUMNS:
                cell.font = BOLD

        for spec, currency_col, rate_col in cost_layout:
            currency_cell = worksheet.cell(excel_row, currency_col, row.get(spec.currency_column))
            currency_cell.alignment = CENTER
            currency_cell.border = THIN_BORDER

            value = row.get(spec.rate_column)
            rate_cell = worksheet.cell(excel_row, rate_col)
            rate_cell.border = THIN_BORDER
            if value is not None and value != "":
                rate_cell.value = value
                rate_cell.number_format = RATE_NUMBER_FORMAT
                rate_cell.alignment = CENTER

    worksheet.freeze_panes = worksheet.cell(DATA_START_ROW, 1)
    worksheet.sheet_view.showGridLines = False


def save_matrix(
    matrix_df: pd.DataFrame,
    cost_specs: list[CostSpec],
    *,
    agreement_title: str,
    output_path: Path | None = None,
    source_stem: str,
) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if output_path is None:
        output_path = OUTPUT_DIR / f"{source_stem}_matrix.xlsx"

    workbook = Workbook()
    write_rate_card_sheet(
        workbook,
        matrix_df,
        cost_specs,
        agreement_title=agreement_title,
    )
    workbook.save(output_path)

    from matrix_diff_highlight import apply_matrix_highlights_if_saved

    apply_matrix_highlights_if_saved(output_path, matrix_df)
    return output_path


def run_build_matrix(
    *,
    source_file: Path | None = None,
    output_path: Path | None = None,
    auto: bool = False,
) -> Path:
    ensure_workspace_dirs()
    files = list_extracted_files()
    file_path = source_file or select_extracted_file(files, auto=auto)

    matrix_df, cost_specs, agreement_title = build_matrix_dataframe(file_path)
    print(f"\nBuilding matrix from {file_path.name}:")
    print(f"  - Base freight lanes: {len(matrix_df[matrix_df['Tab'].astype(str).str.startswith('Base Freight')])}")
    od_rows = matrix_df[matrix_df["Measurement Type"].astype(str).str.startswith("OD")]
    print(f"  - OD handling rows: {len(od_rows)}")
    print(f"  - Cost columns: {len(cost_specs)}")

    saved_path = save_matrix(
        matrix_df,
        cost_specs,
        agreement_title=agreement_title,
        output_path=output_path,
        source_stem=file_path.stem.replace("_extracted", ""),
    )
    print(f"\nSaved matrix ({len(matrix_df)} rows) to: {saved_path}")
    return saved_path, matrix_df


def main() -> int:
    try:
        run_build_matrix()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
