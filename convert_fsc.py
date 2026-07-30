"""
Convert GAR Fuel Surcharge (FSC) workbooks from input/fsc/ into cleaned dataframes.

FSC layout (Fuel Rates sheet):
  - Lane columns: Row Number, Region, Origin/Destination Country/City
  - Monthly and bi-weekly fuel value columns (*Value (USD)*)
  - Side panel: Calc. Rules and FSC Valid Dates (month -> from/to)
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from convert_to_processing import (
    CALC_RULE_FIELD_MAP,
    PERIOD_COLUMN,
    VALID_FROM_COLUMN,
    VALID_TO_COLUMN,
    CalcRules,
    _attach_calc_rules,
    _cell_text,
    _format_metadata_date,
    _normalize_headers,
    prompt_selection,
)
from project_paths import FSC_INPUT_DIR, ensure_workspace_dirs

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}
FSC_SHEET_NAME = "Fuel Rates"
FSC_VALUE_PATTERN = re.compile(r"value\s*\(usd\)", re.IGNORECASE)
FSC_LANE_COLUMNS = (
    "Row Number",
    "Region",
    "Origin Country",
    "Origin City",
    "Destination Country",
    "Destination City",
)

MONTH_ALIASES = {
    "january": "jan",
    "february": "feb",
    "march": "mar",
    "april": "apr",
    "may": "may",
    "june": "jun",
    "july": "jul",
    "august": "aug",
    "september": "sep",
    "october": "oct",
    "november": "nov",
    "december": "dec",
}


@dataclass
class FscMetadata:
    calc_rules: CalcRules = field(default_factory=CalcRules)
    valid_dates: dict[str, tuple[object, object]] = field(default_factory=dict)


def list_fsc_files() -> list[Path]:
    return [
        path
        for path in sorted(FSC_INPUT_DIR.rglob("*"))
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    ]


def _file_label(path: Path) -> str:
    try:
        return str(path.relative_to(FSC_INPUT_DIR))
    except ValueError:
        return path.name


def select_fsc_file(files: list[Path], *, auto: bool = False) -> Path:
    if not files:
        print(f"No Excel files found in: {FSC_INPUT_DIR}")
        sys.exit(1)

    labels = [_file_label(path) for path in files]
    if auto:
        print(f"\nAuto mode: using FSC file {labels[0]}")
        return files[0]

    print("\nAvailable FSC files:")
    for i, label in enumerate(labels, start=1):
        print(f"  {i}. {label}")

    indices = prompt_selection("Select FSC file to convert:", labels)
    if len(indices) != 1:
        print("Please select exactly one file.")
        return select_fsc_file(files)
    return files[indices[0]]


def _row_texts(row: pd.Series) -> list[str]:
    return [_cell_text(value).lower() for value in row]


def _is_fsc_value_header(header: object) -> bool:
    return bool(FSC_VALUE_PATTERN.search(_cell_text(header)))


def _canonical_period_key(text: object) -> str:
    key = re.sub(r"\s+", " ", _cell_text(text).replace("\n", " ")).strip().lower()
    key = re.sub(r"\s+gar\s+\d+", "", key).strip()
    key = re.sub(r"\s*value\s*\(usd\)\s*$", "", key, flags=re.IGNORECASE).strip()
    parts = key.split()
    if parts and parts[0] in MONTH_ALIASES:
        parts[0] = MONTH_ALIASES[parts[0]]
    return " ".join(parts)


def find_fsc_header_row(df_raw: pd.DataFrame, max_scan_rows: int = 10) -> int | None:
    scan_limit = min(len(df_raw), max_scan_rows)
    for row_idx in range(scan_limit):
        texts = set(_row_texts(df_raw.iloc[row_idx]))
        if "row number" in texts and "origin country" in texts:
            return row_idx
    return None


def _find_side_panel_start(df_raw: pd.DataFrame, header_row_idx: int) -> int | None:
    scan_limit = min(len(df_raw), header_row_idx + 10)
    for row_idx in range(scan_limit):
        for col_idx in range(len(df_raw.columns)):
            text = _cell_text(df_raw.iloc[row_idx, col_idx]).lower()
            if text in {"calc. rules", "fsc valid dates"}:
                return col_idx

    header_row = df_raw.iloc[header_row_idx]
    for col_idx, value in enumerate(header_row):
        if _cell_text(value).lower() == "difference mom":
            return col_idx + 1
    return None


def _extract_fsc_calc_rules(df_raw: pd.DataFrame, side_col: int) -> CalcRules:
    calc_header_row_idx: int | None = None
    for row_idx in range(min(10, len(df_raw))):
        if "calculation base" in _cell_text(df_raw.iloc[row_idx, side_col]).lower():
            calc_header_row_idx = row_idx
            break

    if calc_header_row_idx is None or calc_header_row_idx + 1 >= len(df_raw):
        return CalcRules()

    values: dict[str, str] = {}
    header_row = df_raw.iloc[calc_header_row_idx]
    data_row = df_raw.iloc[calc_header_row_idx + 1]
    for offset in range(6):
        col_idx = side_col + offset
        if col_idx >= len(header_row):
            break
        field_key = CALC_RULE_FIELD_MAP.get(_cell_text(header_row.iloc[col_idx]).lower())
        if field_key is None:
            continue
        values[field_key] = _cell_text(data_row.iloc[col_idx])

    return CalcRules(**values)


def _extract_fsc_valid_dates(df_raw: pd.DataFrame, side_col: int) -> dict[str, tuple[object, object]]:
    dates_start: int | None = None
    for row_idx in range(min(40, len(df_raw))):
        if _cell_text(df_raw.iloc[row_idx, side_col]).lower() == "fsc valid dates":
            dates_start = row_idx + 2
            break

    if dates_start is None:
        return {}

    mapping: dict[str, tuple[object, object]] = {}
    for row_idx in range(dates_start, min(len(df_raw), dates_start + 40)):
        month = df_raw.iloc[row_idx, side_col]
        valid_from = df_raw.iloc[row_idx, side_col + 1]
        valid_to = df_raw.iloc[row_idx, side_col + 2]
        if pd.isna(month) or not _cell_text(month):
            break
        if _cell_text(month).lower().startswith("from"):
            continue
        key = _canonical_period_key(month)
        if key:
            mapping[key] = (valid_from, valid_to)
    return mapping


def _lookup_validity(
    header_name: str,
    valid_dates: dict[str, tuple[object, object]],
) -> tuple[object, object]:
    key = _canonical_period_key(header_name)
    if key in valid_dates:
        return valid_dates[key]

    for period_key, validity in valid_dates.items():
        if key.endswith(period_key) or period_key.endswith(key):
            return validity

    return (pd.NA, pd.NA)


def _is_data_row(row: pd.Series, lane_headers: list[str], headers: list[str]) -> bool:
    if "Row Number" not in lane_headers:
        return any(_cell_text(value) for value in row)

    col_idx = headers.index("Row Number")
    if col_idx >= len(row):
        return False
    row_number = _cell_text(row.iloc[col_idx])
    if not row_number:
        return False
    try:
        int(float(row_number))
    except ValueError:
        return False
    return True


def clean_fsc_df(df_raw: pd.DataFrame) -> tuple[pd.DataFrame, FscMetadata]:
    header_row_idx = find_fsc_header_row(df_raw)
    if header_row_idx is None:
        raise RuntimeError("Could not find FSC header row (Row Number / Origin Country).")

    side_col = _find_side_panel_start(df_raw, header_row_idx) or 30
    metadata = FscMetadata(
        calc_rules=_extract_fsc_calc_rules(df_raw, side_col),
        valid_dates=_extract_fsc_valid_dates(df_raw, side_col),
    )

    headers = _normalize_headers(df_raw.iloc[header_row_idx].tolist())
    value_indices = [idx for idx, header in enumerate(headers) if _is_fsc_value_header(header)]
    if not value_indices:
        raise RuntimeError("No fuel value columns found in FSC sheet.")

    first_value_idx = value_indices[0]
    lane_headers = [header for header in headers[:first_value_idx] if header in FSC_LANE_COLUMNS]
    if not lane_headers:
        lane_headers = [
            header
            for header in headers[:first_value_idx]
            if not header.startswith("column_")
        ]

    value_headers = [headers[idx] for idx in value_indices]
    records: list[dict[str, object]] = []
    for row_idx in range(header_row_idx + 1, len(df_raw)):
        row = df_raw.iloc[row_idx]
        if not _is_data_row(row, lane_headers, headers):
            continue

        record: dict[str, object] = {}
        for lane_header in lane_headers:
            col_idx = headers.index(lane_header)
            record[lane_header] = row.iloc[col_idx] if col_idx < len(row) else pd.NA
        for header_name in value_headers:
            col_idx = headers.index(header_name)
            record[header_name] = row.iloc[col_idx] if col_idx < len(row) else pd.NA
        records.append(record)

    if not records:
        raise RuntimeError("No FSC lane rows found.")

    lane_df = pd.DataFrame(records)[lane_headers]
    if len(value_headers) == 1:
        header_name = value_headers[0]
        valid_from, valid_to = _lookup_validity(header_name, metadata.valid_dates)
        result = lane_df.copy()
        result[header_name] = [record[header_name] for record in records]
        result.insert(0, VALID_TO_COLUMN, _format_metadata_date(valid_to))
        result.insert(0, VALID_FROM_COLUMN, _format_metadata_date(valid_from))
        result.insert(0, PERIOD_COLUMN, _canonical_period_key(header_name))
        return _attach_calc_rules(result, metadata.calc_rules), metadata

    parts: list[pd.DataFrame] = [lane_df]
    for value_idx, header_name in enumerate(value_headers):
        suffix = "" if value_idx == 0 else f"_{value_idx + 1}"
        valid_from, valid_to = _lookup_validity(header_name, metadata.valid_dates)
        parts.append(
            pd.DataFrame(
                {
                    f"{PERIOD_COLUMN}{suffix}": _canonical_period_key(header_name),
                    f"{VALID_FROM_COLUMN}{suffix}": _format_metadata_date(valid_from),
                    f"{VALID_TO_COLUMN}{suffix}": _format_metadata_date(valid_to),
                    header_name: [record[header_name] for record in records],
                }
            )
        )

    return _attach_calc_rules(pd.concat(parts, axis=1), metadata.calc_rules), metadata


def collect_fsc_frames(file_path: Path) -> list[tuple[str, pd.DataFrame]]:
    workbook = pd.ExcelFile(file_path)
    sheet_name = FSC_SHEET_NAME if FSC_SHEET_NAME in workbook.sheet_names else workbook.sheet_names[0]
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    df, metadata = clean_fsc_df(raw)

    meta_parts: list[str] = []
    if not metadata.calc_rules.is_empty:
        meta_parts.append(
            "calc rules: "
            f"{metadata.calc_rules.calculation_base} / "
            f"{metadata.calc_rules.fap_charge_code} / "
            f"{metadata.calc_rules.uom}"
        )
    if metadata.valid_dates:
        meta_parts.append(f"validity periods: {len(metadata.valid_dates)}")
    meta_summary = f"; {'; '.join(meta_parts)}" if meta_parts else ""

    print(
        f"  Loaded FSC: {sheet_name} -> 'Fuel Rates' "
        f"({len(df)} rows, {len(df.columns)} columns; "
        f"{len([h for h in df.columns if FSC_VALUE_PATTERN.search(str(h))])} fuel periods"
        f"{meta_summary})"
    )
    return [("Fuel Rates", df)]


def run_convert_fsc(
    *,
    auto: bool = False,
    file_path: Path | None = None,
) -> tuple[Path, list[tuple[str, pd.DataFrame]]]:
    ensure_workspace_dirs()
    files = list_fsc_files()
    selected_file = file_path or select_fsc_file(files, auto=auto)

    if not selected_file.exists():
        raise FileNotFoundError(f"FSC file not found: {selected_file}")

    print(f"\nConverting FSC {_file_label(selected_file)}...")
    frames = collect_fsc_frames(selected_file)
    return selected_file, frames
