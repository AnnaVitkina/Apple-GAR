"""
Convert selected tabs from GAR input workbooks into cleaned dataframes
and save them to the processing/ folder.

Interactive flow:
  1. Choose source: rate only, FSC only, or rate + FSC
  2. Choose input file(s) from input/rate/ and/or input/fsc/
  3. For rate files: review proposed default tabs
  4. Write one multi-sheet workbook to processing/
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from project_paths import FSC_INPUT_DIR, INPUT_DIR, PROCESSING_DIR, RATE_INPUT_DIR, ensure_workspace_dirs

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}
SOURCE_MODE_RATE = "rate"
SOURCE_MODE_FSC = "fsc"
SOURCE_MODE_BOTH = "both"
SOURCE_MODE_CHOICES = (SOURCE_MODE_RATE, SOURCE_MODE_FSC, SOURCE_MODE_BOTH)

DEFAULT_SHEET_EXACT = ("Additional Info", "Base Freight Rates", "Index")

HEADER_MARKERS = (
    "original naming convention",
    "origin city",
    "destination city",
    "service type",
    "charge code supplier",
    "billing account",
    "2026 item_id",
)

SECTION_TITLE_PATTERN = re.compile(
    r"(?:gor\d+|boc\d+)\s+.+\s+common\s+rating",
    re.IGNORECASE,
)
VALUE_HEADER_PATTERN = re.compile(r"^Value\s*\(([^)]+)\)\s*$", re.IGNORECASE)
PEAK_NONPEAK_PATTERN = re.compile(r"\b(?:non-peak|peak)\b", re.IGNORECASE)
VALID_FROM_COLUMN = "Valid From"
VALID_TO_COLUMN = "Valid To"
RC_VALID_FROM_COLUMN = "RC Valid From"
RC_VALID_TO_COLUMN = "RC Valid To"
PERIOD_COLUMN = "Period"
CALC_RULE_COLUMNS = (
    "Calculation Base",
    "Calculation Base Desc.",
    "Price Unit",
    "UoM",
    "FAP Charge Code",
)
CALC_RULE_FIELD_MAP = {
    "calculation base": "calculation_base",
    "calculation base desc.": "calculation_base_desc",
    "price unit": "price_unit",
    "uom": "uom",
    "fap charge code": "fap_charge_code",
}


@dataclass
class CalcRules:
    calculation_base: str = ""
    calculation_base_desc: str = ""
    price_unit: str = ""
    uom: str = ""
    fap_charge_code: str = ""

    def as_columns(self) -> dict[str, str]:
        return {
            "Calculation Base": self.calculation_base,
            "Calculation Base Desc.": self.calculation_base_desc,
            "Price Unit": self.price_unit,
            "UoM": self.uom,
            "FAP Charge Code": self.fap_charge_code,
        }

    @property
    def is_empty(self) -> bool:
        return not any(self.as_columns().values())


@dataclass(frozen=True)
class SheetSelection:
    file_path: Path
    sheet_name: str

    @property
    def label(self) -> str:
        return f"{self.file_path.name} -> {self.sheet_name}"


@dataclass
class SheetMetadata:
    section_titles: list[str] = field(default_factory=list)
    valid_from: object = None
    valid_to: object = None
    calc_rules: CalcRules = field(default_factory=CalcRules)


def _is_valid_from_label(lower: str) -> bool:
    return lower.startswith("from (yy") or lower.startswith("from (dd/mm/yyyy)")


def _is_valid_to_label(lower: str) -> bool:
    return lower.startswith("to (yy") or lower.startswith("to (dd/mm/yyyy)")


def _file_label(path: Path, base_dir: Path = INPUT_DIR) -> str:
    try:
        return str(path.relative_to(base_dir))
    except ValueError:
        return path.name


def list_rate_files() -> list[Path]:
    return [
        path
        for path in sorted(RATE_INPUT_DIR.rglob("*"))
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    ]


def list_input_files() -> list[Path]:
    """Backward-compatible alias for rate files."""
    return list_rate_files()


def parse_selection(raw: str, max_index: int) -> list[int]:
    """Parse '1,3-5' into zero-based indices."""
    raw = raw.strip().lower()
    if raw in {"all", "*"}:
        return list(range(max_index))

    indices: set[int] = set()
    for part in re.split(r"\s*,\s*", raw):
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = int(start_s) - 1
            end = int(end_s) - 1
            if start > end or start < 0 or end >= max_index:
                raise ValueError(f"Invalid range: {part}")
            indices.update(range(start, end + 1))
        else:
            idx = int(part) - 1
            if idx < 0 or idx >= max_index:
                raise ValueError(f"Invalid index: {part}")
            indices.add(idx)
    return sorted(indices)


def prompt_selection(title: str, items: list[str], allow_empty: bool = False) -> list[int]:
    if not items:
        return []

    print(f"\n{title}")
    for i, item in enumerate(items, start=1):
        print(f"  {i}. {item}")

    hint = "Enter numbers (e.g. 1,3 or 1-3) or 'all'"

    while True:
        raw = input(f"{hint}: ").strip()
        if not raw and allow_empty:
            return []
        if not raw:
            print("Please enter at least one choice.")
            continue
        try:
            chosen = parse_selection(raw, len(items))
            if chosen or allow_empty:
                return chosen
            print("Please enter at least one choice.")
        except ValueError as exc:
            print(f"Invalid input: {exc}")


def _is_default_od_tab(sheet_name: str) -> bool:
    """Match main OD rate tabs but skip carrier/lane-specific OD sheets."""
    lower_name = sheet_name.strip().lower()
    if "od" not in lower_name or "rate" not in lower_name:
        return False
    return lower_name.startswith("od ")


def propose_default_tabs(sheet_names: list[str]) -> list[str]:
    """Match default exact tabs and main OD rate tabs."""
    by_lower = {name.lower(): name for name in sheet_names}
    selected: list[str] = []
    seen_lower: set[str] = set()

    for exact_name in DEFAULT_SHEET_EXACT:
        actual = by_lower.get(exact_name.lower())
        if actual is not None and actual.lower() not in seen_lower:
            selected.append(actual)
            seen_lower.add(actual.lower())

    for sheet_name in sheet_names:
        lower_name = sheet_name.strip().lower()
        if lower_name in seen_lower:
            continue
        if _is_default_od_tab(sheet_name):
            selected.append(sheet_name)
            seen_lower.add(lower_name)

    return selected


def select_input_file(files: list[Path], *, auto: bool = False) -> Path:
    if not files:
        print(f"No Excel files found in: {RATE_INPUT_DIR}")
        sys.exit(1)

    labels = [_file_label(path, RATE_INPUT_DIR) for path in files]

    if auto:
        print(f"\nAuto mode: using rate file {labels[0]}")
        return files[0]

    print("\nAvailable rate files:")
    for i, label in enumerate(labels, start=1):
        print(f"  {i}. {label}")

    indices = prompt_selection("Select rate file to convert:", labels)
    if len(indices) != 1:
        print("Please select exactly one file.")
        return select_input_file(files)
    return files[indices[0]]


def prompt_source_mode(*, auto: bool = False) -> str:
    options = [
        ("Rate only (input/rate/)", SOURCE_MODE_RATE),
        ("FSC only (input/fsc/)", SOURCE_MODE_FSC),
        ("Rate + FSC (both)", SOURCE_MODE_BOTH),
    ]

    if auto:
        print("\nAuto mode: using Rate only.")
        return SOURCE_MODE_RATE

    print("\nSelect conversion source:")
    for i, (label, _) in enumerate(options, start=1):
        print(f"  {i}. {label}")

    while True:
        choice = input("Choose option (1-3): ").strip()
        if choice in {"1", "2", "3"}:
            return options[int(choice) - 1][1]
        print("Invalid choice. Enter 1, 2, or 3.")


def print_all_tabs(sheet_names: list[str], selected: list[str]) -> None:
    print("\nAll available tabs:")
    for i, name in enumerate(sheet_names, start=1):
        marker = " [selected]" if name in selected else ""
        print(f"  {i}. {name}{marker}")


def print_current_selection(selected: list[str]) -> None:
    if selected:
        print("\nCurrent tab selection:")
        for name in selected:
            print(f"  - {name}")
    else:
        print("\nNo tabs selected yet.")


def select_tabs_interactive(file_path: Path, sheet_names: list[str]) -> list[str]:
    selected = propose_default_tabs(sheet_names)

    while True:
        print_all_tabs(sheet_names, selected)
        print_current_selection(selected)
        print("\nTab selection options:")
        print("  1. Accept current selection and convert")
        print("  2. Change tabs (replace selection)")
        print("  3. Add tabs to current selection")

        choice = input("Choose option (1-3): ").strip()

        if choice == "1":
            if not selected:
                print("Select at least one tab before converting.")
                continue
            return selected

        if choice == "2":
            indices = prompt_selection(
                f"Choose sheet(s) for {_file_label(file_path, RATE_INPUT_DIR)}:",
                sheet_names,
            )
            selected = [sheet_names[i] for i in indices]
            continue

        if choice == "3":
            remaining = [name for name in sheet_names if name not in selected]
            if not remaining:
                print("All tabs are already selected.")
                continue
            indices = prompt_selection(
                "Choose additional sheet(s) to add:",
                remaining,
            )
            for idx in indices:
                tab_name = remaining[idx]
                if tab_name not in selected:
                    selected.append(tab_name)
            continue

        print("Invalid choice. Enter 1, 2, or 3.")


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _format_metadata_date(value: object) -> object:
    if pd.isna(value):
        return pd.NA
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def _metadata_sortable_date(value: object):
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    text = _cell_text(value)
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text[:10], fmt)
        except ValueError:
            continue
    return text


def _row_texts(row: pd.Series) -> list[str]:
    return [_cell_text(value).lower() for value in row]


def _is_section_title(value: object) -> bool:
    text = _cell_text(value)
    return bool(text and SECTION_TITLE_PATTERN.search(text))


def _find_section_title_in_row(row: pd.Series) -> str | None:
    for value in row:
        text = _cell_text(value)
        if _is_section_title(text):
            return text
    return None


def _extract_calc_rules(df_raw: pd.DataFrame, header_row_idx: int) -> CalcRules:
    """Read Calc. Rules table from preamble rows above the lane header."""
    calc_header_row_idx: int | None = None
    scan_limit = min(header_row_idx, len(df_raw))

    for row_idx in range(scan_limit):
        if "calculation base" in set(_row_texts(df_raw.iloc[row_idx])):
            calc_header_row_idx = row_idx
            break

    if calc_header_row_idx is None or calc_header_row_idx + 1 >= len(df_raw):
        return CalcRules()

    header_row = df_raw.iloc[calc_header_row_idx]
    data_row = df_raw.iloc[calc_header_row_idx + 1]
    values: dict[str, str] = {}

    for col_idx, header in enumerate(header_row):
        field_key = CALC_RULE_FIELD_MAP.get(_cell_text(header).lower())
        if field_key is None:
            continue
        values[field_key] = _cell_text(data_row.iloc[col_idx])

    return CalcRules(**values)


def _dates_after_valid_to_label(row: pd.Series, label_col_idx: int) -> list[object]:
    dates: list[object] = []
    for candidate in row.iloc[label_col_idx + 1 :]:
        if pd.isna(candidate):
            continue
        if _is_valid_to_label(_cell_text(candidate).lower()):
            continue
        if _cell_text(candidate):
            dates.append(candidate)
    return dates


def extract_sheet_metadata(df_raw: pd.DataFrame, max_scan_rows: int = 50) -> SheetMetadata:
    metadata = SheetMetadata()
    scan_limit = min(len(df_raw), max_scan_rows)

    for row_idx in range(scan_limit):
        row = df_raw.iloc[row_idx]
        for col_idx, value in enumerate(row):
            text = _cell_text(value)
            lower = text.lower()

            if _is_section_title(text) and text not in metadata.section_titles:
                metadata.section_titles.append(text)

            if _is_valid_from_label(lower):
                if col_idx + 1 < len(row):
                    candidate = row.iloc[col_idx + 1]
                    if pd.notna(candidate):
                        metadata.valid_from = candidate
                for candidate in row.iloc[col_idx + 1 :]:
                    if pd.notna(candidate) and not _is_valid_from_label(_cell_text(candidate).lower()):
                        metadata.valid_from = candidate
                        break

            if _is_valid_to_label(lower):
                dates_after = _dates_after_valid_to_label(row, col_idx)
                if len(dates_after) == 1:
                    metadata.valid_to = dates_after[0]
                elif metadata.valid_to is None and dates_after:
                    metadata.valid_to = max(dates_after, key=_metadata_sortable_date)

    header_row_idx = find_header_row_index(df_raw, max_scan_rows=max_scan_rows)
    if header_row_idx is not None:
        metadata.calc_rules = _extract_calc_rules(df_raw, header_row_idx)

    return metadata


def _is_additional_info_header_row(row: pd.Series) -> bool:
    texts = set(_row_texts(row))
    return "origin city" in texts or "destination city" in texts


def _normalize_headers(headers: list[object]) -> list[str]:
    cleaned: list[str] = []
    seen: dict[str, int] = {}

    for idx, header in enumerate(headers, start=1):
        value = _cell_text(header)
        if not value or value.lower() == "nan":
            value = f"column_{idx}"

        base = value
        count = seen.get(base, 0)
        if count:
            value = f"{base}_{count + 1}"
        seen[base] = count + 1
        cleaned.append(value)

    return cleaned


def _drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    non_empty_mask = df.apply(
        lambda row: any(_cell_text(value) for value in row),
        axis=1,
    )
    return df.loc[non_empty_mask].copy()


def _drop_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    keep_columns = [
        column
        for column in df.columns
        if any(_cell_text(value) for value in df[column])
    ]
    return df.loc[:, keep_columns].copy()


def _row_contains_marker(row: pd.Series) -> bool:
    texts = set(_row_texts(row))
    return any(marker in texts for marker in HEADER_MARKERS)


def _find_header_row_by_marker(
    df_raw: pd.DataFrame,
    marker: str,
    max_scan_rows: int = 50,
) -> int | None:
    scan_limit = min(len(df_raw), max_scan_rows)
    for row_idx in range(scan_limit):
        if marker in set(_row_texts(df_raw.iloc[row_idx])):
            return row_idx
    return None


def _normalize_period_label(value: object) -> str:
    text = _cell_text(value)
    if not text:
        return ""
    return re.sub(r"\s+", " ", text.replace("\n", " / ")).strip()


def _extract_per_column_period_labels(
    df_raw: pd.DataFrame,
    header_row_idx: int,
) -> dict[int, str]:
    """Map raw column index to peak/non-peak period label from preamble rows."""
    mapping: dict[int, str] = {}

    for row_idx in range(header_row_idx):
        row = df_raw.iloc[row_idx]
        for col_idx, value in enumerate(row):
            text = _cell_text(value)
            if not text or not PEAK_NONPEAK_PATTERN.search(text):
                continue
            mapping[col_idx] = _normalize_period_label(text)

    return mapping


def _extract_per_column_validity(
    df_raw: pd.DataFrame,
    header_row_idx: int,
) -> dict[int, tuple[object, object]]:
    """Map raw column index to (valid_from, valid_to) from header rows above the table."""
    from_row_idx: int | None = None
    to_row_idx: int | None = None

    for row_idx in range(header_row_idx):
        row = df_raw.iloc[row_idx]
        for value in row:
            lower = _cell_text(value).lower()
            if _is_valid_from_label(lower):
                from_row_idx = row_idx
            elif _is_valid_to_label(lower):
                to_row_idx = row_idx

    if from_row_idx is None or to_row_idx is None:
        return {}

    from_row = df_raw.iloc[from_row_idx]
    to_row = df_raw.iloc[to_row_idx]
    mapping: dict[int, tuple[object, object]] = {}
    for col_idx in range(min(len(from_row), len(to_row))):
        valid_from = from_row.iloc[col_idx]
        valid_to = to_row.iloc[col_idx]
        if pd.isna(valid_from) or pd.isna(valid_to):
            continue
        if _is_valid_from_label(_cell_text(valid_from).lower()):
            continue
        mapping[col_idx] = (valid_from, valid_to)
    return mapping


def _apply_internal_value_headers(headers: list[str], internal_row: pd.Series) -> list[str]:
    """Replace display headers like 'GOR Year' with internal names like 'Value (USD)'."""
    updated = list(headers)
    value_count = 0

    for col_idx, header in enumerate(updated):
        if col_idx >= len(internal_row):
            continue
        internal_name = _cell_text(internal_row.iloc[col_idx])
        if not VALUE_HEADER_PATTERN.match(internal_name):
            continue
        value_count += 1
        updated[col_idx] = internal_name if value_count == 1 else f"{internal_name}_{value_count}"

    return updated


def _is_value_header_name(header_name: str) -> bool:
    text = _cell_text(header_name)
    return bool(VALUE_HEADER_PATTERN.match(text) or re.match(r"^Value\s*\([^)]+\)_\d+$", text, re.I))


def _is_od_value_header_name(header_name: str) -> bool:
    text = _cell_text(header_name).lower()
    return "value (usd)" in text or "o/d total value" in text


def _column_suffix(value_idx: int) -> str:
    return "" if value_idx == 0 else f"_{value_idx + 1}"


def _attach_calc_rules(df: pd.DataFrame, calc_rules: CalcRules) -> pd.DataFrame:
    if calc_rules.is_empty:
        return df

    enriched = df.copy()
    for column_name in reversed(CALC_RULE_COLUMNS):
        enriched.insert(0, column_name, calc_rules.as_columns()[column_name])
    return enriched


def _attach_per_value_column_validity(
    df: pd.DataFrame,
    *,
    headers: list[str],
    per_column_validity: dict[int, tuple[object, object]],
    per_column_period: dict[int, str],
    metadata: SheetMetadata,
) -> pd.DataFrame:
    value_column_indices = [
        col_idx
        for col_idx, header in enumerate(headers)
        if _is_value_header_name(header) or _is_od_value_header_name(header)
    ]
    if len(value_column_indices) <= 1:
        return _attach_validity_columns(
            df,
            metadata,
            per_column_validity=per_column_validity,
            per_column_period=per_column_period,
            value_column_indices=value_column_indices,
        )

    value_headers = {headers[col_idx] for col_idx in value_column_indices}
    non_value_columns = [column for column in df.columns if column not in value_headers]
    parts: list[pd.DataFrame] = [df[non_value_columns].copy()]

    for value_idx, col_idx in enumerate(value_column_indices):
        header_name = headers[col_idx]
        if header_name not in df.columns:
            continue

        suffix = _column_suffix(value_idx)
        valid_from, valid_to = per_column_validity.get(
            col_idx,
            (metadata.valid_from, metadata.valid_to),
        )
        period_label = per_column_period.get(col_idx, "")

        group = pd.DataFrame(
            {
                f"{PERIOD_COLUMN}{suffix}": period_label,
                f"{VALID_FROM_COLUMN}{suffix}": _format_metadata_date(valid_from),
                f"{VALID_TO_COLUMN}{suffix}": _format_metadata_date(valid_to),
                header_name: df[header_name],
            }
        )
        parts.append(group)

    return pd.concat(parts, axis=1)


def _attach_rc_validity_columns(df: pd.DataFrame, metadata: SheetMetadata) -> pd.DataFrame:
    """Sheet-level RC agreement dates (distinct from per-period peak/non-peak validity)."""
    if metadata.valid_from is None and metadata.valid_to is None:
        return df

    enriched = df.copy()
    if metadata.valid_to is not None:
        enriched.insert(0, RC_VALID_TO_COLUMN, _format_metadata_date(metadata.valid_to))
    if metadata.valid_from is not None:
        enriched.insert(0, RC_VALID_FROM_COLUMN, _format_metadata_date(metadata.valid_from))
    return enriched


def _attach_validity_columns(
    df: pd.DataFrame,
    metadata: SheetMetadata,
    *,
    per_column_validity: dict[int, tuple[object, object]] | None = None,
    per_column_period: dict[int, str] | None = None,
    value_column_indices: list[int] | None = None,
) -> pd.DataFrame:


    per_column_validity = per_column_validity or {}
    per_column_period = per_column_period or {}

    if value_column_indices:
        col_idx = value_column_indices[0]
        valid_from, valid_to = per_column_validity.get(
            col_idx,
            (metadata.valid_from, metadata.valid_to),
        )
        period_label = per_column_period.get(col_idx, "")
    else:
        valid_from = metadata.valid_from
        valid_to = metadata.valid_to
        period_label = ""

    if (
        metadata.valid_from is None
        and metadata.valid_to is None
        and not per_column_validity
        and not per_column_period
    ):
        return df

    enriched = df.copy()
    if period_label:
        enriched.insert(0, PERIOD_COLUMN, period_label)
    enriched.insert(0, VALID_TO_COLUMN, _format_metadata_date(valid_to))
    enriched.insert(0, VALID_FROM_COLUMN, _format_metadata_date(valid_from))
    return enriched


def clean_reference_tab_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Drop preamble/empty rows and columns for free-form reference tabs."""
    cleaned = _drop_empty_rows(df_raw)
    cleaned = _drop_empty_columns(cleaned)
    return cleaned.reset_index(drop=True)


def clean_index_df(df_raw: pd.DataFrame, metadata: SheetMetadata) -> pd.DataFrame:
    header_row_idx = _find_header_row_by_marker(df_raw, "billing account")
    if header_row_idx is None:
        return clean_reference_tab_df(df_raw)

    headers = _normalize_headers(df_raw.iloc[header_row_idx].tolist())
    df = df_raw.iloc[header_row_idx + 1 :].copy()
    df.columns = headers
    df = _drop_empty_rows(df)
    df = _drop_empty_columns(df)
    return df.reset_index(drop=True)


def clean_additional_info_df(df_raw: pd.DataFrame, metadata: SheetMetadata) -> pd.DataFrame:
    sections: list[pd.DataFrame] = []
    current_title: str | None = metadata.section_titles[0] if metadata.section_titles else None
    current_headers: list[str] | None = None
    current_rows: list[list[object]] = []

    def flush_section() -> None:
        nonlocal current_rows, current_headers, current_title
        if current_headers is None or not current_rows:
            current_rows = []
            return

        section_df = pd.DataFrame(current_rows, columns=current_headers)
        section_df.insert(0, "Section Title", current_title or pd.NA)
        sections.append(section_df)
        current_rows = []

    for _, row in df_raw.iterrows():
        title = _find_section_title_in_row(row)
        if title is not None:
            flush_section()
            current_title = title
            current_headers = None
            continue

        if _is_additional_info_header_row(row):
            flush_section()
            current_headers = _normalize_headers(row.tolist())
            continue

        if current_headers is None:
            continue

        if any(_cell_text(value) for value in row):
            current_rows.append(row.tolist())

    flush_section()

    if not sections:
        return clean_reference_tab_df(df_raw)

    combined = pd.concat(sections, ignore_index=True)
    combined = _drop_empty_rows(combined)
    combined = _drop_empty_columns(combined)
    return combined.reset_index(drop=True)


def clean_standard_tab_df(df_raw: pd.DataFrame, metadata: SheetMetadata) -> pd.DataFrame:
    """Remove preamble rows above the header and fully empty rows/columns."""
    if df_raw.empty:
        return df_raw.copy()

    header_row_idx = find_header_row_index(df_raw)
    if header_row_idx is None:
        cleaned = _drop_empty_rows(df_raw)
        cleaned.columns = _normalize_headers(list(cleaned.columns))
        cleaned = cleaned.reset_index(drop=True)
        return _attach_validity_columns(cleaned, metadata)

    headers = _normalize_headers(df_raw.iloc[header_row_idx].tolist())
    per_column_validity = _extract_per_column_validity(df_raw, header_row_idx)
    per_column_period = _extract_per_column_period_labels(df_raw, header_row_idx)
    data_start = header_row_idx + 1
    if data_start < len(df_raw) and _is_internal_field_row(df_raw.iloc[data_start]):
        headers = _apply_internal_value_headers(headers, df_raw.iloc[data_start])
        data_start += 1

    df = df_raw.iloc[data_start:].copy()
    df.columns = headers
    df = _drop_empty_rows(df)
    df = _drop_empty_columns(df)
    df = df.reset_index(drop=True)

    value_headers = [
        header
        for header in headers
        if _is_value_header_name(header) or _is_od_value_header_name(header)
    ]
    if len(value_headers) > 1:
        result = _attach_per_value_column_validity(
            df,
            headers=headers,
            per_column_validity=per_column_validity,
            per_column_period=per_column_period,
            metadata=metadata,
        )
    else:
        result = _attach_validity_columns(
            df,
            metadata,
            per_column_validity=per_column_validity,
            per_column_period=per_column_period,
            value_column_indices=[
                col_idx
                for col_idx, header in enumerate(headers)
                if _is_value_header_name(header) or _is_od_value_header_name(header)
            ],
        )
    result = _attach_calc_rules(result, metadata.calc_rules)
    return _attach_rc_validity_columns(result, metadata)


def clean_tab_df(df_raw: pd.DataFrame, sheet_name: str | None = None) -> pd.DataFrame:
    metadata = extract_sheet_metadata(df_raw)
    if sheet_name:
        lower_name = sheet_name.strip().lower()
        if lower_name == "additional info":
            return clean_additional_info_df(df_raw, metadata)
        if lower_name == "index":
            return clean_index_df(df_raw, metadata)
    return clean_standard_tab_df(df_raw, metadata)


def find_header_row_index(df_raw: pd.DataFrame, max_scan_rows: int = 30) -> int | None:
    scan_limit = min(len(df_raw), max_scan_rows)
    for row_idx in range(scan_limit):
        if _row_contains_marker(df_raw.iloc[row_idx]):
            return row_idx
    return None


def _is_internal_field_row(row: pd.Series) -> bool:
    """Skip Geodis internal field-name row directly below display headers."""
    first_values = [_cell_text(value).lower() for value in row[:6]]
    return any(value in {"item_name", "service_type", "origin_region"} for value in first_values)


def tab_to_df(file_path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    return clean_tab_df(raw, sheet_name=sheet_name)


MAX_SHEET_NAME_LEN = 31


def sanitize_sheet_part(text: str) -> str:
    return re.sub(r"[\[\]:*?/\\]", "_", text).strip() or "Sheet"


def output_sheet_name(sheet_name: str, used: set[str]) -> str:
    base = sanitize_sheet_part(sheet_name)
    if base not in used:
        used.add(base)
        return base

    for n in range(2, 1000):
        suffix = f"_{n}"
        candidate = sanitize_sheet_part(sheet_name)[: MAX_SHEET_NAME_LEN - len(suffix)] + suffix
        if candidate not in used:
            used.add(candidate)
            return candidate

    raise RuntimeError(f"Could not create a unique sheet name for: {sheet_name}")


def collect_frames(
    file_path: Path,
    sheet_names: list[str],
) -> list[tuple[str, pd.DataFrame]]:
    frames: list[tuple[str, pd.DataFrame]] = []
    used_names: set[str] = set()

    for sheet_name in sheet_names:
        try:
            raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            metadata = extract_sheet_metadata(raw)
            df = clean_tab_df(raw, sheet_name=sheet_name)
            removed_rows = len(raw) - len(df)
        except Exception as exc:
            print(f"Skipping {sheet_name}: could not read sheet ({exc})")
            continue

        label = output_sheet_name(sheet_name, used_names)
        frames.append((label, df))
        meta_parts: list[str] = []
        if metadata.section_titles:
            meta_parts.append(f"sections: {', '.join(metadata.section_titles)}")
        if metadata.valid_from is not None or metadata.valid_to is not None:
            meta_parts.append(
                "validity: "
                f"{_format_metadata_date(metadata.valid_from)} to "
                f"{_format_metadata_date(metadata.valid_to)}"
            )
        if not metadata.calc_rules.is_empty:
            meta_parts.append(
                "calc rules: "
                f"{metadata.calc_rules.calculation_base} / "
                f"{metadata.calc_rules.fap_charge_code} / "
                f"{metadata.calc_rules.uom}"
            )
        meta_summary = f"; {'; '.join(meta_parts)}" if meta_parts else ""
        print(
            f"  Loaded: {sheet_name} -> '{label}' "
            f"({len(df)} rows, {len(df.columns)} columns; removed {removed_rows} preamble/empty rows"
            f"{meta_summary})"
        )

    return frames


def save_to_processing(
    output_stem: str,
    frames: list[tuple[str, pd.DataFrame]],
) -> Path:
    output_path = PROCESSING_DIR / f"{output_stem}_extracted.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in frames:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  Wrote tab: {sheet_name}")

    return output_path


def _convert_rate_file(
    *,
    auto: bool,
    file_path: Path | None,
) -> tuple[Path | None, list[tuple[str, pd.DataFrame]]]:
    files = list_rate_files()
    selected_file = file_path or select_input_file(files, auto=auto)

    if not selected_file.exists():
        raise FileNotFoundError(f"Rate file not found: {selected_file}")

    workbook = pd.ExcelFile(selected_file)
    if auto:
        selected_tabs = propose_default_tabs(workbook.sheet_names)
        if not selected_tabs:
            raise RuntimeError(
                f"No default tabs matched in {selected_file.name}. "
                "Run without --auto to choose tabs manually."
            )
        print(f"Auto mode: converting rate tabs: {', '.join(selected_tabs)}")
    else:
        selected_tabs = select_tabs_interactive(selected_file, workbook.sheet_names)

    print(f"\nConverting rate {_file_label(selected_file, RATE_INPUT_DIR)}...")
    frames = collect_frames(selected_file, selected_tabs)
    if not frames:
        raise RuntimeError("Nothing to save. No rate sheets could be loaded.")
    return selected_file, frames


def run_convert(
    *,
    auto: bool = False,
    file_path: Path | None = None,
    fsc_file_path: Path | None = None,
    source_mode: str | None = None,
) -> Path:
    ensure_workspace_dirs()
    from convert_fsc import run_convert_fsc

    mode = source_mode or prompt_source_mode(auto=auto)
    if mode not in SOURCE_MODE_CHOICES:
        raise ValueError(f"Invalid source mode: {mode}")

    frames: list[tuple[str, pd.DataFrame]] = []
    output_stem: str | None = None

    if mode in {SOURCE_MODE_RATE, SOURCE_MODE_BOTH}:
        rate_file, rate_frames = _convert_rate_file(auto=auto, file_path=file_path)
        frames.extend(rate_frames)
        output_stem = rate_file.stem

    if mode in {SOURCE_MODE_FSC, SOURCE_MODE_BOTH}:
        fsc_file, fsc_frames = run_convert_fsc(auto=auto, file_path=fsc_file_path)
        frames.extend(fsc_frames)
        if output_stem is None:
            output_stem = fsc_file.stem

    if not frames or output_stem is None:
        raise RuntimeError("Nothing to save.")

    output_path = save_to_processing(output_stem, frames)
    print(f"\nSaved {len(frames)} sheet(s) to: {output_path}")
    return output_path


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert GAR input Excel tabs to cleaned workbooks in processing/.",
    )
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Use default tabs and skip interactive prompts.",
    )
    parser.add_argument(
        "--file",
        type=Path,
        default=None,
        help="Optional rate workbook path (relative to input/rate/ or absolute).",
    )
    parser.add_argument(
        "--fsc-file",
        type=Path,
        default=None,
        help="Optional FSC workbook path (relative to input/fsc/ or absolute).",
    )
    parser.add_argument(
        "--source",
        choices=SOURCE_MODE_CHOICES,
        default=None,
        help="Conversion source: rate, fsc, or both.",
    )
    return parser.parse_args()


def _resolve_input_file(file_arg: Path | None, *, base_dir: Path) -> Path | None:
    if file_arg is None:
        return None
    if file_arg.is_file():
        return file_arg
    candidate = base_dir / file_arg
    if candidate.is_file():
        return candidate
    raise FileNotFoundError(f"Input file not found: {file_arg}")


def main() -> int:
    try:
        args = _parse_args()
        run_convert(
            auto=args.auto,
            file_path=_resolve_input_file(args.file, base_dir=RATE_INPUT_DIR),
            fsc_file_path=_resolve_input_file(args.fsc_file, base_dir=FSC_INPUT_DIR),
            source_mode=args.source,
        )
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
