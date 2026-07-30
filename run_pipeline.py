"""
GAR — end-to-end rate pipeline.

Steps:
  1. Select source (rate / FSC / both) and convert to processing/
  2. Build matrix workbook in output/
  3. Build postal code zones txt from Additional Info and highlight matrix cities

Usage (local):
  python run_pipeline.py
  python run_pipeline.py --auto
  python run_pipeline.py --convert-only
  python run_pipeline.py --matrix-only
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from pathlib import Path

_CODE_DIR = Path(os.environ.get("GAR_CODE_DIR", "/content/Apple-GAR")).resolve()
try:
    _CODE_DIR = Path(__file__).resolve().parent
except NameError:
    pass
if str(_CODE_DIR) not in sys.path:
    sys.path.insert(0, str(_CODE_DIR))

_PIPELINE_MODULES = (
    "project_paths",
    "build_matrix",
    "build_postal_code_zones",
    "convert_fsc",
    "convert_to_processing",
)


def _bootstrap_paths():
    for module_name in _PIPELINE_MODULES:
        sys.modules.pop(module_name, None)

    import project_paths

    project_paths.configure_paths_from_env()
    return project_paths


_project_paths = _bootstrap_paths()
configure_paths_from_env = _project_paths.configure_paths_from_env
print_path_config = _project_paths.print_path_config

from build_matrix import run_build_matrix
from build_postal_code_zones import resolve_rate_workbook, run_build_postal_code_zones
from convert_to_processing import run_convert


@dataclass(frozen=True)
class PipelineResult:
    processing_path: Path | None
    output_path: Path | None
    postal_zones_path: Path | None = None


def run_pipeline(
    *,
    auto: bool = False,
    convert_only: bool = False,
    matrix_only: bool = False,
    processing_path: Path | None = None,
    output_path: Path | None = None,
    source_mode: str | None = None,
    rate_file_path: Path | None = None,
    fsc_file_path: Path | None = None,
) -> PipelineResult:
    if convert_only and matrix_only:
        raise ValueError("Use only one of --convert-only or --matrix-only.")

    configure_paths_from_env()
    print_path_config()
    saved_processing_path = processing_path
    saved_output_path: Path | None = None
    saved_matrix_df = None
    saved_postal_path: Path | None = None

    if not matrix_only:
        print("\n=== Step 1/3: Convert input to processing ===")
        saved_processing_path = run_convert(
            auto=auto,
            file_path=rate_file_path,
            fsc_file_path=fsc_file_path,
            source_mode=source_mode,
        )

    if not convert_only:
        matrix_step = "Step 2/3" if not matrix_only else "Step 1/2"
        postal_step = "Step 3/3" if not matrix_only else "Step 2/2"
        print(f"\n=== {matrix_step}: Build matrix ===")
        saved_output_path, saved_matrix_df = run_build_matrix(
            source_file=saved_processing_path,
            output_path=output_path,
            auto=auto or matrix_only,
        )

        print(f"\n=== {postal_step}: Build postal code zones ===")
        if saved_processing_path is None:
            raise RuntimeError("Processing workbook path is required to build postal code zones.")
        rate_workbook = resolve_rate_workbook(saved_processing_path, rate_file_path)
        saved_postal_path = run_build_postal_code_zones(
            rate_file_path=rate_workbook,
            matrix_path=saved_output_path,
            matrix_df=saved_matrix_df,
        )

    print("\n=== Pipeline complete ===")
    if saved_processing_path is not None:
        print(f"  Processing workbook: {saved_processing_path}")
    if saved_output_path is not None:
        print(f"  Output workbook:     {saved_output_path}")
    if saved_postal_path is not None:
        print(f"  Postal zones:        {saved_postal_path}")

    return PipelineResult(
        processing_path=saved_processing_path,
        output_path=saved_output_path,
        postal_zones_path=saved_postal_path,
    )


def _parse_args() -> argparse.Namespace:
    from convert_to_processing import SOURCE_MODE_CHOICES

    parser = argparse.ArgumentParser(description="Run the GAR end-to-end rate pipeline.")
    parser.add_argument("--auto", action="store_true", help="Skip prompts where possible.")
    parser.add_argument("--convert-only", action="store_true", help="Only run input -> processing.")
    parser.add_argument("--matrix-only", action="store_true", help="Only build matrix from processing file.")
    parser.add_argument("--processing", type=Path, default=None, help="Processing workbook for --matrix-only.")
    parser.add_argument("--output", type=Path, default=None, help="Optional output matrix workbook path.")
    parser.add_argument("--source", choices=SOURCE_MODE_CHOICES, default=None, help="rate, fsc, or both.")
    parser.add_argument("--file", type=Path, default=None, help="Rate file for conversion.")
    parser.add_argument("--fsc-file", type=Path, default=None, help="FSC file for conversion.")
    return parser.parse_args()


def _env_flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "on"}


def _running_in_notebook() -> bool:
    if "colab_kernel_launcher" in Path(sys.argv[0]).name:
        return True
    if any(arg == "-f" for arg in sys.argv):
        return True
    return "ipykernel" in sys.modules or "IPython" in sys.modules


def main() -> int:
    try:
        args = _parse_args()
        from convert_to_processing import _resolve_input_file
        from project_paths import FSC_INPUT_DIR, RATE_INPUT_DIR

        run_pipeline(
            auto=args.auto,
            convert_only=args.convert_only,
            matrix_only=args.matrix_only,
            processing_path=args.processing,
            output_path=args.output,
            source_mode=args.source,
            rate_file_path=_resolve_input_file(args.file, base_dir=RATE_INPUT_DIR),
            fsc_file_path=_resolve_input_file(args.fsc_file, base_dir=FSC_INPUT_DIR),
        )
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    if _running_in_notebook():
        from convert_to_processing import _resolve_input_file
        from project_paths import FSC_INPUT_DIR, RATE_INPUT_DIR

        run_pipeline(
            auto=_env_flag("GAR_AUTO"),
            convert_only=_env_flag("GAR_CONVERT_ONLY"),
            matrix_only=_env_flag("GAR_MATRIX_ONLY"),
            source_mode=os.environ.get("GAR_SOURCE"),
            rate_file_path=_resolve_input_file(
                Path(os.environ["GAR_RATE_FILE"]) if os.environ.get("GAR_RATE_FILE") else None,
                base_dir=RATE_INPUT_DIR,
            ),
            fsc_file_path=_resolve_input_file(
                Path(os.environ["GAR_FSC_FILE"]) if os.environ.get("GAR_FSC_FILE") else None,
                base_dir=FSC_INPUT_DIR,
            ),
        )
    else:
        raise SystemExit(main())
