"""Project folder paths for GAR rate conversion."""

from __future__ import annotations

import os
from pathlib import Path

_DEFAULT_ROOT = Path(__file__).resolve().parent

ROOT = _DEFAULT_ROOT
INPUT_DIR = ROOT / "input"
RATE_INPUT_DIR = INPUT_DIR / "rate"
FSC_INPUT_DIR = INPUT_DIR / "fsc"
PROCESSING_DIR = ROOT / "processing"
OUTPUT_DIR = ROOT / "output"

# Default Google Drive location for RMT GAR data in Colab.
COLAB_DRIVE_BASE = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Apple/RMT_GAR"
)


def _path_from_env(name: str) -> Path | None:
    value = os.environ.get(name)
    if not value:
        return None
    return Path(value).expanduser()


def _colab_drive_base() -> Path | None:
    candidates: list[Path] = []
    env_base = _path_from_env("GAR_DRIVE_BASE")
    if env_base is not None:
        candidates.append(env_base)
    candidates.append(COLAB_DRIVE_BASE)

    for candidate in candidates:
        if (candidate / "input").is_dir():
            return candidate
    return None


def configure_paths(
    *,
    root: Path | str | None = None,
    input_dir: Path | str | None = None,
    rate_input_dir: Path | str | None = None,
    fsc_input_dir: Path | str | None = None,
    processing_dir: Path | str | None = None,
    output_dir: Path | str | None = None,
) -> None:
    """Override data folder locations (for Colab / Google Drive)."""
    global ROOT, INPUT_DIR, RATE_INPUT_DIR, FSC_INPUT_DIR, PROCESSING_DIR, OUTPUT_DIR

    if root is not None:
        ROOT = Path(root).expanduser().resolve()
    if input_dir is not None:
        INPUT_DIR = Path(input_dir).expanduser().resolve()
    if processing_dir is not None:
        PROCESSING_DIR = Path(processing_dir).expanduser().resolve()
    if output_dir is not None:
        OUTPUT_DIR = Path(output_dir).expanduser().resolve()

    if rate_input_dir is not None:
        RATE_INPUT_DIR = Path(rate_input_dir).expanduser().resolve()
    else:
        RATE_INPUT_DIR = INPUT_DIR / "rate"

    if fsc_input_dir is not None:
        FSC_INPUT_DIR = Path(fsc_input_dir).expanduser().resolve()
    else:
        FSC_INPUT_DIR = INPUT_DIR / "fsc"


def configure_paths_from_env() -> None:
    """Apply GAR_* environment variables and Colab Drive defaults when available."""
    root = _path_from_env("GAR_ROOT")
    input_dir = _path_from_env("GAR_INPUT_DIR")
    rate_input_dir = _path_from_env("GAR_RATE_INPUT_DIR")
    fsc_input_dir = _path_from_env("GAR_FSC_INPUT_DIR")
    processing_dir = _path_from_env("GAR_PROCESSING_DIR")
    output_dir = _path_from_env("GAR_OUTPUT_DIR")

    if input_dir is None and processing_dir is None and output_dir is None:
        drive_base = _colab_drive_base()
        if drive_base is not None:
            input_dir = drive_base / "input"
            processing_dir = drive_base / "processing"
            output_dir = drive_base / "output"

    if any(
        path is not None
        for path in (root, input_dir, rate_input_dir, fsc_input_dir, processing_dir, output_dir)
    ):
        configure_paths(
            root=root or ROOT,
            input_dir=input_dir or (root / "input" if root is not None else INPUT_DIR),
            rate_input_dir=rate_input_dir,
            fsc_input_dir=fsc_input_dir,
            processing_dir=processing_dir or (root / "processing" if root is not None else PROCESSING_DIR),
            output_dir=output_dir or (root / "output" if root is not None else OUTPUT_DIR),
        )


def ensure_workspace_dirs() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    RATE_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    FSC_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def print_path_config() -> None:
    print("GAR paths:")
    print(f"  Root:       {ROOT}")
    print(f"  Input:      {INPUT_DIR}")
    print(f"  Rate input: {RATE_INPUT_DIR}")
    print(f"  FSC input:  {FSC_INPUT_DIR}")
    print(f"  Processing: {PROCESSING_DIR}")
    print(f"  Output:     {OUTPUT_DIR}")
