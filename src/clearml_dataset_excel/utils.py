from __future__ import annotations

import sys
from pathlib import Path
from typing import Any


def non_empty_str(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and value != value:  # NaN
        return None
    text = str(value).strip()
    return text or None


def is_url(text: str) -> bool:
    lowered = text.lower()
    return lowered.startswith(("http://", "https://", "s3://", "gs://", "azure://", "file://"))


def try_relative_to(path: Path, base: Path) -> bool:
    try:
        path.relative_to(base)
        return True
    except ValueError:
        return False


def resolve_local_base_folder(source_path: Path, base_dir: Path | None) -> Path:
    if base_dir and try_relative_to(source_path, base_dir):
        return base_dir
    if source_path.is_dir():
        return source_path
    return source_path.parent


def get_macos_quarantine(path: Path) -> str | None:
    """
    Return the macOS quarantine xattr value (com.apple.quarantine) if present, else None.
    """
    if sys.platform != "darwin":
        return None
    try:
        import subprocess

        r = subprocess.run(
            ["xattr", "-p", "com.apple.quarantine", path.as_posix()],
            capture_output=True,
            text=True,
        )
        if r.returncode != 0:
            return None
        v = (r.stdout or "").strip()
        return v or None
    except (FileNotFoundError, OSError):
        return None


def clear_macos_quarantine(path: Path) -> bool:
    """
    Best-effort: remove macOS quarantine xattr (com.apple.quarantine).

    This is useful when Excel blocks macros in files flagged as downloaded from the internet.
    Returns True if the xattr was removed successfully.
    """
    if sys.platform != "darwin":
        return False
    try:
        import subprocess

        r = subprocess.run(
            ["xattr", "-d", "com.apple.quarantine", path.as_posix()],
            capture_output=True,
            text=True,
        )
        return r.returncode == 0
    except (FileNotFoundError, OSError):
        return False
