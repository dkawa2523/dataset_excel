from __future__ import annotations

from pathlib import Path
from urllib.parse import urlparse
from typing import Any, Mapping

from .utils import is_url, try_relative_to


def path_parts_for_template(source_text: str, *, source_path: Path | None, base_dir: Path | None) -> dict[str, str]:
    if is_url(source_text):
        parsed = urlparse(source_text)
        parsed_path = Path(parsed.path)
        basename = parsed_path.name
        stem = parsed_path.stem
        suffix = parsed_path.suffix
        relpath = parsed.path.lstrip("/") or basename
        return {
            "basename": basename,
            "stem": stem,
            "suffix": suffix,
            "relpath": relpath,
            "url": source_text,
            "scheme": parsed.scheme,
            "netloc": parsed.netloc,
        }

    if not source_path:
        source_path = Path(source_text)

    basename = source_path.name
    stem = source_path.stem
    suffix = source_path.suffix
    relpath: str = basename
    if base_dir and try_relative_to(source_path, base_dir):
        relpath = source_path.relative_to(base_dir).as_posix()
    return {
        "basename": basename,
        "stem": stem,
        "suffix": suffix,
        "relpath": relpath,
    }


def render_dataset_path_template(
    template: str,
    row: Mapping[str, Any],
    *,
    row_index: int,
    source_text: str,
    source_path: Path | None,
    base_dir: Path | None,
) -> str:
    values = dict(row)
    values.update(path_parts_for_template(source_text, source_path=source_path, base_dir=base_dir))
    try:
        rendered = template.format(**values)
    except KeyError as e:
        key = e.args[0]
        raise ValueError(f"Row {row_index}: dataset path template is missing key: {key}") from e
    return rendered.strip().lstrip("/")

