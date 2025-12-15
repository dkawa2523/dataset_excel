from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from .template import render_dataset_path_template
from .utils import is_url, non_empty_str, resolve_local_base_folder
from .wildcards import has_glob_magic, matches_any_wildcard, split_glob_root_and_pattern


@dataclass(frozen=True)
class ResolvedItem:
    source: str
    dataset_path: str | None
    local_base_folder: str | None
    wildcard: str | None


def resolve_items(
    rows: Iterable[Mapping[str, Any]],
    *,
    path_col: str,
    dataset_path_col: str | None,
    dataset_path_template: str | None,
    base_dir: Path | None,
    skip_missing: bool,
) -> tuple[list[ResolvedItem], int]:
    items: list[ResolvedItem] = []
    skipped = 0
    for i, row in enumerate(rows, start=1):
        raw_path = non_empty_str(row.get(path_col))
        if not raw_path:
            continue

        resolved_dataset_path: str | None
        source_text = raw_path
        if is_url(source_text):
            if dataset_path_template:
                resolved_dataset_path = render_dataset_path_template(
                    dataset_path_template,
                    row,
                    row_index=i,
                    source_text=source_text,
                    source_path=None,
                    base_dir=None,
                )
            elif dataset_path_col:
                resolved_dataset_path = non_empty_str(row.get(dataset_path_col))
                if resolved_dataset_path:
                    resolved_dataset_path = resolved_dataset_path.lstrip("/")
            else:
                resolved_dataset_path = None
            items.append(
                ResolvedItem(
                    source=source_text,
                    dataset_path=resolved_dataset_path,
                    local_base_folder=None,
                    wildcard=None,
                )
            )
            continue

        source_path = Path(source_text)
        if not source_path.is_absolute() and base_dir:
            source_path = (base_dir / source_path).expanduser().resolve()
        else:
            source_path = source_path.expanduser().resolve()

        wildcard: str | None = None
        template_source_path: Path = source_path
        if has_glob_magic(source_text):
            glob_root, wildcard = split_glob_root_and_pattern(source_path)
            template_source_path = glob_root
            if not glob_root.exists() or not glob_root.is_dir():
                if skip_missing:
                    print(f"Warning: Row {i}: glob root not found, skipped: {glob_root}", file=sys.stderr)
                    skipped += 1
                    continue
                raise FileNotFoundError(f"Row {i}: glob root not found: {glob_root}")

        if dataset_path_template:
            resolved_dataset_path = render_dataset_path_template(
                dataset_path_template,
                row,
                row_index=i,
                source_text=source_text,
                source_path=template_source_path,
                base_dir=base_dir,
            )
        elif dataset_path_col:
            resolved_dataset_path = non_empty_str(row.get(dataset_path_col))
            if resolved_dataset_path:
                resolved_dataset_path = resolved_dataset_path.lstrip("/")
        else:
            resolved_dataset_path = None

        if wildcard is not None:
            local_base_folder = resolve_local_base_folder(template_source_path, base_dir)
            items.append(
                ResolvedItem(
                    source=template_source_path.as_posix(),
                    dataset_path=resolved_dataset_path,
                    local_base_folder=local_base_folder.as_posix(),
                    wildcard=wildcard,
                )
            )
            continue

        if not source_path.exists():
            if skip_missing:
                print(f"Warning: Row {i}: path not found, skipped: {source_path}", file=sys.stderr)
                skipped += 1
                continue
            raise FileNotFoundError(f"Row {i}: path not found: {source_path}")

        local_base_folder = resolve_local_base_folder(source_path, base_dir)
        items.append(
            ResolvedItem(
                source=source_path.as_posix(),
                dataset_path=resolved_dataset_path,
                local_base_folder=local_base_folder.as_posix(),
                wildcard=None,
            )
        )

    return items, skipped


def iter_local_files(path: Path, *, recursive: bool) -> list[Path]:
    if path.is_file():
        return [path]
    if not path.is_dir():
        return []
    iterator = path.rglob("*") if recursive else path.glob("*")
    return sorted({p for p in iterator if p.is_file()})


def iter_local_files_with_wildcards(path: Path, *, wildcards: list[str] | None, recursive: bool) -> list[Path]:
    if path.is_file():
        return [path]
    if not path.is_dir():
        return []
    if not wildcards:
        return iter_local_files(path, recursive=recursive)

    files: set[Path] = set()
    for w in wildcards:
        iterator = path.rglob(w) if recursive else path.glob(w)
        files.update({p for p in iterator if p.is_file()})
    return sorted(files)


def calc_dataset_relpath(*, file_path: Path, local_base_folder: Path, dataset_path: str | None) -> str:
    prefix = Path(dataset_path or ".")
    return (prefix / file_path.relative_to(local_base_folder)).as_posix()


def collect_local_dataset_paths(
    items: Sequence[ResolvedItem],
    *,
    recursive: bool,
    include: list[str] | None,
    exclude: list[str] | None,
) -> tuple[dict[str, str], dict[str, list[str]], int, int]:
    unique_dataset_paths: dict[str, str] = {}
    collisions: dict[str, list[str]] = {}
    matched_local_files = 0
    excluded_local_files = 0

    for it in items:
        if is_url(it.source):
            continue

        source_path = Path(it.source)
        local_base_folder = Path(it.local_base_folder or source_path.parent)

        item_wildcards: list[str] | None
        if it.wildcard:
            item_wildcards = [it.wildcard]
        elif source_path.is_dir() and include:
            item_wildcards = include
        else:
            item_wildcards = None

        if source_path.is_file():
            rel = source_path.relative_to(local_base_folder).as_posix()
            if include and not matches_any_wildcard(rel, include, recursive=recursive):
                files: list[Path] = []
            else:
                files = [source_path]
        else:
            files = iter_local_files_with_wildcards(source_path, wildcards=item_wildcards, recursive=recursive)

        matched_local_files += len(files)
        for f in files:
            dataset_relpath = calc_dataset_relpath(
                file_path=f, local_base_folder=local_base_folder, dataset_path=it.dataset_path
            )
            if exclude and matches_any_wildcard(dataset_relpath, exclude, recursive=recursive):
                excluded_local_files += 1
                continue

            existing = unique_dataset_paths.get(dataset_relpath)
            if existing and existing != f.as_posix():
                collisions.setdefault(dataset_relpath, [existing]).append(f.as_posix())
            else:
                unique_dataset_paths.setdefault(dataset_relpath, f.as_posix())

    return unique_dataset_paths, collisions, matched_local_files, excluded_local_files

