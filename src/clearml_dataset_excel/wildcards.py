from __future__ import annotations

import glob
from fnmatch import fnmatch
from pathlib import Path


def matches_any_wildcard(path: str, wildcards: str | list[str] | None, *, recursive: bool = True) -> bool:
    path = Path(path).as_posix()
    if wildcards is None:
        wildcards = ["*"]
    if not isinstance(wildcards, list):
        wildcards = [wildcards]
    wildcards = [str(w) for w in wildcards]
    if not recursive:
        path_segments = path.split("/")
        for wildcard in wildcards:
            wildcard_segments = str(wildcard).split("/")
            if len(path_segments) != len(wildcard_segments):
                continue
            if all(fnmatch(p, w) for p, w in zip(path_segments, wildcard_segments)):
                return True
        return False

    for wildcard in wildcards:
        wildcard = str(wildcard)
        wildcard_file = wildcard.split("/")[-1]
        wildcard_dir = wildcard[: -len(wildcard_file)] + "*"
        if fnmatch(path, wildcard_dir) and fnmatch("/" + path, "*/" + wildcard_file):
            return True
    return False


def has_glob_magic(text: str) -> bool:
    return glob.has_magic(text)


def split_glob_root_and_pattern(pattern_path: Path) -> tuple[Path, str]:
    parts = pattern_path.parts
    for idx, part in enumerate(parts):
        if glob.has_magic(part):
            root = Path(*parts[:idx]) if idx > 0 else Path(pattern_path.anchor or ".")
            rel_pattern = Path(*parts[idx:]).as_posix()
            return root, rel_pattern
    raise ValueError(f"Not a glob pattern: {pattern_path}")

