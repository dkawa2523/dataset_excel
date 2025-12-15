from __future__ import annotations

import sys
from pathlib import Path


def _ensure_src_on_path() -> None:
    repo_root = Path(__file__).resolve().parent
    src = repo_root / "src"
    if src.exists():
        sys.path.insert(0, src.as_posix())


def main(argv: list[str] | None = None) -> int:
    _ensure_src_on_path()
    from clearml_dataset_excel.cli import main as cli_main

    args = list(argv) if argv is not None else sys.argv[1:]
    return int(cli_main(["agent", "reprocess", *args]))


if __name__ == "__main__":
    raise SystemExit(main())

