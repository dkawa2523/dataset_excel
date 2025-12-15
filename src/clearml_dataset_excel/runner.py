from __future__ import annotations

import sys

from .cli import main


def run(argv: list[str] | None = None) -> int:
    """
    Entry point for packaging (e.g. PyInstaller) to provide a Python-less runner executable on Windows.
    """
    args = sys.argv[1:] if argv is None else list(argv)
    return int(main(args))


if __name__ == "__main__":
    raise SystemExit(run())

