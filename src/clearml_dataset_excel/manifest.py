from __future__ import annotations

import csv
import sys
from pathlib import Path
from typing import Any, Mapping, Sequence


def write_manifest_csv(path: Path, rows: Sequence[Mapping[str, Any]]) -> None:
    keys: list[str] = []
    seen = set()
    for r in rows:
        for k in r.keys():
            if k not in seen:
                seen.add(k)
                keys.append(str(k))

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            writer.writerow({str(k): ("" if v is None else v) for k, v in r.items()})


def read_rows_from_manifest(manifest_path: Path, sheet_name: str | None) -> tuple[list[dict[str, Any]], list[str]]:
    try:
        import pandas as pd
    except Exception as e:  # pragma: no cover
        raise RuntimeError("pandas is required. Install dependencies first.") from e

    suffix = manifest_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        kwargs: dict[str, Any] = {}
        if sheet_name:
            kwargs["sheet_name"] = sheet_name
        try:
            # Use cached values for formulas (Excel add-ins/custom functions are not evaluated here).
            df = pd.read_excel(
                manifest_path,
                engine="openpyxl",
                engine_kwargs={"data_only": True},
                **kwargs,
            )
        except (ImportError, ModuleNotFoundError) as e:
            raise RuntimeError("Reading .xlsx requires 'openpyxl'. Install dependencies first.") from e
        except ValueError as e:
            raise RuntimeError(str(e)) from e
    elif suffix in {".csv", ".tsv"}:
        if sheet_name:
            print("Warning: --sheet is ignored for .csv/.tsv input", file=sys.stderr)
        sep = "\t" if suffix == ".tsv" else ","
        df = pd.read_csv(manifest_path, sep=sep)
    else:
        raise RuntimeError(f"Unsupported manifest extension: {manifest_path.suffix}")

    df = df.where(df.notnull(), None)
    rows: list[dict[str, Any]] = df.to_dict(orient="records")
    return rows, [str(c) for c in df.columns.to_list()]
