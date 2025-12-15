from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

from .expr import ExprError, eval_expr
from .format_spec import Aggregate, DatasetFormatSpec, FileSpec
from .manifest import read_rows_from_manifest
from .utils import is_url


class ProcessingError(RuntimeError):
    pass


@dataclass(frozen=True)
class ProcessOutputs:
    output_dir: Path
    canonical_csv: Path
    conditions_csv: Path
    consolidated_excel: Path
    uploaded_files: list[Path]
    raw_path_map: dict[str, Path]


def _coerce_condition_series(series, dtype: str):  # type: ignore[no-untyped-def]
    import pandas as pd

    dt = dtype.lower().strip()
    if dt in {"str", "string", "path"}:
        return series.apply(lambda v: None if v is None else str(v))
    if dt in {"int", "int64", "integer"}:
        return pd.to_numeric(series, errors="coerce").astype("Int64")
    if dt in {"float", "float64", "number"}:
        return pd.to_numeric(series, errors="coerce").astype("Float64")
    if dt in {"bool", "boolean"}:
        def _to_bool(v: Any) -> bool | None:
            if v is None:
                return None
            if isinstance(v, bool):
                return v
            if isinstance(v, (int, float)) and v in (0, 1):
                return bool(v)
            if isinstance(v, str):
                s = v.strip().lower()
                if s in {"true", "t", "1", "yes", "y"}:
                    return True
                if s in {"false", "f", "0", "no", "n"}:
                    return False
            return None

        return series.apply(_to_bool)
    if dt in {"date", "datetime"}:
        return pd.to_datetime(series, errors="coerce")
    return series


def _coerce_measure_series(series, dtype: str):  # type: ignore[no-untyped-def]
    import pandas as pd

    dt = dtype.lower().strip()
    if dt in {"str", "string"}:
        return series.apply(lambda v: None if v is None else str(v))
    if dt in {"int", "int64", "integer"}:
        return pd.to_numeric(series, errors="coerce").astype("Int64")
    if dt in {"float", "float64", "number"}:
        return pd.to_numeric(series, errors="coerce").astype("Float64")
    if dt in {"bool", "boolean"}:
        return _coerce_condition_series(series, "bool")
    if dt in {"date", "datetime"}:
        return pd.to_datetime(series, errors="coerce")
    return series


def _resolve_local_path(text: str, *, base_dir: Path) -> Path:
    expanded = Path(text).expanduser()
    if expanded.is_absolute():
        return expanded.resolve()
    return (base_dir / expanded).resolve()


def _fallback_resolve_missing_path(*, raw_path: str, search_root: Path) -> Path | None:
    """
    Best-effort path resolution when the path stored in Excel does not exist locally.
    This is mainly for ClearML Agent re-processing, where Excel might contain absolute
    paths from another machine.
    """
    t = str(raw_path).strip()
    if not t:
        return None

    normalized = t.replace("\\", "/")

    # 1) Try interpreting as a relative path under search_root.
    first_segment = normalized.split("/", 1)[0]
    if not normalized.startswith("/") and not first_segment.endswith(":"):
        candidate = (search_root / Path(normalized)).resolve()
        if candidate.exists():
            return candidate

    # 2) Search by basename. Also support ClearML staging external naming: 000_<basename>.
    basename = normalized.split("/")[-1]
    if not basename:
        return None

    patterns = [basename, f"*_{basename}"]
    matches: list[Path] = []
    for pat in patterns:
        for p in search_root.rglob(pat):
            if p.is_file():
                matches.append(p.resolve())

    unique: list[Path] = []
    seen: set[str] = set()
    for p in matches:
        s = p.as_posix()
        if s not in seen:
            seen.add(s)
            unique.append(p)

    if len(unique) == 1:
        return unique[0]
    return None


def _read_measurement_file(path: Path, file_spec: FileSpec):  # type: ignore[no-untyped-def]
    import pandas as pd

    fmt = file_spec.format.lower().strip()
    if fmt in {"csv", "tsv"}:
        sep = "\t" if fmt == "tsv" else ","
        opts = dict(file_spec.read)
        opts.setdefault("sep", sep)
        return pd.read_csv(path, **opts)
    if fmt in {"xlsx", "xlsm", "xls", "excel"}:
        opts = dict(file_spec.read)
        if file_spec.sheet:
            opts.setdefault("sheet_name", file_spec.sheet)
        return pd.read_excel(path, engine="openpyxl", engine_kwargs={"data_only": True}, **opts)
    if fmt in {"image", "jpg", "jpeg", "png", "tif", "tiff", "bmp"}:
        return None
    raise ProcessingError(f"Unsupported format for {file_spec.file_id}: {file_spec.format}")


def _normalize_measurement_df(df, file_spec: FileSpec):  # type: ignore[no-untyped-def]
    import pandas as pd

    if df is None:
        return None

    df = df.where(df.notnull(), None)

    out = pd.DataFrame()
    for axis in ("x", "y", "z", "t"):
        src = getattr(file_spec.axes, axis)
        dtype = getattr(file_spec.axis_types, axis, None)
        if src is None:
            out[axis] = pd.NA
        else:
            if src not in df.columns:
                raise ProcessingError(f"{file_spec.file_id}: missing axis column in file: {src}")
            s = df[src]
            if isinstance(dtype, str) and dtype.strip():
                out[axis] = _coerce_measure_series(s, dtype)
            else:
                out[axis] = s

    for t in file_spec.targets:
        if t.source not in df.columns:
            raise ProcessingError(f"{file_spec.file_id}: missing target source column in file: {t.source}")
        out[t.name] = _coerce_measure_series(df[t.source], t.dtype)

    # Derived columns evaluated on canonical namespace (axes + targets + previous derived)
    for d in file_spec.derived:
        namespace = {k: out[k] for k in out.columns}
        try:
            v = eval_expr(d.expr, namespace)
        except ExprError as e:
            raise ProcessingError(f"{file_spec.file_id}: derived '{d.name}' failed: {e}") from e
        out[d.name] = _coerce_measure_series(v, d.dtype)  # type: ignore[arg-type]

    return out


def _compute_aggregate(df, agg: Aggregate):  # type: ignore[no-untyped-def]
    import numpy as np
    import pandas as pd

    if agg.source not in df.columns:
        raise ProcessingError(f"aggregate '{agg.name}': unknown source column: {agg.source}")

    op = agg.op.lower().strip()
    s = pd.to_numeric(df[agg.source], errors="coerce")

    if op == "mean":
        return float(s.mean())
    if op == "max":
        return float(s.max())
    if op == "min":
        return float(s.min())
    if op == "sum":
        return float(s.sum())
    if op in {"integral", "integrate", "integrate_trapz", "trapz"}:
        wrt = agg.wrt or "t"
        if wrt not in df.columns:
            raise ProcessingError(f"aggregate '{agg.name}': wrt column not found: {wrt}")
        x = pd.to_numeric(df[wrt], errors="coerce").to_numpy(dtype=float)
        y = s.to_numpy(dtype=float)
        mask = ~(np.isnan(x) | np.isnan(y))
        if int(mask.sum()) < 2:
            return float("nan")
        x2 = x[mask]
        y2 = y[mask]
        order = np.argsort(x2)
        trapezoid = getattr(np, "trapezoid", np.trapz)
        return float(trapezoid(y2[order], x2[order]))

    raise ProcessingError(f"aggregate '{agg.name}': unsupported op: {agg.op}")


def process_condition_excel(
    spec: DatasetFormatSpec,
    excel_path: str | Path,
    *,
    sheet_name: str | None = None,
    output_root: str | Path | None = None,
    check_files_exist: bool = True,
    path_fallback_map: Mapping[str, str] | None = None,
    fallback_search_root: str | Path | None = None,
) -> ProcessOutputs:
    """
    Read a filled condition Excel and referenced measurement files, then produce:
    - conditions.csv (original + aggregates)
    - canonical.csv (long table with x/y/z/t and all targets)
    - consolidated.xlsx (Conditions + Canonical sheets)
    """
    import pandas as pd

    excel_path = Path(excel_path).expanduser().resolve()
    if not excel_path.exists():
        raise ProcessingError(f"Condition Excel not found: {excel_path}")

    rows, _ = read_rows_from_manifest(excel_path, sheet_name or spec.template.condition_sheet)
    cond_df = pd.DataFrame(rows)

    # Ensure all condition columns exist
    col_names = {c.name for c in spec.condition_columns}
    missing_cols = sorted(col_names - set(cond_df.columns))
    if missing_cols:
        raise ProcessingError(f"Missing condition columns in Excel: {missing_cols}")

    # Coerce and validate required columns
    for c in spec.condition_columns:
        cond_df[c.name] = _coerce_condition_series(cond_df[c.name], c.dtype)
        if c.required:
            missing_mask = cond_df[c.name].isna() | (cond_df[c.name].astype(str).str.strip() == "")
            if bool(missing_mask.any()):
                bad_rows = missing_mask[missing_mask].index.to_list()[:20]
                raise ProcessingError(f"Required column '{c.name}' has missing values at rows: {bad_rows}")

    output_root_path = Path(output_root).expanduser().resolve() if output_root else excel_path.parent
    output_dir = output_root_path / spec.output.output_dirname
    output_dir.mkdir(parents=True, exist_ok=True)

    fallback_root_path = Path(fallback_search_root).expanduser().resolve() if fallback_search_root else None

    file_path_cols = spec.file_path_columns()

    # Prepare conditions output (optionally drop path columns)
    if spec.output.include_file_path_columns:
        conditions_out = cond_df.copy()
    else:
        conditions_out = cond_df.drop(columns=sorted(file_path_cols), errors="ignore").copy()

    canonical_frames: list[pd.DataFrame] = []
    uploaded_files: list[Path] = [excel_path]
    raw_path_map: dict[str, Path] = {}

    for row_idx, row in cond_df.iterrows():
        row_dict = row.to_dict()

        per_file_frames: list[tuple[FileSpec, pd.DataFrame]] = []
        per_row_aggs: dict[str, Any] = {}

        for f in spec.files:
            raw_path = row_dict.get(f.path_column)
            if not isinstance(raw_path, str):
                raw_path = "" if raw_path is None else str(raw_path)
            raw_path = raw_path.strip()
            if not raw_path:
                continue
            if is_url(raw_path):
                raise ProcessingError(f"Row {row_idx}: URL paths are not supported (upload-all required): {raw_path}")

            resolved = _resolve_local_path(raw_path, base_dir=excel_path.parent)
            if check_files_exist and not resolved.exists():
                if path_fallback_map is not None:
                    alt_text = path_fallback_map.get(raw_path)
                    if alt_text is None:
                        alt_text = path_fallback_map.get(raw_path.replace("\\", "/"))
                    if alt_text is None:
                        alt_text = path_fallback_map.get(raw_path.replace("/", "\\"))
                    if isinstance(alt_text, str) and alt_text.strip():
                        alt_path = Path(alt_text).expanduser()
                        if not alt_path.is_absolute():
                            alt_path = (excel_path.parent / alt_path).resolve()
                        else:
                            alt_path = alt_path.resolve()
                        if alt_path.exists():
                            resolved = alt_path
                if fallback_root_path is not None:
                    alt = _fallback_resolve_missing_path(raw_path=raw_path, search_root=fallback_root_path)
                    if alt is not None and alt.exists():
                        resolved = alt
                if not resolved.exists():
                    raise ProcessingError(f"Row {row_idx}: file not found for column '{f.path_column}': {resolved}")
            prev = raw_path_map.get(raw_path)
            if prev is None:
                raw_path_map[raw_path] = resolved
            elif prev.resolve() != resolved.resolve():
                raise ProcessingError(
                    f"Row {row_idx}: same path value resolves to different files: {raw_path} -> {prev} vs {resolved}"
                )
            uploaded_files.append(resolved)

            df_src = _read_measurement_file(resolved, f)
            df_norm = _normalize_measurement_df(df_src, f)
            if df_norm is None:
                continue

            # Aggregates
            for agg in f.aggregates:
                out_col = agg.output_column or agg.name
                key = out_col
                if key in per_row_aggs:
                    raise ProcessingError(f"Row {row_idx}: duplicate aggregate output column: {key}")
                per_row_aggs[key] = _compute_aggregate(df_norm, agg)

            per_file_frames.append((f, df_norm))

        # Write aggregates into conditions output
        for k, v in per_row_aggs.items():
            if k not in conditions_out.columns:
                conditions_out[k] = pd.NA
            conditions_out.at[row_idx, k] = v

        if not per_file_frames:
            continue

        combine_mode = (spec.output.combine_mode or "auto").strip().lower()
        if combine_mode not in {"auto", "merge", "append"}:
            raise ProcessingError(f"Invalid output.combine_mode: {combine_mode}")

        # Group frames by their defined axes set (per file)
        groups: dict[frozenset[str], list[tuple[FileSpec, pd.DataFrame]]] = {}
        for f, df in per_file_frames:
            groups.setdefault(frozenset(f.axes.defined_axes()), []).append((f, df))

        if combine_mode == "merge" and len(groups) != 1:
            details = {f.file_id: sorted(f.axes.defined_axes()) for f, _ in per_file_frames}
            raise ProcessingError(f"Row {row_idx}: axes mismatch across files (combine_mode=merge): {details}")

        merged_or_appended: list[pd.DataFrame] = []
        for axes_key, items_in_group in groups.items():
            join_keys = [a for a in ("x", "y", "z", "t") if a in set(axes_key)]

            # Decide whether to merge within the group
            do_merge = combine_mode in {"merge", "auto"} and len(items_in_group) > 1 and bool(join_keys)
            if not do_merge:
                merged_or_appended.extend([df.copy() for _, df in items_in_group])
                continue

            merged = None
            for f, df in items_in_group:
                keep_cols = join_keys + [c for c in df.columns if c not in {"x", "y", "z", "t"}]
                df2 = df[keep_cols].copy()
                if df2.duplicated(subset=join_keys).any():
                    raise ProcessingError(
                        f"Row {row_idx}: non-unique points for join keys {join_keys} in file '{f.file_id}'"
                    )
                if merged is None:
                    merged = df2
                else:
                    merged = merged.merge(df2, on=join_keys, how="outer")

            assert merged is not None
            # Re-add missing axes columns (as NaN) to match x/y/z/t contract
            for axis in ("x", "y", "z", "t"):
                if axis not in merged.columns:
                    merged[axis] = pd.NA
            merged_or_appended.append(merged)

        attach_cols = [
            c.name
            for c in spec.condition_columns
            if spec.output.include_file_path_columns or c.name not in file_path_cols
        ]

        for df in merged_or_appended:
            for c in attach_cols:
                df[c] = row_dict.get(c)
            for k, v in per_row_aggs.items():
                df[k] = v
            df["__condition_row"] = int(row_idx)
            canonical_frames.append(df)

    if canonical_frames:
        canonical_out = pd.concat(canonical_frames, ignore_index=True)
    else:
        canonical_out = pd.DataFrame(columns=["x", "y", "z", "t"] + [c.name for c in spec.condition_columns])

    canonical_csv = output_dir / spec.output.canonical_filename
    conditions_csv = output_dir / spec.output.conditions_filename
    consolidated_excel = output_dir / spec.output.consolidated_excel_filename

    conditions_out.to_csv(conditions_csv, index=False, encoding="utf-8")
    canonical_out.to_csv(canonical_csv, index=False, encoding="utf-8")

    with pd.ExcelWriter(consolidated_excel, engine="openpyxl") as writer:
        conditions_out.to_excel(writer, index=False, sheet_name="Conditions")
        canonical_out.to_excel(writer, index=False, sheet_name="Canonical")

    # De-duplicate upload list while preserving order
    seen: set[str] = set()
    unique_files: list[Path] = []
    for p in uploaded_files:
        key = p.resolve().as_posix()
        if key in seen:
            continue
        seen.add(key)
        unique_files.append(p)

    return ProcessOutputs(
        output_dir=output_dir,
        canonical_csv=canonical_csv,
        conditions_csv=conditions_csv,
        consolidated_excel=consolidated_excel,
        uploaded_files=unique_files,
        raw_path_map=raw_path_map,
    )
