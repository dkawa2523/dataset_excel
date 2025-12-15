from __future__ import annotations

import os
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from .format_processor import ProcessOutputs
from .format_spec import DatasetFormatSpec, spec_to_yaml_dict, with_clearml_values


def _link_or_copy(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    try:
        os.link(src, dst)
    except OSError:
        shutil.copy2(src, dst)


def _ensure_empty_dir(path: Path, *, overwrite: bool) -> None:
    if path.exists():
        try:
            has_children = next(path.iterdir(), None) is not None
        except Exception:
            has_children = True
        if has_children:
            if not overwrite:
                raise FileExistsError(f"Stage dir already exists and is not empty: {path}")
            shutil.rmtree(path)
            path.mkdir(parents=True, exist_ok=True)
        return
    path.mkdir(parents=True, exist_ok=True)


def _find_requirements_txt() -> Path | None:
    candidates = [
        (Path.cwd() / "requirements.txt").resolve(),
        # When running from a source checkout: <repo>/src/clearml_dataset_excel/format_clearml.py -> <repo>/requirements.txt
        (Path(__file__).resolve().parents[2] / "requirements.txt").resolve(),
    ]
    for p in candidates:
        try:
            if p.exists() and p.is_file():
                return p
        except Exception:
            continue
    return None


def _ensure_output_uri_ready(output_uri: str | None) -> None:
    """
    Best-effort: ensure file:// output_uri points to an existing directory.
    ClearML's local file storage requires the target path to exist.
    """
    if not output_uri:
        return
    if not isinstance(output_uri, str):
        return
    if not output_uri.startswith("file://"):
        return

    # output_uri is expected to be like: file:///abs/path (posix) or file:///C:/path (windows)
    uri_path = output_uri[len("file://") :]
    if not uri_path:
        return
    # Handle file://localhost/...
    if uri_path.startswith("localhost/"):
        uri_path = "/" + uri_path[len("localhost/") :]

    try:
        p = Path(uri_path).expanduser()
        if not p.is_absolute():
            # Safety: avoid creating relative paths for file:// URIs.
            return
        p.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        raise RuntimeError(f"Failed to prepare output_uri directory: {output_uri} ({e})") from e


def _is_parent_not_finalized_error(e: Exception) -> bool:
    msg = str(e).lower()
    return "not finalized" in msg or "finalized/closed" in msg


def stage_template_payload_to_dir(
    *,
    stage_dir: Path,
    spec_path: Path,
    spec: DatasetFormatSpec,
    template_excel: Path,
    template_excel_windows: Path | None = None,
    addin_xlam_windows: Path | None = None,
    vba_module: Path | None = None,
    runner_exe_windows: Path | None = None,
    overwrite: bool = False,
) -> Path:
    stage = Path(stage_dir).expanduser().resolve()
    _ensure_empty_dir(stage, overwrite=overwrite)

    _link_or_copy(spec_path, stage / "spec" / spec_path.name)
    _link_or_copy(template_excel, stage / "template" / template_excel.name)
    if template_excel_windows is not None:
        _link_or_copy(template_excel_windows, stage / "template" / template_excel_windows.name)
    if addin_xlam_windows is not None:
        _link_or_copy(addin_xlam_windows, stage / "template" / addin_xlam_windows.name)
    if vba_module is not None:
        _link_or_copy(vba_module, stage / "template" / vba_module.name)
    if runner_exe_windows is not None:
        _link_or_copy(runner_exe_windows, stage / "template" / runner_exe_windows.name)
    if spec.addin.enabled:
        spec_name = spec.addin.spec_filename or spec_path.name
        _link_or_copy(spec_path, stage / "template" / spec_name)

    meta = {
        "payload_version": 1,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "spec_path": f"spec/{spec_path.name}",
        "template_excel": f"template/{template_excel.name}",
        "template_excel_windows": f"template/{template_excel_windows.name}" if template_excel_windows is not None else None,
        "addin_xlam_windows": f"template/{addin_xlam_windows.name}" if addin_xlam_windows is not None else None,
        "vba_module": f"template/{vba_module.name}" if vba_module is not None else None,
        "runner_exe_windows": f"template/{runner_exe_windows.name}" if runner_exe_windows is not None else None,
        "template_spec": f"template/{(spec.addin.spec_filename or spec_path.name)}" if spec.addin.enabled else None,
    }
    (stage / "payload.json").write_text(_json_dumps(meta) + "\n", encoding="utf-8")
    return stage


def stage_dataset_payload_to_dir(
    *,
    stage_dir: Path,
    spec_path: Path,
    spec: DatasetFormatSpec,
    condition_excel: Path,
    outputs: ProcessOutputs,
    template_excel: Path,
    template_excel_windows: Path | None = None,
    addin_xlam_windows: Path | None = None,
    vba_module: Path | None = None,
    runner_exe_windows: Path | None = None,
    overwrite: bool = False,
) -> Path:
    stage = Path(stage_dir).expanduser().resolve()
    _ensure_empty_dir(stage, overwrite=overwrite)

    # Spec + template
    _link_or_copy(spec_path, stage / "spec" / spec_path.name)
    _link_or_copy(template_excel, stage / "template" / template_excel.name)
    if template_excel_windows is not None:
        _link_or_copy(template_excel_windows, stage / "template" / template_excel_windows.name)
    if addin_xlam_windows is not None:
        _link_or_copy(addin_xlam_windows, stage / "template" / addin_xlam_windows.name)
    if vba_module is not None:
        _link_or_copy(vba_module, stage / "template" / vba_module.name)
    if runner_exe_windows is not None:
        _link_or_copy(runner_exe_windows, stage / "template" / runner_exe_windows.name)
    if spec.addin.enabled:
        # Place spec next to the template so the VBA module can resolve it via ThisWorkbook.Path
        spec_name = spec.addin.spec_filename or spec_path.name
        _link_or_copy(spec_path, stage / "template" / spec_name)

    # Inputs
    base_dir = condition_excel.parent.resolve()
    input_dir = stage / "input"
    seen_src: set[str] = set()
    src_abs_to_dest_rel: dict[str, str] = {}
    for i, p in enumerate(outputs.uploaded_files):
        p = p.resolve()
        p_key = p.as_posix()
        if p_key in seen_src:
            continue
        seen_src.add(p_key)
        try:
            rel = p.relative_to(base_dir)
            dest = input_dir / rel
        except ValueError:
            dest = input_dir / "external" / f"{i:03d}_{p.name}"
        if dest.exists():
            try:
                if dest.samefile(p):
                    continue
            except Exception:
                pass
            raise RuntimeError(f"Staging collision: {dest} <- {p}")
        _link_or_copy(p, dest)
        src_abs_to_dest_rel[p_key] = dest.relative_to(stage).as_posix()

    # Processed outputs
    processed_dir = stage / "processed"
    _link_or_copy(outputs.conditions_csv, processed_dir / outputs.conditions_csv.name)
    _link_or_copy(outputs.canonical_csv, processed_dir / outputs.canonical_csv.name)
    _link_or_copy(outputs.consolidated_excel, processed_dir / outputs.consolidated_excel.name)

    # Convenience metadata
    raw_path_map: dict[str, str] = {}
    for raw, resolved in outputs.raw_path_map.items():
        resolved_key = resolved.resolve().as_posix()
        dest_rel = src_abs_to_dest_rel.get(resolved_key)
        if dest_rel is not None:
            raw_path_map[str(raw)] = dest_rel

    excel_key = condition_excel.resolve().as_posix()
    excel_rel = src_abs_to_dest_rel.get(excel_key) or f"input/{condition_excel.name}"

    meta = {
        "payload_version": 1,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "spec_path": f"spec/{spec_path.name}",
        "template_excel": f"template/{template_excel.name}",
        "template_excel_windows": f"template/{template_excel_windows.name}" if template_excel_windows is not None else None,
        "addin_xlam_windows": f"template/{addin_xlam_windows.name}" if addin_xlam_windows is not None else None,
        "vba_module": f"template/{vba_module.name}" if vba_module is not None else None,
        "runner_exe_windows": f"template/{runner_exe_windows.name}" if runner_exe_windows is not None else None,
        "template_spec": f"template/{(spec.addin.spec_filename or spec_path.name)}" if spec.addin.enabled else None,
        "condition_excel": excel_rel,
        "path_map": raw_path_map,
        "conditions_csv": f"processed/{outputs.conditions_csv.name}",
        "canonical_csv": f"processed/{outputs.canonical_csv.name}",
        "consolidated_excel": f"processed/{outputs.consolidated_excel.name}",
    }
    (stage / "payload.json").write_text(_json_dumps(meta) + "\n", encoding="utf-8")
    return stage


def stage_template_payload(
    *,
    spec_path: Path,
    spec: DatasetFormatSpec,
    template_excel: Path,
    template_excel_windows: Path | None = None,
    addin_xlam_windows: Path | None = None,
    vba_module: Path | None = None,
    runner_exe_windows: Path | None = None,
) -> tempfile.TemporaryDirectory:
    """
    Create a staging directory containing spec/template (for distribution).

    Layout:
      - spec/<spec.yaml>
      - template/<template file>
      - template/<spec copy for VBA, optional>
      - template/<vba module (.bas), optional>
      - template/<clearml_dataset_excel_runner.exe, optional>
    """
    td = tempfile.TemporaryDirectory(prefix="clearml_dataset_excel_template_stage_")
    stage_template_payload_to_dir(
        stage_dir=Path(td.name).resolve(),
        spec_path=spec_path,
        spec=spec,
        template_excel=template_excel,
        template_excel_windows=template_excel_windows,
        addin_xlam_windows=addin_xlam_windows,
        vba_module=vba_module,
        runner_exe_windows=runner_exe_windows,
        overwrite=False,
    )
    return td


def stage_dataset_payload(
    *,
    spec_path: Path,
    spec: DatasetFormatSpec,
    condition_excel: Path,
    outputs: ProcessOutputs,
    template_excel: Path,
    template_excel_windows: Path | None = None,
    addin_xlam_windows: Path | None = None,
    vba_module: Path | None = None,
    runner_exe_windows: Path | None = None,
) -> tempfile.TemporaryDirectory:
    """
    Create a staging directory containing all inputs/outputs to be uploaded to ClearML Dataset.

    Layout:
      - spec/spec.yaml
      - template/<template file>
      - template/<vba module (.bas), optional>
      - template/<clearml_dataset_excel_runner.exe, optional>
      - input/<condition excel + referenced files (relative to condition excel dir if possible)>
      - processed/<outputs>
    """
    td = tempfile.TemporaryDirectory(prefix="clearml_dataset_excel_stage_")
    stage_dataset_payload_to_dir(
        stage_dir=Path(td.name).resolve(),
        spec_path=spec_path,
        spec=spec,
        condition_excel=condition_excel,
        outputs=outputs,
        template_excel=template_excel,
        template_excel_windows=template_excel_windows,
        addin_xlam_windows=addin_xlam_windows,
        vba_module=vba_module,
        runner_exe_windows=runner_exe_windows,
        overwrite=False,
    )
    return td


def _json_dumps(obj: Any) -> str:
    import json

    return json.dumps(obj, ensure_ascii=False, indent=2, sort_keys=True)


def upload_dataset(
    *,
    stage_dir: Path,
    spec: DatasetFormatSpec,
    dataset_project: str,
    dataset_name: str,
    base_dataset_id: str | None = None,
    tags: list[str] | None = None,
    output_uri: str | None = None,
    description: str | None = None,
    max_workers: int | None = None,
    verbose: bool = False,
) -> str:
    """
    Create (or create a writable copy of) a ClearML Dataset, upload staged files, and finalize.
    Returns dataset id.
    """
    try:
        from clearml import Dataset
    except Exception as e:  # pragma: no cover
        raise RuntimeError("clearml is required. Install dependencies first.") from e

    resolved_output_uri = output_uri if output_uri is not None else (spec.clearml.output_uri if spec.clearml else None)
    resolved_tags = tags if tags is not None else (spec.clearml.tags if spec.clearml else None)
    _ensure_output_uri_ready(resolved_output_uri)
    spec_for_task = with_clearml_values(
        spec,
        dataset_project=dataset_project,
        dataset_name=dataset_name,
        output_uri=resolved_output_uri,
        tags=resolved_tags,
    )

    dataset = None
    if base_dataset_id:
        try:
            dataset = Dataset.get(dataset_id=base_dataset_id, writable_copy=True)
        except Exception as e:
            if _is_parent_not_finalized_error(e):
                try:
                    dataset = Dataset.get(dataset_id=base_dataset_id)
                except Exception:
                    dataset = None
            else:
                dataset = None

    # Get latest dataset and create a writable copy (new version) if exists
    if dataset is None:
        try:
            dataset = Dataset.get(dataset_project=dataset_project, dataset_name=dataset_name, writable_copy=True)
        except Exception as e:
            if _is_parent_not_finalized_error(e):
                try:
                    dataset = Dataset.get(dataset_project=dataset_project, dataset_name=dataset_name)
                except Exception:
                    dataset = None
            else:
                dataset = None

    # Create a new dataset if it doesn't exist (or if we failed to recover).
    if dataset is None:
        dataset = Dataset.create(
            dataset_name=dataset_name,
            dataset_project=dataset_project,
            dataset_tags=resolved_tags or None,
            output_uri=resolved_output_uri,
            description=description,
            use_current_task=bool(spec_for_task.clearml and spec_for_task.clearml.use_current_task),
        )

    # Attach configuration + artifacts to the Dataset's task (if accessible)
    task = getattr(dataset, "_task", None)

    # Best-effort: write dataset/task link into the visible template Info sheet
    try:
        from .format_excel import annotate_template_with_clearml_info

        template_paths: list[Path] = []
        try:
            import json

            meta = json.loads((stage_dir / "payload.json").read_text(encoding="utf-8"))
            if isinstance(meta, dict):
                for k in ("template_excel", "template_excel_windows"):
                    rel = meta.get(k)
                    if isinstance(rel, str) and rel.strip():
                        template_paths.append((stage_dir / rel).resolve())
        except Exception:
            pass
        if not template_paths:
            template_paths = [stage_dir / "template" / spec.template.template_filename]

        web_url = None
        if task is not None and hasattr(task, "get_output_log_web_page"):
            try:
                web_url = task.get_output_log_web_page()
            except Exception:
                web_url = None

        for template_path in template_paths:
            annotate_template_with_clearml_info(
                template_path,
                dataset_project=dataset_project,
                dataset_name=dataset_name,
                dataset_id=str(getattr(dataset, "id", "")) or None,
                clearml_web_url=web_url,
            )
    except Exception:
        pass

    if task is not None:
        try:
            task.connect_configuration(spec_to_yaml_dict(spec_for_task), name="dataset_format_spec")
        except Exception:
            pass

        # Also store the full spec YAML into Hyperparameters for easy editing in clone/enqueue workflows.
        try:
            if hasattr(task, "connect"):
                import yaml

                spec_yaml = yaml.safe_dump(
                    spec_to_yaml_dict(spec_for_task),
                    allow_unicode=True,
                    sort_keys=False,
                )
                task.connect({"yaml": spec_yaml}, name="dataset_format_spec")
        except Exception:
            pass

        # Store dataset identifiers into Hyperparameters so cloned tasks can still locate the source dataset.
        try:
            if hasattr(task, "connect"):
                task.connect(
                    {
                        "dataset_id": str(getattr(dataset, "id", "")) or "",
                        "dataset_project": str(dataset_project) if dataset_project is not None else "",
                        "dataset_name": str(dataset_name) if dataset_name is not None else "",
                        "base_dataset_id": "" if base_dataset_id is None else str(base_dataset_id),
                        # Defaults for clone/enqueue override (editable in UI)
                        "output_dataset_project": str(dataset_project) if dataset_project is not None else "",
                        "output_dataset_name": str(dataset_name) if dataset_name is not None else "",
                        "output_uri": "" if resolved_output_uri is None else str(resolved_output_uri),
                        "output_tags": list(resolved_tags or []),
                    },
                    name="clearml_dataset_excel",
                )
        except Exception:
            pass

        # Make the dataset task runnable via clearml-agent when spec.clearml.execution is provided.
        # This enables clone/enqueue workflows for `clearml-dataset-excel agent reprocess`.
        try:
            if spec.clearml is not None and spec.clearml.execution is not None:
                entry_point = spec_for_task.clearml.execution.entry_point or "clearml_agent_reprocess.py"
                task.set_script(
                    repository=spec_for_task.clearml.execution.repository,
                    branch=spec_for_task.clearml.execution.branch,
                    commit=spec_for_task.clearml.execution.commit,
                    working_dir=spec_for_task.clearml.execution.working_dir,
                    entry_point=entry_point,
                )
                # Embed requirements into the Task by pointing to a local requirements.txt (if available).
                # ClearML will read the file at this point and store packages; the path is not used remotely.
                if hasattr(task, "set_packages"):
                    req = _find_requirements_txt()
                    if req is not None:
                        task.set_packages(req.as_posix())
        except Exception:
            pass

        # Upload key artifacts for easy download (spec/template/consolidated)
        try:
            template_file = stage_dir / "template"
            spec_file = stage_dir / "spec"
            processed_file = stage_dir / "processed"
            for p in _iter_files(spec_file):
                task.upload_artifact(name=p.name, artifact_object=p.as_posix(), wait_on_upload=True)
            for p in _iter_files(template_file):
                task.upload_artifact(name=p.name, artifact_object=p.as_posix(), wait_on_upload=True)
            for p in _iter_files(processed_file):
                task.upload_artifact(name=p.name, artifact_object=p.as_posix(), wait_on_upload=True)
        except Exception:
            pass

    # Upload staged payload into dataset
    dataset.add_files(
        path=stage_dir.as_posix(),
        local_base_folder=stage_dir.as_posix(),
        dataset_path=".",
        recursive=True,
        verbose=verbose,
        max_workers=max_workers,
    )

    # Report coverage/stats and a few debug samples
    try:
        _report_dataset_stats(dataset=dataset, stage_dir=stage_dir, spec=spec_for_task)
    except Exception:
        pass

    dataset.upload(show_progress=verbose, verbose=verbose, output_url=resolved_output_uri, max_workers=max_workers)
    ok = dataset.finalize(verbose=verbose, auto_upload=False)
    if not ok:
        raise RuntimeError("Dataset finalize failed")
    return str(dataset.id)


def _iter_files(folder: Path) -> Iterable[Path]:
    if not folder.exists():
        return []
    return sorted([p for p in folder.rglob("*") if p.is_file()])


def _report_dataset_stats(*, dataset: Any, stage_dir: Path, spec: DatasetFormatSpec) -> None:
    import pandas as pd

    logger = dataset.get_logger()

    processed_dir = stage_dir / "processed"
    cond_csv = processed_dir / spec.output.conditions_filename
    canon_csv = processed_dir / spec.output.canonical_filename

    if cond_csv.exists():
        cond_df = pd.read_csv(cond_csv)
        logger.report_scalar("Counts", "conditions_rows", float(len(cond_df)), 0)
        logger.report_scalar("Counts", "conditions_columns", float(len(cond_df.columns)), 0)
        cov = _missing_table(cond_df)
        logger.report_table("Coverage", "conditions", iteration=0, table_plot=cov)
        _report_numeric_stats(logger, title="Summary", series_prefix="conditions", df=cond_df)

        # File-path column coverage (fill rate per file spec)
        try:
            rows: list[dict[str, object]] = []
            for f in spec.files:
                col = f.path_column
                if col not in cond_df.columns:
                    continue
                s0 = cond_df[col]
                try:
                    s = s0.astype(str)
                    filled = (~s0.isna()) & (s.str.strip() != "")
                    unique_paths = int(s[filled].nunique(dropna=True))
                except Exception:
                    filled = ~s0.isna()
                    unique_paths = int(s0[filled].nunique(dropna=True))
                total = int(len(cond_df))
                filled_n = int(filled.sum())
                rows.append(
                    {
                        "file_id": f.file_id,
                        "path_column": col,
                        "total_rows": total,
                        "filled_rows": filled_n,
                        "filled_rate": (float(filled_n) / float(total)) if total else 0.0,
                        "unique_paths": unique_paths,
                    }
                )
            if rows and hasattr(logger, "report_table"):
                out = pd.DataFrame(rows).sort_values(["file_id", "path_column"])
                logger.report_table("Summary", "file_path_column_coverage", iteration=0, table_plot=out)
        except Exception:
            pass
    else:
        cond_df = None

    if canon_csv.exists():
        canon_df = pd.read_csv(canon_csv)
        logger.report_scalar("Counts", "canonical_rows", float(len(canon_df)), 0)
        logger.report_scalar("Counts", "canonical_columns", float(len(canon_df.columns)), 0)
        cov = _missing_table(canon_df)
        logger.report_table("Coverage", "canonical", iteration=0, table_plot=cov)
        _report_numeric_stats(logger, title="Summary", series_prefix="canonical", df=canon_df)

        # Points per condition row (if available)
        try:
            if "__condition_row" in canon_df.columns:
                counts = canon_df.groupby("__condition_row").size()
                logger.report_scalar("Counts", "conditions_with_points", float(len(counts)), 0)
                if hasattr(logger, "report_histogram"):
                    logger.report_histogram(
                        "Summary",
                        "canonical/points_per_condition_row",
                        values=[int(x) for x in counts.tolist()],
                        iteration=0,
                    )
                if hasattr(logger, "report_table") and len(counts) > 0:
                    top = (
                        counts.sort_values(ascending=False)
                        .head(200)
                        .reset_index()
                        .rename(columns={0: "points"})
                    )
                    logger.report_table("Summary", "canonical_points_per_condition_top", iteration=0, table_plot=top)
        except Exception:
            pass
    else:
        canon_df = None

    # Debug images (if any)
    input_dir = stage_dir / "input"
    if input_dir.exists():
        # File extension summary (quick sanity check)
        try:
            exts: dict[str, int] = {}
            for p in input_dir.rglob("*"):
                if not p.is_file():
                    continue
                ext = p.suffix.lower() or "<none>"
                exts[ext] = int(exts.get(ext, 0) + 1)
            if exts and hasattr(logger, "report_table"):
                rows = [{"ext": k, "count": int(v)} for k, v in sorted(exts.items(), key=lambda kv: (-kv[1], kv[0]))]
                logger.report_table("Summary", "input_file_ext_counts", iteration=0, table_plot=pd.DataFrame(rows))
        except Exception:
            pass

        exts = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}
        imgs = [p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in exts]
        for i, p in enumerate(sorted(imgs)[:20]):
            logger.report_image("Debug Samples", "images", iteration=i, local_path=p.as_posix(), delete_after_upload=False)

        # Debug tabular samples (CSV/TSV head)
        try:
            tab_exts = {".csv", ".tsv"}
            tabs = [p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in tab_exts]
            for i, p in enumerate(sorted(tabs)[:10]):
                try:
                    sep = "\t" if p.suffix.lower() == ".tsv" else ","
                    df = pd.read_csv(p, sep=sep, nrows=50)
                    logger.report_table("Debug Samples", f"tabular_head/{p.name}", iteration=i, table_plot=df)
                except Exception:
                    continue
        except Exception:
            pass


def _missing_table(df):  # type: ignore[no-untyped-def]
    import pandas as pd

    total = int(len(df))
    rows: list[dict[str, object]] = []
    for c in df.columns:
        s = df[c]
        missing = s.isna()
        # Treat empty strings as missing (after stripping)
        try:
            missing = missing | (s.astype(str).str.strip() == "")
        except Exception:
            pass
        m = int(missing.sum())
        rows.append(
            {
                "column": str(c),
                "total": total,
                "missing": m,
                "filled": int(total - m),
                "missing_rate": (float(m) / float(total)) if total else 0.0,
            }
        )
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(["missing_rate", "missing", "column"], ascending=[False, False, True])
    return out


def _report_numeric_stats(logger: Any, *, title: str, series_prefix: str, df):  # type: ignore[no-untyped-def]
    import pandas as pd

    rows: list[dict[str, object]] = []
    for c in df.columns:
        s0 = df[c]
        s = pd.to_numeric(s0, errors="coerce")
        n = int(s.notna().sum())
        if n == 0:
            continue
        rows.append(
            {
                "column": str(c),
                "count": n,
                "mean": float(s.mean()),
                "std": float(s.std(ddof=0)),
                "min": float(s.min()),
                "p50": float(s.quantile(0.5)),
                "max": float(s.max()),
            }
        )

        if hasattr(logger, "report_histogram"):
            vals = s.dropna()
            if len(vals) > 20000:
                vals = vals.sample(20000, random_state=0)
            if len(vals) > 0:
                logger.report_histogram(title, f"{series_prefix}/{c}", values=vals.tolist(), iteration=0)

    if rows and hasattr(logger, "report_table"):
        out = pd.DataFrame(rows).sort_values(["column"])
        logger.report_table(title, f"{series_prefix}_numeric", iteration=0, table_plot=out)
