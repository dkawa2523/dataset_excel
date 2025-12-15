from __future__ import annotations

import json
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

from .format_clearml import stage_dataset_payload, stage_dataset_payload_to_dir, upload_dataset
from .format_excel import generate_condition_template, generate_condition_template_from_excel, generate_windows_addin_xlam
from .format_processor import ProcessingError, process_condition_excel
from .format_spec import (
    DatasetFormatSpec,
    SpecError,
    load_format_spec,
    load_format_spec_from_mapping,
    with_clearml_values,
    write_spec_yaml,
)
from .vba_addin import write_vba_module
from .vba_embedder import embed_vba_module_into_xlsm


class AgentError(RuntimeError):
    pass


@dataclass(frozen=True)
class AgentStageResult:
    output_dir: Path
    stage_dir: Path


def _read_payload_meta(dataset_root: Path) -> dict[str, Any]:
    payload = dataset_root / "payload.json"
    if not payload.exists():
        return {}
    try:
        meta = json.loads(payload.read_text(encoding="utf-8"))
        return meta if isinstance(meta, dict) else {}
    except Exception:
        return {}


def _get_task_configuration_dict(task: Any, name: str) -> Mapping[str, Any]:
    if hasattr(task, "get_configuration_object_as_dict"):
        try:
            cfg = task.get_configuration_object_as_dict(name)
            if isinstance(cfg, dict):
                return cfg
        except Exception:
            pass

    if hasattr(task, "get_configuration_object"):
        try:
            cfg = task.get_configuration_object(name)
            if isinstance(cfg, dict):
                return cfg
            if isinstance(cfg, str) and cfg.strip():
                # Best-effort parsing (ClearML UI sometimes stores YAML as text)
                try:
                    return json.loads(cfg)
                except Exception:
                    try:
                        import yaml
                    except Exception as e:  # pragma: no cover
                        raise RuntimeError("Parsing spec requires 'PyYAML'. Install dependencies first.") from e
                    parsed = yaml.safe_load(cfg)
                    if isinstance(parsed, dict):
                        return parsed
        except Exception:
            pass

    raise AgentError(f"Task configuration '{name}' not found or not a mapping")


def _get_task_hyperparam_yaml(task: Any, section_name: str) -> str | None:
    key = f"{section_name}/yaml"
    if hasattr(task, "get_parameter"):
        try:
            v = task.get_parameter(key)
            if isinstance(v, str) and v.strip():
                return v
        except Exception:
            pass

    if hasattr(task, "get_parameters_as_dict"):
        try:
            params = task.get_parameters_as_dict(cast=False)
            section = params.get(section_name)
            if isinstance(section, dict):
                v = section.get("yaml")
                if isinstance(v, str) and v.strip():
                    return v
        except Exception:
            pass

    return None


def _get_task_param(task: Any, key: str) -> Any:
    if hasattr(task, "get_parameter"):
        try:
            return task.get_parameter(key)
        except Exception:
            return None
    if hasattr(task, "get_parameters_as_dict"):
        try:
            params = task.get_parameters_as_dict(cast=False)
            if not isinstance(params, dict):
                return None
            if "/" not in key:
                return params.get(key)
            section, name = key.split("/", 1)
            sec = params.get(section)
            if isinstance(sec, dict):
                return sec.get(name)
        except Exception:
            return None
    return None


def _get_task_param_str(task: Any, key: str) -> str | None:
    v = _get_task_param(task, key)
    if isinstance(v, str) and v.strip():
        return v.strip()
    return None


def _get_task_param_str_list(task: Any, key: str) -> list[str] | None:
    v = _get_task_param(task, key)
    if v is None:
        return None
    if isinstance(v, list):
        out: list[str] = []
        for x in v:
            if isinstance(x, str) and x.strip():
                out.append(x.strip())
        return out
    if isinstance(v, str) and v.strip():
        t = v.strip()
        try:
            parsed = json.loads(t)
            if isinstance(parsed, list):
                out = [str(x).strip() for x in parsed if str(x).strip()]
                return out
        except Exception:
            pass
        # Comma-separated fallback
        return [p.strip() for p in t.split(",") if p.strip()]
    return None


def _infer_dataset_id(task: Any) -> str | None:
    # Dataset tasks commonly have the dataset id equal to the task id.
    ds_id = getattr(task, "id", None)
    if isinstance(ds_id, str) and ds_id.strip():
        return ds_id.strip()
    return None


def _locate_condition_excel(dataset_root: Path, *, override: str | None = None) -> Path:
    if override:
        p = Path(override).expanduser()
        if not p.is_absolute():
            p = dataset_root / p
        return p.resolve()

    meta = _read_payload_meta(dataset_root)
    rel = meta.get("condition_excel")
    if isinstance(rel, str) and rel.strip():
        return (dataset_root / rel).resolve()

    input_dir = dataset_root / "input"
    if input_dir.exists():
        for ext in (".xlsm", ".xlsx", ".csv", ".tsv"):
            candidates = sorted([p for p in input_dir.glob(f"*{ext}") if p.is_file()])
            if candidates:
                return candidates[0].resolve()

    raise AgentError("Failed to locate condition Excel (payload.json missing and no input/*.xlsm|xlsx|csv|tsv found)")


def _locate_spec_yaml(dataset_root: Path, meta: Mapping[str, Any]) -> Path:
    rel = meta.get("spec_path")
    if isinstance(rel, str) and rel.strip():
        p = (dataset_root / rel).resolve()
        if p.exists():
            return p

    rel = meta.get("template_spec")
    if isinstance(rel, str) and rel.strip():
        p = (dataset_root / rel).resolve()
        if p.exists():
            return p

    spec_dir = dataset_root / "spec"
    if spec_dir.exists():
        candidates = sorted(
            [p for p in spec_dir.iterdir() if p.is_file() and p.suffix.lower() in {".yaml", ".yml"}]
        )
        if candidates:
            return candidates[0].resolve()

    raise AgentError("Spec YAML not found in dataset payload (expected payload.json spec_path or spec/*.yaml)")


def _locate_template_excel(dataset_root: Path, meta: Mapping[str, Any], *, prefer_name: str | None = None) -> Path | None:
    rel = meta.get("template_excel")
    if isinstance(rel, str) and rel.strip():
        p = (dataset_root / rel).resolve()
        if p.exists():
            return p

    template_dir = dataset_root / "template"
    if template_dir.exists():
        candidates = sorted(
            [p for p in template_dir.iterdir() if p.is_file() and p.suffix.lower() in {".xlsm", ".xlsx"}]
        )
        if prefer_name:
            for p in candidates:
                if p.name == prefer_name:
                    return p.resolve()
        if candidates:
            return candidates[0].resolve()
    return None


def _locate_runner_exe_windows(dataset_root: Path, meta: Mapping[str, Any]) -> Path | None:
    rel = meta.get("runner_exe_windows")
    if isinstance(rel, str) and rel.strip():
        p = (dataset_root / rel).resolve()
        if p.exists():
            return p

    template_dir = dataset_root / "template"
    if template_dir.exists():
        preferred = (template_dir / "clearml_dataset_excel_runner.exe").resolve()
        if preferred.exists() and preferred.is_file():
            return preferred

        candidates = sorted([p for p in template_dir.glob("*.exe") if p.is_file()])
        if candidates:
            return candidates[0].resolve()

    return None


def _write_spec_yaml(spec: DatasetFormatSpec, *, output_dir: Path) -> Path:
    name = Path(spec.addin.spec_filename or "spec.yaml").name
    return write_spec_yaml(spec, output_dir / name)


@dataclass(frozen=True)
class _ReprocessInputs:
    base_dataset: Any
    dataset_root: Path
    meta: dict[str, Any]
    spec: DatasetFormatSpec
    excel_path: Path
    fallback_map: dict[str, str] | None


def _prepare_reprocess_inputs(
    *,
    task: Any,
    dataset_id: str,
    config_name: str,
    condition_excel: str | None,
) -> _ReprocessInputs:
    try:
        from clearml import Dataset
    except Exception as e:  # pragma: no cover
        raise RuntimeError("clearml is required. Install dependencies first.") from e

    base_dataset = Dataset.get(dataset_id=dataset_id)
    dataset_root = Path(base_dataset.get_local_copy()).expanduser().resolve()
    meta = _read_payload_meta(dataset_root)

    pv = meta.get("payload_version", 0)
    try:
        pv_int = int(pv)
    except Exception:
        pv_int = 0
    if pv_int > 1:
        raise AgentError(f"Unsupported payload_version in payload.json: {pv_int}")

    raw_config: Mapping[str, Any] | None = None
    hp_yaml = _get_task_hyperparam_yaml(task, config_name)
    if isinstance(hp_yaml, str) and hp_yaml.strip():
        try:
            import yaml

            parsed = yaml.safe_load(hp_yaml)
            if isinstance(parsed, dict):
                raw_config = parsed
        except Exception:
            raw_config = None

    if raw_config is None:
        try:
            raw_config = _get_task_configuration_dict(task, config_name)
        except AgentError:
            raw_config = None

    try:
        if raw_config:
            spec = load_format_spec_from_mapping(raw_config)
        else:
            spec = load_format_spec(_locate_spec_yaml(dataset_root, meta))
    except SpecError as e:
        raise AgentError(str(e)) from e

    excel_path = _locate_condition_excel(dataset_root, override=condition_excel)
    if not excel_path.exists():
        raise AgentError(f"Condition Excel not found in dataset: {excel_path}")

    fallback_map: dict[str, str] | None = None
    path_map = meta.get("path_map")
    if isinstance(path_map, dict) and path_map:
        fallback_map = {}
        for k, v in path_map.items():
            if not isinstance(k, str) or not isinstance(v, str):
                continue
            p = (dataset_root / v).resolve()
            fallback_map[k] = p.as_posix()

    return _ReprocessInputs(
        base_dataset=base_dataset,
        dataset_root=dataset_root,
        meta=meta,
        spec=spec,
        excel_path=excel_path,
        fallback_map=fallback_map,
    )


def stage_reprocess_dataset_from_task(
    *,
    task: Any,
    dataset_id: str,
    config_name: str = "dataset_format_spec",
    condition_excel: str | None = None,
    sheet_name: str | None = None,
    output_root: str | Path | None = None,
    stage_dir: str | Path | None = None,
    overwrite_stage: bool = False,
) -> AgentStageResult:
    inputs = _prepare_reprocess_inputs(
        task=task,
        dataset_id=dataset_id,
        config_name=config_name,
        condition_excel=condition_excel,
    )

    output_root_path = Path(output_root).expanduser().resolve() if output_root else Path.cwd().resolve()

    try:
        outputs = process_condition_excel(
            inputs.spec,
            inputs.excel_path,
            sheet_name=sheet_name,
            output_root=output_root_path,
            check_files_exist=True,
            path_fallback_map=inputs.fallback_map,
            fallback_search_root=inputs.excel_path.parent,
        )
    except ProcessingError as e:
        raise AgentError(str(e)) from e

    with tempfile.TemporaryDirectory(prefix="clearml_dataset_excel_agent_stage_") as td:
        td_path = Path(td).resolve()

        template_path = td_path / inputs.spec.template.template_filename
        base_template = _locate_template_excel(inputs.dataset_root, inputs.meta, prefer_name=inputs.spec.template.template_filename)
        if base_template is not None and base_template.suffix.lower() == ".xlsm":
            generate_condition_template_from_excel(base_template, inputs.spec, template_path, overwrite=True)
        else:
            generate_condition_template(inputs.spec, template_path, overwrite=True)

        if inputs.spec.addin.enabled and inputs.spec.addin.embed_vba and template_path.suffix.lower() == ".xlsm":
            try:
                embed_vba_module_into_xlsm(
                    excel_path=template_path,
                    bas_path=None,
                    overwrite=True,
                    template_excel=None,
                )
            except Exception:
                # Best-effort: keep the .bas next to the template as a manual fallback.
                pass

        win_template_path = None
        win_addin_path = None
        if inputs.spec.addin.enabled and inputs.spec.addin.windows_mode == "addin":
            try:
                if inputs.spec.addin.windows_template_filename:
                    win_template_path = td_path / inputs.spec.addin.windows_template_filename
                    generate_condition_template(inputs.spec, win_template_path, overwrite=True)
                win_addin_path = td_path / inputs.spec.addin.windows_addin_filename
                generate_windows_addin_xlam(win_addin_path, overwrite=True)
            except Exception:
                win_template_path = None
                win_addin_path = None

        vba_path = None
        if inputs.spec.addin.enabled:
            vba_path = td_path / inputs.spec.addin.vba_module_filename
            write_vba_module(vba_path, meta_sheet_name=inputs.spec.template.meta_sheet)

        spec_path = _write_spec_yaml(inputs.spec, output_dir=td_path)
        runner_exe_windows = _locate_runner_exe_windows(inputs.dataset_root, inputs.meta) if inputs.spec.addin.enabled else None

        stage_dir_path = Path(stage_dir).expanduser().resolve() if stage_dir else outputs.output_dir / "_clearml_stage"
        stage_dataset_payload_to_dir(
            stage_dir=stage_dir_path,
            spec_path=spec_path,
            spec=inputs.spec,
            condition_excel=inputs.excel_path,
            outputs=outputs,
            template_excel=template_path,
            template_excel_windows=win_template_path,
            addin_xlam_windows=win_addin_path,
            vba_module=vba_path,
            runner_exe_windows=runner_exe_windows,
            overwrite=overwrite_stage,
        )

    return AgentStageResult(output_dir=outputs.output_dir.resolve(), stage_dir=stage_dir_path.resolve())


def reprocess_dataset_from_task(
    *,
    task: Any,
    dataset_id: str,
    dataset_project: str | None = None,
    dataset_name: str | None = None,
    config_name: str = "dataset_format_spec",
    condition_excel: str | None = None,
    sheet_name: str | None = None,
    output_uri: str | None = None,
    tags: list[str] | None = None,
    description: str | None = None,
    max_workers: int | None = None,
    verbose: bool = False,
) -> str:
    inputs = _prepare_reprocess_inputs(
        task=task,
        dataset_id=dataset_id,
        config_name=config_name,
        condition_excel=condition_excel,
    )

    base_project = getattr(inputs.base_dataset, "project", None)
    base_name = getattr(inputs.base_dataset, "name", None)

    task_project = None
    if hasattr(task, "get_project_name"):
        try:
            task_project = task.get_project_name()
        except Exception:
            task_project = None
    if not task_project:
        task_project = getattr(task, "project", None)

    # Optional output overrides from task hyperparameters (useful in clone/enqueue without script args).
    hp_out_project = _get_task_param_str(task, "clearml_dataset_excel/output_dataset_project")
    hp_out_name = _get_task_param_str(task, "clearml_dataset_excel/output_dataset_name")
    hp_out_uri = _get_task_param_str(task, "clearml_dataset_excel/output_uri")
    hp_out_tags = _get_task_param_str_list(task, "clearml_dataset_excel/output_tags")

    resolved_project = (
        dataset_project
        or hp_out_project
        or (task_project if isinstance(task_project, str) and task_project else None)
        or (base_project if isinstance(base_project, str) and base_project else None)
    )
    if not resolved_project:
        resolved_project = inputs.spec.clearml.dataset_project if inputs.spec.clearml else None

    resolved_name = dataset_name or hp_out_name or (base_name if isinstance(base_name, str) and base_name else None)
    if not resolved_name:
        resolved_name = inputs.spec.clearml.dataset_name if inputs.spec.clearml else None

    resolved_output_uri = output_uri or hp_out_uri or (inputs.spec.clearml.output_uri if inputs.spec.clearml else None)
    if tags is not None:
        resolved_tags = tags
    elif hp_out_tags is not None:
        resolved_tags = hp_out_tags
    else:
        resolved_tags = inputs.spec.clearml.tags if inputs.spec.clearml else None
    if not resolved_project or not resolved_name:
        raise AgentError(
            "dataset_project/dataset_name is required (set in base dataset, spec.clearml, or CLI override)"
        )

    spec_for_upload = with_clearml_values(
        inputs.spec,
        dataset_project=resolved_project,
        dataset_name=resolved_name,
        output_uri=resolved_output_uri,
        tags=resolved_tags,
    )

    use_base_dataset_id = (
        isinstance(base_project, str)
        and isinstance(base_name, str)
        and base_project
        and base_name
        and resolved_project == base_project
        and resolved_name == base_name
    )

    with tempfile.TemporaryDirectory(prefix="clearml_dataset_excel_agent_") as td:
        td_path = Path(td).resolve()

        try:
            outputs = process_condition_excel(
                spec_for_upload,
                inputs.excel_path,
                sheet_name=sheet_name,
                output_root=td_path,
                check_files_exist=True,
                path_fallback_map=inputs.fallback_map,
                fallback_search_root=inputs.excel_path.parent,
            )
        except ProcessingError as e:
            raise AgentError(str(e)) from e

        template_path = td_path / spec_for_upload.template.template_filename
        base_template = _locate_template_excel(
            inputs.dataset_root,
            inputs.meta,
            prefer_name=spec_for_upload.template.template_filename,
        )
        if base_template is not None and base_template.suffix.lower() == ".xlsm":
            generate_condition_template_from_excel(base_template, spec_for_upload, template_path, overwrite=True)
        else:
            generate_condition_template(spec_for_upload, template_path, overwrite=True)

        if spec_for_upload.addin.enabled and spec_for_upload.addin.embed_vba and template_path.suffix.lower() == ".xlsm":
            try:
                embed_vba_module_into_xlsm(
                    excel_path=template_path,
                    bas_path=None,
                    overwrite=True,
                    template_excel=None,
                )
            except Exception:
                pass

        win_template_path = None
        win_addin_path = None
        if spec_for_upload.addin.enabled and spec_for_upload.addin.windows_mode == "addin":
            try:
                if spec_for_upload.addin.windows_template_filename:
                    win_template_path = td_path / spec_for_upload.addin.windows_template_filename
                    generate_condition_template(spec_for_upload, win_template_path, overwrite=True)
                win_addin_path = td_path / spec_for_upload.addin.windows_addin_filename
                generate_windows_addin_xlam(win_addin_path, overwrite=True)
            except Exception:
                win_template_path = None
                win_addin_path = None

        vba_path = None
        if spec_for_upload.addin.enabled:
            vba_path = td_path / spec_for_upload.addin.vba_module_filename
            write_vba_module(vba_path, meta_sheet_name=spec_for_upload.template.meta_sheet)

        spec_path = _write_spec_yaml(spec_for_upload, output_dir=td_path)
        runner_exe_windows = _locate_runner_exe_windows(inputs.dataset_root, inputs.meta) if spec_for_upload.addin.enabled else None

        stage_td = stage_dataset_payload(
            spec_path=spec_path,
            spec=spec_for_upload,
            condition_excel=inputs.excel_path,
            outputs=outputs,
            template_excel=template_path,
            template_excel_windows=win_template_path,
            addin_xlam_windows=win_addin_path,
            vba_module=vba_path,
            runner_exe_windows=runner_exe_windows,
        )
        try:
            stage_dir = Path(stage_td.name).resolve()
            ds_id = upload_dataset(
                stage_dir=stage_dir,
                spec=spec_for_upload,
                dataset_project=resolved_project,
                dataset_name=resolved_name,
                base_dataset_id=(dataset_id if use_base_dataset_id else None),
                tags=resolved_tags,
                output_uri=resolved_output_uri,
                description=description,
                max_workers=max_workers,
                verbose=verbose,
            )
        finally:
            stage_td.cleanup()

    return ds_id


def infer_or_raise_dataset_id(task: Any, dataset_id: str | None) -> str:
    if dataset_id:
        return dataset_id

    # Prefer an explicit hyperparameter stored by clearml-dataset-excel (works for cloned tasks)
    try:
        v = None
        if hasattr(task, "get_parameter"):
            try:
                v = task.get_parameter("clearml_dataset_excel/dataset_id")
            except Exception:
                v = None
        if not v and hasattr(task, "get_parameters_as_dict"):
            try:
                params = task.get_parameters_as_dict(cast=False)
                sec = params.get("clearml_dataset_excel")
                if isinstance(sec, dict):
                    v = sec.get("dataset_id")
            except Exception:
                v = None
        if isinstance(v, str) and v.strip():
            return v.strip()
    except Exception:
        pass

    inferred = _infer_dataset_id(task)
    if inferred:
        return inferred
    raise AgentError("--dataset-id is required (failed to infer from Task)")
