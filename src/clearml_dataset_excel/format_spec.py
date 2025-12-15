from __future__ import annotations

from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any, Mapping


class SpecError(ValueError):
    pass


def _as_bool(v: Any, *, path: str) -> bool:
    if isinstance(v, bool):
        return v
    if v in (0, 1):
        return bool(v)
    raise SpecError(f"{path}: expected bool, got {type(v).__name__}")


def _as_str(v: Any, *, path: str) -> str:
    if isinstance(v, str) and v.strip():
        return v
    raise SpecError(f"{path}: expected non-empty string, got {type(v).__name__}")


def _as_opt_str(v: Any, *, path: str) -> str | None:
    if v is None:
        return None
    if isinstance(v, str) and v.strip():
        return v
    raise SpecError(f"{path}: expected string|null, got {type(v).__name__}")


def _as_list(v: Any, *, path: str) -> list[Any]:
    if v is None:
        return []
    if isinstance(v, list):
        return v
    raise SpecError(f"{path}: expected list, got {type(v).__name__}")


def _as_dict(v: Any, *, path: str) -> dict[str, Any]:
    if v is None:
        return {}
    if isinstance(v, dict):
        return dict(v)
    raise SpecError(f"{path}: expected mapping, got {type(v).__name__}")


def _as_number(v: Any, *, path: str) -> float:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    raise SpecError(f"{path}: expected number, got {type(v).__name__}")


def _normalize_key(key: Any) -> str:
    return str(key).strip().replace("-", "_")


def _normalize_mapping(m: Mapping[str, Any]) -> dict[str, Any]:
    return {_normalize_key(k): v for k, v in m.items()}


def _parse_axis_source_and_type(v: Any, *, path: str) -> tuple[str | None, str | None]:
    """
    Accept either:
      - null
      - "<column_name>"
      - {source: "<column_name>", type: "<dtype>"} (type is optional; dtype is accepted as alias)
    """
    if v is None:
        return None, None
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None, None
        return s, None
    if isinstance(v, Mapping):
        m = _normalize_mapping(v)
        src = m.get("source", m.get("column", m.get("name")))
        source = _as_str(src, path=f"{path}.source")
        dtype_raw = m.get("type", m.get("dtype"))
        dtype = _as_opt_str(dtype_raw, path=f"{path}.type")
        return source, dtype
    raise SpecError(f"{path}: expected string|null|mapping, got {type(v).__name__}")


@dataclass(frozen=True)
class ConditionColumn:
    name: str
    dtype: str = "str"
    required: bool = False
    description: str | None = None
    enum: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class AxisMapping:
    x: str | None = None
    y: str | None = None
    z: str | None = None
    t: str | None = None

    def defined_axes(self) -> set[str]:
        axes: set[str] = set()
        for k in ("x", "y", "z", "t"):
            if getattr(self, k) is not None:
                axes.add(k)
        return axes


@dataclass(frozen=True)
class AxisTypeMapping:
    x: str | None = None
    y: str | None = None
    z: str | None = None
    t: str | None = None


@dataclass(frozen=True)
class TargetColumn:
    name: str
    source: str
    dtype: str = "float"


@dataclass(frozen=True)
class DerivedColumn:
    name: str
    expr: str
    dtype: str = "float"


@dataclass(frozen=True)
class Aggregate:
    name: str
    source: str
    op: str
    wrt: str | None = None
    output_column: str | None = None


@dataclass(frozen=True)
class FileSpec:
    file_id: str
    path_column: str
    format: str = "csv"
    sheet: str | None = None
    read: dict[str, Any] = field(default_factory=dict)
    axes: AxisMapping = field(default_factory=AxisMapping)
    axis_types: AxisTypeMapping = field(default_factory=AxisTypeMapping)
    targets: list[TargetColumn] = field(default_factory=list)
    derived: list[DerivedColumn] = field(default_factory=list)
    aggregates: list[Aggregate] = field(default_factory=list)

    def all_target_names(self) -> list[str]:
        return [t.name for t in self.targets] + [d.name for d in self.derived]


@dataclass(frozen=True)
class ClearMLSpec:
    dataset_project: str
    dataset_name: str
    output_uri: str | None = None
    tags: list[str] = field(default_factory=list)
    use_current_task: bool = False
    execution: "ExecutionSpec | None" = None


@dataclass(frozen=True)
class ExecutionSpec:
    repository: str
    branch: str | None = None
    commit: str | None = None
    working_dir: str | None = None
    entry_point: str | None = None


@dataclass(frozen=True)
class TemplateSpec:
    condition_sheet: str = "Conditions"
    meta_sheet: str = "_meta"
    template_filename: str = "condition_template.xlsm"


@dataclass(frozen=True)
class AddinSpec:
    enabled: bool = False
    target_os: str = "auto"  # auto|mac|windows
    spec_filename: str | None = None
    vba_module_filename: str = "clearml_dataset_excel_addin.bas"
    vba_template_excel: str | None = None
    embed_vba: bool = False
    command: str | None = None
    command_mac: str | None = None
    command_windows: str | None = None
    windows_mode: str = "macro"  # macro|addin
    windows_template_filename: str | None = None  # macro-free template for Windows add-in mode (.xlsx)
    windows_addin_filename: str = "clearml_dataset_excel_addin.xlam"


@dataclass(frozen=True)
class OutputSpec:
    output_dirname: str = "processed"
    canonical_filename: str = "canonical.csv"
    conditions_filename: str = "conditions.csv"
    consolidated_excel_filename: str = "consolidated.xlsx"
    include_file_path_columns: bool = True
    combine_mode: str = "auto"  # auto|merge|append


@dataclass(frozen=True)
class DatasetFormatSpec:
    schema_version: int
    clearml: ClearMLSpec | None
    template: TemplateSpec
    addin: AddinSpec
    condition_columns: list[ConditionColumn]
    files: list[FileSpec]
    output: OutputSpec

    def file_path_columns(self) -> set[str]:
        return {f.path_column for f in self.files}


def parse_format_spec(raw: Mapping[str, Any], *, spec_path: Path | None = None) -> DatasetFormatSpec:
    if not isinstance(raw, Mapping):
        raise SpecError("Spec root must be a mapping (dict)")
    raw = _normalize_mapping(raw)

    schema_version = raw.get("schema_version")
    if schema_version != 1:
        raise SpecError("schema_version must be 1")

    clearml_raw = _as_dict(raw.get("clearml"), path="clearml")
    clearml: ClearMLSpec | None
    if clearml_raw:
        clearml_raw = _normalize_mapping(clearml_raw)
        output_uri = _as_opt_str(clearml_raw.get("output_uri"), path="clearml.output_uri")
        if spec_path is not None and isinstance(output_uri, str) and output_uri.startswith("file://"):
            uri_path = output_uri[len("file://") :]
            if uri_path and not uri_path.startswith("/"):
                abs_path = (spec_path.parent / Path(uri_path)).expanduser().resolve()
                output_uri = "file://" + abs_path.as_posix()
        exec_raw = _normalize_mapping(_as_dict(clearml_raw.get("execution"), path="clearml.execution"))
        execution = None
        if exec_raw:
            repository = _as_str(exec_raw.get("repository"), path="clearml.execution.repository")
            execution = ExecutionSpec(
                repository=repository,
                branch=_as_opt_str(exec_raw.get("branch"), path="clearml.execution.branch"),
                commit=_as_opt_str(exec_raw.get("commit"), path="clearml.execution.commit"),
                working_dir=_as_opt_str(exec_raw.get("working_dir"), path="clearml.execution.working_dir"),
                entry_point=_as_opt_str(exec_raw.get("entry_point"), path="clearml.execution.entry_point"),
            )

        clearml = ClearMLSpec(
            dataset_project=_as_str(clearml_raw.get("dataset_project"), path="clearml.dataset_project"),
            dataset_name=_as_str(clearml_raw.get("dataset_name"), path="clearml.dataset_name"),
            output_uri=output_uri,
            tags=[str(x) for x in _as_list(clearml_raw.get("tags"), path="clearml.tags")],
            use_current_task=_as_bool(clearml_raw.get("use_current_task", False), path="clearml.use_current_task"),
            execution=execution,
        )
    else:
        clearml = None

    template_raw = _normalize_mapping(_as_dict(raw.get("template"), path="template"))
    template = TemplateSpec(
        condition_sheet=str(template_raw.get("condition_sheet", "Conditions")),
        meta_sheet=str(template_raw.get("meta_sheet", "_meta")),
        template_filename=str(template_raw.get("template_filename", "condition_template.xlsm")),
    )

    addin_raw = _normalize_mapping(_as_dict(raw.get("addin"), path="addin"))
    target_os = str(addin_raw.get("target_os", addin_raw.get("os", "auto"))).strip().lower() if addin_raw else "auto"
    if target_os in {"win"}:
        target_os = "windows"
    if target_os not in {"auto", "mac", "windows"}:
        raise SpecError("addin.target_os must be one of: auto, mac, windows")

    spec_filename = addin_raw.get("spec_filename") if addin_raw else None
    if spec_filename is not None:
        spec_filename = _as_opt_str(spec_filename, path="addin.spec_filename")
    if not spec_filename:
        spec_filename = spec_path.name if spec_path is not None else "spec.yaml"
    if Path(str(spec_filename)).name != str(spec_filename):
        raise SpecError("addin.spec_filename must be a file name (no directory components)")

    addin_enabled = _as_bool(addin_raw.get("enabled", False), path="addin.enabled") if addin_raw else False
    addin_command = _as_opt_str(addin_raw.get("command"), path="addin.command") if addin_raw else None
    addin_command_mac = _as_opt_str(addin_raw.get("command_mac"), path="addin.command_mac") if addin_raw else None
    addin_command_windows = _as_opt_str(addin_raw.get("command_windows"), path="addin.command_windows") if addin_raw else None

    # More robust defaults for end-user Excel execution: only when user did not specify any command.
    if addin_enabled and not addin_command and not addin_command_mac:
        addin_command_mac = '/usr/bin/env python3 -m clearml_dataset_excel.cli run --spec "${SPEC}" --excel "${EXCEL}"'

    addin = AddinSpec(
        enabled=addin_enabled,
        target_os=target_os,
        spec_filename=spec_filename,
        vba_module_filename=str(addin_raw.get("vba_module_filename", "clearml_dataset_excel_addin.bas"))
        if addin_raw
        else "clearml_dataset_excel_addin.bas",
        vba_template_excel=_as_opt_str(addin_raw.get("vba_template_excel"), path="addin.vba_template_excel")
        if addin_raw
        else None,
        embed_vba=_as_bool(addin_raw.get("embed_vba", False), path="addin.embed_vba") if addin_raw else False,
        command=addin_command,
        command_mac=addin_command_mac,
        command_windows=addin_command_windows,
        windows_mode="macro",
        windows_template_filename=None,
        windows_addin_filename="clearml_dataset_excel_addin.xlam",
    )

    # Windows distribution mode: workbook macro (legacy) vs .xlam add-in (recommended on Windows).
    win_mode = str(
        addin_raw.get("windows_mode", addin_raw.get("mode_windows", addin_raw.get("win_mode", "macro")))
        if addin_raw
        else "macro"
    ).strip().lower()
    if win_mode in {"workbook", "macro"}:
        win_mode = "macro"
    elif win_mode in {"addin", "add-in", "xlam"}:
        win_mode = "addin"
    else:
        raise SpecError("addin.windows_mode must be one of: macro, addin")

    win_template = (
        _as_opt_str(
            addin_raw.get("windows_template_filename", addin_raw.get("template_filename_windows")),
            path="addin.windows_template_filename",
        )
        if addin_raw
        else None
    )
    win_addin = (
        _as_opt_str(
            addin_raw.get("windows_addin_filename", addin_raw.get("addin_filename_windows")),
            path="addin.windows_addin_filename",
        )
        if addin_raw
        else None
    )
    if not win_addin:
        win_addin = "clearml_dataset_excel_addin.xlam"

    if win_mode == "addin":
        if not win_template:
            try:
                from pathlib import Path as _Path

                win_template = _Path(template.template_filename).with_suffix(".xlsx").name
            except Exception:
                win_template = "condition_template.xlsx"
        if not str(win_template).lower().endswith(".xlsx"):
            raise SpecError("addin.windows_template_filename must end with .xlsx when addin.windows_mode=addin")
        if not str(win_addin).lower().endswith(".xlam"):
            raise SpecError("addin.windows_addin_filename must end with .xlam when addin.windows_mode=addin")

    addin = replace(
        addin,
        windows_mode=win_mode,
        windows_template_filename=win_template,
        windows_addin_filename=win_addin,
    )

    if addin.enabled and addin.windows_mode == "addin" and not addin.command and not addin.command_windows:
        addin = replace(
            addin,
            command_windows=(
                'if exist "clearml_dataset_excel_runner.exe" '
                '("clearml_dataset_excel_runner.exe" run --spec "${SPEC}" --excel "${EXCEL}") '
                'else (clearml-dataset-excel run --spec "${SPEC}" --excel "${EXCEL}")'
            ),
        )

    # The bundled default vbaProject.bin (for embed_vba and .xlam generation) assumes meta sheet name "_meta".
    if addin.windows_mode == "addin" and template.meta_sheet != "_meta":
        raise SpecError("addin.windows_mode=addin requires template.meta_sheet to be '_meta' (bundled .xlam macro expects _meta).")
    if addin.embed_vba and not addin.vba_template_excel and template.meta_sheet != "_meta":
        raise SpecError("addin.embed_vba=true requires template.meta_sheet to be '_meta' when using bundled vbaProject.bin.")

    output_raw = _normalize_mapping(_as_dict(raw.get("output"), path="output"))
    combine_mode = str(output_raw.get("combine_mode", "auto")).strip().lower()
    if combine_mode not in {"auto", "merge", "append"}:
        raise SpecError("output.combine_mode must be one of: auto, merge, append")
    output = OutputSpec(
        output_dirname=str(output_raw.get("output_dirname", "processed")),
        canonical_filename=str(output_raw.get("canonical_filename", "canonical.csv")),
        conditions_filename=str(output_raw.get("conditions_filename", "conditions.csv")),
        consolidated_excel_filename=str(output_raw.get("consolidated_excel_filename", "consolidated.xlsx")),
        include_file_path_columns=_as_bool(
            output_raw.get("include_file_path_columns", True), path="output.include_file_path_columns"
        ),
        combine_mode=combine_mode,
    )

    cond_from_root = raw.get("condition")
    if cond_from_root is None and "condition_columns" in raw:
        cond_from_root = {"columns": raw.get("condition_columns")}
    cond_raw = _normalize_mapping(_as_dict(cond_from_root, path="condition"))
    columns_raw = _as_list(cond_raw.get("columns"), path="condition.columns")
    if not columns_raw:
        raise SpecError("condition.columns is required and must be non-empty")

    condition_columns: list[ConditionColumn] = []
    seen_col_names: set[str] = set()
    for i, c in enumerate(columns_raw):
        if not isinstance(c, dict):
            raise SpecError(f"condition.columns[{i}]: expected mapping")
        c = _normalize_mapping(c)
        name = _as_str(c.get("name"), path=f"condition.columns[{i}].name")
        if name in seen_col_names:
            raise SpecError(f"condition.columns[{i}].name: duplicate column name: {name}")
        seen_col_names.add(name)
        dtype = str(c.get("type", c.get("dtype", "str")))
        required = _as_bool(c.get("required", False), path=f"condition.columns[{i}].required")
        description = _as_opt_str(c.get("description"), path=f"condition.columns[{i}].description")
        enum = [str(x) for x in _as_list(c.get("enum"), path=f"condition.columns[{i}].enum")]
        condition_columns.append(
            ConditionColumn(name=name, dtype=dtype, required=required, description=description, enum=enum)
        )

    files_raw = _as_list(raw.get("files"), path="files")
    if not files_raw:
        raise SpecError("files is required and must be non-empty")

    files: list[FileSpec] = []
    seen_file_ids: set[str] = set()
    file_path_columns: set[str] = set()
    for i, f in enumerate(files_raw):
        if not isinstance(f, dict):
            raise SpecError(f"files[{i}]: expected mapping")
        f = _normalize_mapping(f)
        file_id = _as_str(f.get("id", f.get("file_id")), path=f"files[{i}].id")
        if file_id in seen_file_ids:
            raise SpecError(f"files[{i}].id: duplicate id: {file_id}")
        seen_file_ids.add(file_id)

        path_column = _as_str(f.get("path_column"), path=f"files[{i}].path_column")
        if path_column in file_path_columns:
            raise SpecError(f"files[{i}].path_column: duplicate path_column: {path_column}")
        file_path_columns.add(path_column)

        fmt = str(f.get("format", "csv")).lower()
        sheet = _as_opt_str(f.get("sheet"), path=f"files[{i}].sheet")
        read_opts = _as_dict(f.get("read"), path=f"files[{i}].read")

        mapping_src = f.get("mapping")
        if mapping_src is None and any(k in f for k in ("axes", "targets", "derived", "aggregates")):
            mapping_src = {
                "axes": f.get("axes"),
                "targets": f.get("targets"),
                "derived": f.get("derived"),
                "aggregates": f.get("aggregates"),
            }
        mapping = _normalize_mapping(_as_dict(mapping_src, path=f"files[{i}].mapping"))
        axes_raw = _normalize_mapping(_as_dict(mapping.get("axes"), path=f"files[{i}].mapping.axes"))
        axes_kwargs: dict[str, str | None] = {}
        axis_type_kwargs: dict[str, str | None] = {}
        for axis in ("x", "y", "z", "t"):
            src, dtype = _parse_axis_source_and_type(axes_raw.get(axis), path=f"files[{i}].mapping.axes.{axis}")
            axes_kwargs[axis] = src
            axis_type_kwargs[axis] = dtype
        axes = AxisMapping(**axes_kwargs)
        axis_types = AxisTypeMapping(**axis_type_kwargs)

        targets_raw = _as_list(mapping.get("targets"), path=f"files[{i}].mapping.targets")
        targets: list[TargetColumn] = []
        seen_targets: set[str] = set()
        for j, t in enumerate(targets_raw):
            if not isinstance(t, dict):
                raise SpecError(f"files[{i}].mapping.targets[{j}]: expected mapping")
            t = _normalize_mapping(t)
            t_name = _as_str(t.get("name"), path=f"files[{i}].mapping.targets[{j}].name")
            if t_name in seen_targets:
                raise SpecError(f"files[{i}].mapping.targets[{j}].name: duplicate target name: {t_name}")
            seen_targets.add(t_name)
            t_source = _as_str(t.get("source"), path=f"files[{i}].mapping.targets[{j}].source")
            t_dtype = str(t.get("type", t.get("dtype", "float")))
            targets.append(TargetColumn(name=t_name, source=t_source, dtype=t_dtype))

        derived_raw = _as_list(mapping.get("derived"), path=f"files[{i}].mapping.derived")
        derived: list[DerivedColumn] = []
        for j, d in enumerate(derived_raw):
            if not isinstance(d, dict):
                raise SpecError(f"files[{i}].mapping.derived[{j}]: expected mapping")
            d = _normalize_mapping(d)
            d_name = _as_str(d.get("name"), path=f"files[{i}].mapping.derived[{j}].name")
            if d_name in seen_targets:
                raise SpecError(f"files[{i}].mapping.derived[{j}].name: duplicate name: {d_name}")
            seen_targets.add(d_name)
            expr = _as_str(d.get("expr"), path=f"files[{i}].mapping.derived[{j}].expr")
            d_dtype = str(d.get("type", d.get("dtype", "float")))
            derived.append(DerivedColumn(name=d_name, expr=expr, dtype=d_dtype))

        aggregates_raw = _as_list(mapping.get("aggregates"), path=f"files[{i}].mapping.aggregates")
        aggregates: list[Aggregate] = []
        for j, a in enumerate(aggregates_raw):
            if not isinstance(a, dict):
                raise SpecError(f"files[{i}].mapping.aggregates[{j}]: expected mapping")
            a = _normalize_mapping(a)
            a_name = _as_str(a.get("name"), path=f"files[{i}].mapping.aggregates[{j}].name")
            a_source = _as_str(a.get("source"), path=f"files[{i}].mapping.aggregates[{j}].source")
            op = str(a.get("op")).strip()
            if not op:
                raise SpecError(f"files[{i}].mapping.aggregates[{j}].op: required")
            wrt = _as_opt_str(a.get("wrt"), path=f"files[{i}].mapping.aggregates[{j}].wrt")
            output_column = _as_opt_str(
                a.get("output_column"), path=f"files[{i}].mapping.aggregates[{j}].output_column"
            )
            aggregates.append(Aggregate(name=a_name, source=a_source, op=op, wrt=wrt, output_column=output_column))

        files.append(
            FileSpec(
                file_id=file_id,
                path_column=path_column,
                format=fmt,
                sheet=sheet,
                read=read_opts,
                axes=axes,
                axis_types=axis_types,
                targets=targets,
                derived=derived,
                aggregates=aggregates,
            )
        )

    # Validate that every file path column is defined in condition columns
    cond_col_names = {c.name for c in condition_columns}
    missing = sorted(file_path_columns - cond_col_names)
    if missing:
        raise SpecError(f"files[*].path_column not found in condition.columns: {missing}")

    # Validate global uniqueness of canonical target/derived names across files
    seen_global_targets: dict[str, str] = {}
    for f in files:
        for name in f.all_target_names():
            prev = seen_global_targets.get(name)
            if prev and prev != f.file_id:
                raise SpecError(f"Duplicate target/derived name across files: {name} (files: {prev}, {f.file_id})")
            seen_global_targets[name] = f.file_id

    return DatasetFormatSpec(
        schema_version=1,
        clearml=clearml,
        template=template,
        addin=addin,
        condition_columns=condition_columns,
        files=files,
        output=output,
    )


def load_format_spec(path: str | Path) -> DatasetFormatSpec:
    spec_path = Path(path).expanduser().resolve()
    raw_text = spec_path.read_text(encoding="utf-8")

    try:
        import yaml
    except Exception as e:  # pragma: no cover
        raise RuntimeError("YAML spec requires 'PyYAML'. Install dependencies first.") from e

    raw = yaml.safe_load(raw_text)
    if not isinstance(raw, dict):
        raise SpecError("Spec root must be a mapping (YAML dict)")
    return parse_format_spec(raw, spec_path=spec_path)


def load_format_spec_from_mapping(raw: Mapping[str, Any]) -> DatasetFormatSpec:
    return parse_format_spec(raw, spec_path=None)


def spec_to_yaml_dict(spec: DatasetFormatSpec) -> dict[str, Any]:
    out: dict[str, Any] = {"schema_version": int(spec.schema_version)}

    if spec.clearml is not None:
        out["clearml"] = {
            "dataset_project": spec.clearml.dataset_project,
            "dataset_name": spec.clearml.dataset_name,
            "output_uri": spec.clearml.output_uri,
            "tags": list(spec.clearml.tags),
            "use_current_task": bool(spec.clearml.use_current_task),
        }
        if spec.clearml.execution is not None:
            out["clearml"]["execution"] = {
                "repository": spec.clearml.execution.repository,
                "branch": spec.clearml.execution.branch,
                "commit": spec.clearml.execution.commit,
                "working_dir": spec.clearml.execution.working_dir,
                "entry_point": spec.clearml.execution.entry_point,
            }

    out["template"] = {
        "condition_sheet": spec.template.condition_sheet,
        "meta_sheet": spec.template.meta_sheet,
        "template_filename": spec.template.template_filename,
    }

    out["addin"] = {
        "enabled": bool(spec.addin.enabled),
        "target_os": spec.addin.target_os,
        "spec_filename": spec.addin.spec_filename,
        "vba_module_filename": spec.addin.vba_module_filename,
        "vba_template_excel": spec.addin.vba_template_excel,
        "embed_vba": bool(spec.addin.embed_vba),
        "command": spec.addin.command,
        "command_mac": spec.addin.command_mac,
        "command_windows": spec.addin.command_windows,
        "windows_mode": spec.addin.windows_mode,
        "windows_template_filename": spec.addin.windows_template_filename,
        "windows_addin_filename": spec.addin.windows_addin_filename,
    }

    out["condition"] = {
        "columns": [
            {
                "name": c.name,
                "type": c.dtype,
                "required": bool(c.required),
                "description": c.description,
                "enum": list(c.enum),
            }
            for c in spec.condition_columns
        ]
    }

    files_out: list[dict[str, Any]] = []
    for f in spec.files:
        axes: dict[str, Any] = {}
        for k in ("x", "y", "z", "t"):
            v = getattr(f.axes, k)
            if v is not None:
                dt = getattr(f.axis_types, k, None)
                if isinstance(dt, str) and dt.strip():
                    axes[k] = {"source": v, "type": dt}
                else:
                    axes[k] = v

        files_out.append(
            {
                "id": f.file_id,
                "path_column": f.path_column,
                "format": f.format,
                "sheet": f.sheet,
                "read": dict(f.read),
                "mapping": {
                    "axes": axes,
                    "targets": [{"name": t.name, "source": t.source, "type": t.dtype} for t in f.targets],
                    "derived": [{"name": d.name, "expr": d.expr, "type": d.dtype} for d in f.derived],
                    "aggregates": [
                        {
                            "name": a.name,
                            "source": a.source,
                            "op": a.op,
                            "wrt": a.wrt,
                            "output_column": a.output_column,
                        }
                        for a in f.aggregates
                    ],
                },
            }
        )

    out["files"] = files_out

    out["output"] = {
        "output_dirname": spec.output.output_dirname,
        "canonical_filename": spec.output.canonical_filename,
        "conditions_filename": spec.output.conditions_filename,
        "consolidated_excel_filename": spec.output.consolidated_excel_filename,
        "include_file_path_columns": bool(spec.output.include_file_path_columns),
        "combine_mode": spec.output.combine_mode,
    }

    return out


def with_clearml_values(
    spec: DatasetFormatSpec,
    *,
    dataset_project: str,
    dataset_name: str,
    output_uri: str | None = None,
    tags: list[str] | None = None,
) -> DatasetFormatSpec:
    """
    Return a copy of the spec where `clearml.dataset_project/dataset_name` (and optionally `output_uri/tags`)
    are updated to the provided values.
    """
    if spec.clearml is None:
        clearml = ClearMLSpec(
            dataset_project=str(dataset_project),
            dataset_name=str(dataset_name),
            output_uri=output_uri,
            tags=list(tags or []),
            use_current_task=False,
            execution=None,
        )
    else:
        clearml = replace(
            spec.clearml,
            dataset_project=str(dataset_project),
            dataset_name=str(dataset_name),
            output_uri=output_uri,
            tags=list(tags) if tags is not None else list(spec.clearml.tags),
        )
    return replace(spec, clearml=clearml)


def dump_spec_yaml(spec: DatasetFormatSpec) -> str:
    try:
        import yaml
    except Exception as e:  # pragma: no cover
        raise RuntimeError("Writing spec requires 'PyYAML'. Install dependencies first.") from e

    return yaml.safe_dump(spec_to_yaml_dict(spec), allow_unicode=True, sort_keys=False) + "\n"


def write_spec_yaml(spec: DatasetFormatSpec, path: str | Path) -> Path:
    out = Path(path).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(dump_spec_yaml(spec), encoding="utf-8")
    return out
