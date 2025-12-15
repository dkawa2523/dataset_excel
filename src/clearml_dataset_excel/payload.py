from __future__ import annotations

import json
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any


class PayloadError(RuntimeError):
    pass


@dataclass(frozen=True)
class PayloadMeta:
    root: Path
    meta: dict[str, Any]

    @property
    def payload_version(self) -> int:
        v = self.meta.get("payload_version", 0)
        try:
            return int(v)
        except Exception:
            return 0

    def relpath(self, key: str) -> str | None:
        v = self.meta.get(key)
        if isinstance(v, str) and v.strip():
            return v
        return None

    def abspath(self, key: str) -> Path | None:
        rel = self.relpath(key)
        if not rel:
            return None
        return (self.root / rel).resolve()

    def exists(self, key: str) -> bool:
        p = self.abspath(key)
        return bool(p and p.exists())


def load_payload_meta(root: str | Path) -> PayloadMeta:
    root_path = Path(root).expanduser().resolve()
    payload_path = root_path / "payload.json"
    if not payload_path.exists():
        raise PayloadError(f"payload.json not found: {payload_path}")
    try:
        meta = json.loads(payload_path.read_text(encoding="utf-8"))
    except Exception as e:
        raise PayloadError(f"Failed to parse payload.json: {payload_path}") from e
    if not isinstance(meta, dict):
        raise PayloadError("payload.json must be a JSON object")
    return PayloadMeta(root=root_path, meta=meta)


def validate_payload(root: str | Path) -> list[str]:
    """
    Validate payload.json and referenced files. Returns a list of human-readable errors.
    """
    errors: list[str] = []
    try:
        payload = load_payload_meta(root)
    except PayloadError as e:
        return [str(e)]

    if payload.payload_version > 1:
        errors.append(f"Unsupported payload_version: {payload.payload_version}")

    # Required in all payloads
    for k in ("spec_path", "template_excel"):
        rel = payload.relpath(k)
        if not rel:
            errors.append(f"Missing key: {k}")
        else:
            p = payload.abspath(k)
            if p is None or not p.exists():
                errors.append(f"Missing file for {k}: {rel}")

    # Optional template companions
    for k in ("template_spec", "vba_module", "template_excel_windows", "addin_xlam_windows", "runner_exe_windows"):
        rel = payload.relpath(k)
        if rel:
            p = payload.abspath(k)
            if p is None or not p.exists():
                errors.append(f"Missing file for {k}: {rel}")

    # Dataset payload keys (optional for template-only distribution)
    for k in ("condition_excel", "conditions_csv", "canonical_csv", "consolidated_excel"):
        rel = payload.relpath(k)
        if rel:
            p = payload.abspath(k)
            if p is None or not p.exists():
                errors.append(f"Missing file for {k}: {rel}")

    # path_map integrity
    pm = payload.meta.get("path_map")
    if pm is not None and not isinstance(pm, dict):
        errors.append("path_map must be an object (raw path -> relative path)")
    elif isinstance(pm, dict):
        for raw, rel in list(pm.items())[:2000]:
            if not isinstance(raw, str) or not isinstance(rel, str):
                errors.append("path_map contains non-string key/value")
                break
            if not rel.strip():
                errors.append(f"path_map has empty value for: {raw!r}")
                break
            p = (payload.root / rel).resolve()
            if not p.exists():
                errors.append(f"path_map points to missing file: {raw!r} -> {rel}")
                break

    return errors


def validate_payload_deep(root: str | Path) -> list[str]:
    """
    Deep validation:
    - validate payload.json + file existence
    - load spec YAML
    - re-run processing in a temporary directory using payload path_map as fallback
    This checks that the dataset payload is actually reproducible on another machine.
    """
    errors = validate_payload(root)
    if errors:
        return errors

    try:
        payload = load_payload_meta(root)
    except PayloadError as e:
        return [str(e)]

    spec_path = payload.abspath("spec_path")
    if spec_path is None:
        return ["Missing key: spec_path"]

    try:
        from .format_spec import load_format_spec
    except Exception as e:  # pragma: no cover
        return [f"Deep validation requires format spec loader: {e}"]

    try:
        spec = load_format_spec(spec_path)
    except Exception as e:
        return [f"Deep validation failed: invalid spec ({spec_path.name}): {e}"]

    condition_excel = payload.abspath("condition_excel")
    if condition_excel is None:
        return ["Deep validation requires 'condition_excel' in payload.json (dataset payload only)."]

    pm = payload.meta.get("path_map")
    fallback_map: dict[str, str] | None = None
    if isinstance(pm, dict) and pm:
        fallback_map = {}
        for raw, rel in pm.items():
            if not isinstance(raw, str) or not isinstance(rel, str) or not rel.strip():
                continue
            p = (payload.root / rel).resolve()
            fallback_map[raw] = p.as_posix()

    try:
        from .format_processor import process_condition_excel
    except Exception as e:  # pragma: no cover
        return [f"Deep validation requires processor: {e}"]

    try:
        with tempfile.TemporaryDirectory(prefix="clearml_dataset_excel_payload_deep_") as td:
            process_condition_excel(
                spec,
                condition_excel,
                output_root=td,
                check_files_exist=True,
                path_fallback_map=fallback_map,
                fallback_search_root=condition_excel.parent,
            )
    except Exception as e:
        return [f"Deep validation failed: {e}"]

    # Validate template structure matches the spec (best-effort).
    try:
        template_excel = payload.abspath("template_excel")
        if template_excel is not None and template_excel.exists():
            import openpyxl

            wb = openpyxl.load_workbook(template_excel, keep_vba=True)
            try:
                required_sheets = {"Info", spec.template.condition_sheet, spec.template.meta_sheet}
                missing = sorted(required_sheets - set(wb.sheetnames))
                if missing:
                    return [f"Deep validation failed: template is missing sheets: {missing}"]
                if spec.addin.enabled:
                    if template_excel.suffix.lower() not in {".xlsm", ".xlsx"}:
                        return [
                            f"Deep validation failed: template_excel must be .xlsm/.xlsx when addin.enabled=true: {template_excel.name}"
                        ]

                    if spec.addin.windows_mode == "addin":
                        win_template = payload.abspath("template_excel_windows")
                        if win_template is None or not win_template.exists():
                            return [
                                "Deep validation failed: addin.windows_mode=addin requires template_excel_windows in payload.json"
                            ]
                        if win_template.suffix.lower() != ".xlsx":
                            return [
                                f"Deep validation failed: template_excel_windows must be .xlsx: {win_template.name}"
                            ]

                        win_addin = payload.abspath("addin_xlam_windows")
                        if win_addin is None or not win_addin.exists():
                            return [
                                "Deep validation failed: addin.windows_mode=addin requires addin_xlam_windows in payload.json"
                            ]
                        if win_addin.suffix.lower() != ".xlam":
                            return [
                                f"Deep validation failed: addin_xlam_windows must be .xlam: {win_addin.name}"
                            ]
            finally:
                try:
                    vba_archive = getattr(wb, "vba_archive", None)
                    if vba_archive is not None:
                        vba_archive.close()
                except Exception:
                    pass
    except Exception as e:
        return [f"Deep validation failed: failed to read template Excel: {e}"]

    return []
