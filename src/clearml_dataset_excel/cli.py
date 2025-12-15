from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Sequence

from .agent import AgentError, infer_or_raise_dataset_id, reprocess_dataset_from_task, stage_reprocess_dataset_from_task
from .config import load_yaml_config
from .format_clearml import stage_dataset_payload, stage_dataset_payload_to_dir, stage_template_payload, upload_dataset
from .format_excel import generate_condition_template, generate_condition_template_from_excel, generate_windows_addin_xlam
from .format_processor import ProcessingError, process_condition_excel
from .format_spec import SpecError, load_format_spec, with_clearml_values, write_spec_yaml
from .manifest import read_rows_from_manifest, write_manifest_csv
from .resolver import (
    ResolvedItem,
    calc_dataset_relpath,
    collect_local_dataset_paths,
    iter_local_files_with_wildcards,
    resolve_items,
)
from .utils import is_url
from .vba_addin import write_vba_module
from .vba_embedder import embed_vba_module_into_xlsm
from .wildcards import matches_any_wildcard
from .payload import PayloadError, load_payload_meta, validate_payload, validate_payload_deep
from .addin_inspect import inspect_addin_excel


def _create_manifest_parser(defaults: dict[str, Any] | None = None) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="clearml-dataset-excel")

    defaults = defaults or {}

    def _default_list(key: str) -> list[str]:
        v = defaults.get(key) or []
        return list(v) if isinstance(v, list) else [str(v)]

    parser.add_argument("--config", default=defaults.get("config"), help="YAML config file path")
    parser.add_argument(
        "--excel",
        "--manifest",
        dest="manifest",
        default=defaults.get("manifest"),
        required=defaults.get("manifest") is None,
        help="Input manifest path (.xlsx/.csv/.tsv)",
    )
    parser.add_argument("--sheet", default=defaults.get("sheet"), help="Excel sheet name (default: first sheet)")
    parser.add_argument("--base-dir", default=defaults.get("base_dir"), help="Base directory for resolving relative paths")

    parser.add_argument(
        "--path-col",
        default=defaults.get("path_col", "path"),
        help="Column name that contains file/dir path (default: path)",
    )
    parser.add_argument(
        "--dataset-path-col",
        default=defaults.get("dataset_path_col", "dataset_path"),
        help="Column name for dataset path prefix (default: dataset_path)",
    )
    parser.add_argument(
        "--dataset-path-template",
        default=defaults.get("dataset_path_template"),
        help="Python format template for dataset path prefix (overrides --dataset-path-col)",
    )

    parser.add_argument(
        "--dataset-project",
        default=defaults.get("dataset_project"),
        required=defaults.get("dataset_project") is None,
        help="ClearML dataset project name",
    )
    parser.add_argument(
        "--dataset-name",
        default=defaults.get("dataset_name"),
        required=defaults.get("dataset_name") is None,
        help="ClearML dataset name",
    )
    parser.add_argument("--dataset-version", default=defaults.get("dataset_version"), help="ClearML dataset version")
    parser.add_argument("--output-uri", default=defaults.get("output_uri"), help="Dataset output URI (e.g. s3://bucket/path)")
    parser.add_argument("--description", default=defaults.get("description"), help="Dataset description")
    parser.add_argument("--tags", action="append", default=_default_list("tags"), help="Dataset tag (can be repeated)")
    parser.add_argument("--parent", action="append", default=_default_list("parent"), help="Parent dataset ID (can be repeated)")
    parser.add_argument("--use-current-task", action="store_true", default=bool(defaults.get("use_current_task")), help="Use current ClearML Task for dataset")
    parser.add_argument("--max-workers", type=int, default=defaults.get("max_workers"), help="Max worker threads for hashing/upload")

    parser.add_argument("--manifest-name", default=defaults.get("manifest_name", "manifest.csv"), help="Manifest filename stored in dataset root")
    parser.add_argument("--no-manifest", action="store_true", default=bool(defaults.get("no_manifest")), help="Do not add manifest.csv into dataset")

    parser.add_argument(
        "--include",
        action="append",
        default=_default_list("include"),
        help="Include wildcard for directory sources (can be repeated, e.g. --include '*.jpg')",
    )
    parser.add_argument(
        "--exclude",
        action="append",
        default=_default_list("exclude"),
        help="Exclude wildcard on dataset paths / external links (can be repeated)",
    )
    parser.add_argument("--skip-missing", action="store_true", default=bool(defaults.get("skip_missing")), help="Skip missing local paths instead of failing")
    parser.add_argument("--recursive", action=argparse.BooleanOptionalAction, default=defaults.get("recursive", True))
    parser.add_argument(
        "--collision-policy",
        choices=["ignore", "warn", "error"],
        default=defaults.get("collision_policy", "warn"),
        help="How to handle local dataset path collisions (default: warn)",
    )
    parser.add_argument("--auto-upload", action=argparse.BooleanOptionalAction, default=defaults.get("auto_upload", True))
    parser.add_argument("--dry-run", action="store_true", default=bool(defaults.get("dry_run")), help="Print actions without creating dataset")
    parser.add_argument("--dry-run-list", action="store_true", default=bool(defaults.get("dry_run_list")), help="List resolved dataset paths for local files")
    parser.add_argument(
        "--dry-run-max-files",
        type=int,
        default=int(defaults.get("dry_run_max_files", 200)),
        help="Max file lines printed in --dry-run-list",
    )
    parser.add_argument(
        "--dry-run-max-items",
        type=int,
        default=int(defaults.get("dry_run_max_items", 20)),
        help="Max item lines printed in --dry-run",
    )
    parser.add_argument("--verbose", action="store_true", default=bool(defaults.get("verbose")))

    return parser


def _iter_candidate_files(
    item: ResolvedItem,
    *,
    recursive: bool,
    include: list[str] | None,
) -> tuple[Path, Path, list[Path], list[str] | None]:
    source_path = Path(item.source)
    local_base_folder = Path(item.local_base_folder or source_path.parent)

    item_wildcards: list[str] | None
    if item.wildcard:
        item_wildcards = [item.wildcard]
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

    return source_path, local_base_folder, files, item_wildcards


def _print_dry_run(*, items: list[ResolvedItem], rows: list[dict], skipped: int, args: argparse.Namespace) -> int:
    local_items = [it for it in items if not is_url(it.source)]
    external_items = [it for it in items if is_url(it.source)]

    print(f"Rows: {len(rows)}")
    print(f"Items: {len(items)} (local={len(local_items)} external={len(external_items)} skipped={skipped})")

    matched_local_files = 0
    selected_local_files = 0
    excluded_local_files = 0
    unique_dataset_paths: dict[str, str] = {}
    collisions: dict[str, list[str]] = {}

    printed_files = 0
    printed_items = 0
    max_items = max(0, args.dry_run_max_items)
    max_files = max(0, args.dry_run_max_files)

    for it in items:
        if is_url(it.source):
            if printed_items < max_items:
                print(f"- external source={it.source} dataset_path={it.dataset_path}")
                printed_items += 1
            continue

        source_path, local_base_folder, files, _ = _iter_candidate_files(
            it,
            recursive=args.recursive,
            include=args.include or None,
        )
        matched_local_files += len(files)

        item_was_printed = printed_items < max_items
        if item_was_printed:
            if source_path.is_dir():
                glob_note = f" glob={it.wildcard}" if it.wildcard else ""
                print(
                    f"- dir source={it.source} dataset_path={it.dataset_path} local_base_folder={local_base_folder} files={len(files)}{glob_note}"
                )
            else:
                dataset_relpath = calc_dataset_relpath(
                    file_path=source_path,
                    local_base_folder=local_base_folder,
                    dataset_path=it.dataset_path,
                )
                include_skipped = args.include and not files
                exclude_skipped = args.exclude and matches_any_wildcard(
                    dataset_relpath, args.exclude, recursive=args.recursive
                )
                if include_skipped:
                    print(
                        f"- file source={it.source} dataset_path={it.dataset_path} local_base_folder={local_base_folder} skipped (include)"
                    )
                elif exclude_skipped:
                    print(
                        f"- file source={it.source} dataset_path={it.dataset_path} local_base_folder={local_base_folder} skipped (exclude) -> {dataset_relpath}"
                    )
                else:
                    print(
                        f"- file source={it.source} dataset_path={it.dataset_path} local_base_folder={local_base_folder} -> {dataset_relpath}"
                    )
            printed_items += 1

        for f in files:
            dataset_relpath = calc_dataset_relpath(
                file_path=f,
                local_base_folder=local_base_folder,
                dataset_path=it.dataset_path,
            )
            if args.exclude and matches_any_wildcard(dataset_relpath, args.exclude, recursive=args.recursive):
                excluded_local_files += 1
                continue

            selected_local_files += 1
            existing = unique_dataset_paths.get(dataset_relpath)
            if existing and existing != f.as_posix():
                collisions.setdefault(dataset_relpath, [existing]).append(f.as_posix())
            else:
                unique_dataset_paths.setdefault(dataset_relpath, f.as_posix())

            if args.dry_run_list and item_was_printed and printed_files < max_files:
                print(f"{dataset_relpath}\t<- {f.as_posix()}")
                printed_files += 1

    if len(items) > max_items:
        print(f"... (printed {max_items} items, {len(items) - max_items} more items)")
    if args.dry_run_list and printed_files >= max_files:
        print(f"... (printed {printed_files} files, truncated by --dry-run-max-files)")

    print(
        f"Local files: {selected_local_files} (matched={matched_local_files} excluded={excluded_local_files} unique dataset paths: {len(unique_dataset_paths)})"
    )

    if collisions:
        print(f"Warning: dataset path collisions detected: {len(collisions)}")
        for i, (relpath, sources) in enumerate(collisions.items()):
            if i >= 20:
                print(f"... ({len(collisions) - 20} more collisions)")
                break
            print(f"  - {relpath} <- {sources[0]} (and {len(sources) - 1} more)")

    return 0


def _main_manifest(argv: Sequence[str] | None = None) -> int:
    # Parse --config first, then apply YAML defaults (CLI args override YAML)
    pre_parser = argparse.ArgumentParser(add_help=False)
    pre_parser.add_argument("--config", default=None)
    pre_args, _ = pre_parser.parse_known_args(argv)

    defaults: dict[str, Any] = {}
    if pre_args.config:
        defaults = load_yaml_config(pre_args.config)
        defaults["config"] = pre_args.config

    args = _create_manifest_parser(defaults).parse_args(argv)

    manifest_path = Path(args.manifest).expanduser().resolve()
    if not manifest_path.exists():
        print(f"Manifest not found: {manifest_path}", file=sys.stderr)
        return 2

    base_dir = Path(args.base_dir).expanduser().resolve() if args.base_dir else None
    if base_dir and not base_dir.is_dir():
        print(f"--base-dir is not a directory: {base_dir}", file=sys.stderr)
        return 2

    rows, columns = read_rows_from_manifest(manifest_path, args.sheet)
    if args.path_col not in columns:
        print(f"Missing required column: {args.path_col}", file=sys.stderr)
        return 2

    items, skipped = resolve_items(
        rows,
        path_col=args.path_col,
        dataset_path_col=args.dataset_path_col if args.dataset_path_col else None,
        dataset_path_template=args.dataset_path_template,
        base_dir=base_dir,
        skip_missing=args.skip_missing,
    )

    if args.dry_run:
        return _print_dry_run(items=items, rows=rows, skipped=skipped, args=args)

    if args.collision_policy != "ignore":
        _, collisions, matched_local_files, excluded_local_files = collect_local_dataset_paths(
            items,
            recursive=args.recursive,
            include=args.include or None,
            exclude=args.exclude or None,
        )
        if collisions:
            print(f"Warning: dataset path collisions detected: {len(collisions)}", file=sys.stderr)
            for i, (relpath, sources) in enumerate(collisions.items()):
                if i >= 20:
                    print(f"... ({len(collisions) - 20} more collisions)", file=sys.stderr)
                    break
                print(f"  - {relpath} <- {sources[0]} (and {len(sources) - 1} more)", file=sys.stderr)
            print(f"(local matched={matched_local_files} excluded={excluded_local_files})", file=sys.stderr)
            if args.collision_policy == "error":
                return 1

    try:
        from clearml import Dataset
    except Exception as e:  # pragma: no cover
        raise RuntimeError("clearml is required. Install dependencies first.") from e

    parent_datasets = args.parent or None
    dataset = Dataset.create(
        dataset_name=args.dataset_name,
        dataset_project=args.dataset_project,
        dataset_tags=args.tags or None,
        parent_datasets=parent_datasets,
        use_current_task=args.use_current_task,
        dataset_version=args.dataset_version,
        output_uri=args.output_uri,
        description=args.description,
    )

    temp_dir_obj = None
    try:
        if not args.no_manifest:
            from tempfile import TemporaryDirectory

            temp_dir_obj = TemporaryDirectory(prefix="clearml_dataset_excel_")
            manifest_out = Path(temp_dir_obj.name) / args.manifest_name
            write_manifest_csv(manifest_out, rows)
            dataset.add_files(manifest_out.as_posix(), dataset_path=".", max_workers=args.max_workers)

        for it in items:
            if is_url(it.source):
                dataset.add_external_files(
                    source_url=it.source,
                    dataset_path=it.dataset_path,
                    recursive=args.recursive,
                    verbose=args.verbose,
                    max_workers=args.max_workers,
                )
                continue

            source_path = Path(it.source)
            local_base_folder = Path(it.local_base_folder) if it.local_base_folder else source_path.parent

            if source_path.is_file():
                rel = source_path.relative_to(local_base_folder).as_posix()
                if args.include and not matches_any_wildcard(rel, args.include, recursive=args.recursive):
                    continue
                dataset_relpath = calc_dataset_relpath(
                    file_path=source_path, local_base_folder=local_base_folder, dataset_path=it.dataset_path
                )
                if args.exclude and matches_any_wildcard(dataset_relpath, args.exclude, recursive=args.recursive):
                    continue
                dataset.add_files(
                    path=it.source,
                    local_base_folder=it.local_base_folder,
                    dataset_path=it.dataset_path,
                    recursive=args.recursive,
                    verbose=args.verbose,
                    max_workers=args.max_workers,
                )
            else:
                wildcards: list[str] | None
                if it.wildcard:
                    wildcards = [it.wildcard]
                elif args.include:
                    wildcards = args.include
                else:
                    wildcards = None

                dataset.add_files(
                    path=it.source,
                    wildcard=wildcards,
                    local_base_folder=it.local_base_folder,
                    dataset_path=it.dataset_path,
                    recursive=args.recursive,
                    verbose=args.verbose,
                    max_workers=args.max_workers,
                )

        if args.exclude:
            for pattern in args.exclude:
                dataset.remove_files(dataset_path=pattern, recursive=args.recursive, verbose=args.verbose)

        if not args.auto_upload:
            dataset.upload(
                show_progress=args.verbose,
                verbose=args.verbose,
                output_url=args.output_uri,
                max_workers=args.max_workers,
            )

        ok = dataset.finalize(verbose=args.verbose, auto_upload=args.auto_upload)
        if not ok:
            print("Dataset finalize failed", file=sys.stderr)
            return 1
    finally:
        if temp_dir_obj:
            temp_dir_obj.cleanup()

    print(dataset.id)
    return 0


def _create_template_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="clearml-dataset-excel template")
    sub = parser.add_subparsers(dest="template_cmd", required=True)

    gen = sub.add_parser("generate", help="Generate condition Excel template from YAML spec")
    gen.add_argument("--spec", required=True, help="YAML spec path")
    gen.add_argument("--output", default=None, help="Output .xlsx/.xlsm path (default: from spec)")
    gen.add_argument(
        "--base-excel",
        default=None,
        help="Base .xlsm/.xlsx to copy for template generation (preserves VBA/extra sheets if present)",
    )
    gen.add_argument("--overwrite", action="store_true", help="Overwrite output file if exists")

    pkg = sub.add_parser("package", help="Generate templates and bundle into a distributable .zip")
    pkg.add_argument("--spec", required=True, help="YAML spec path")
    pkg.add_argument("--output", default=None, help="Output .zip path (default: <spec_dir>/template_package.zip)")
    pkg.add_argument(
        "--base-excel",
        default=None,
        help="Base .xlsm/.xlsx to copy for template generation (preserves VBA/extra sheets if present)",
    )
    pkg.add_argument("--overwrite", action="store_true", help="Overwrite output .zip if exists")

    val = sub.add_parser("validate", help="Validate YAML spec (static)")
    val.add_argument("--spec", required=True, help="YAML spec path")
    return parser


def _resolve_spec_relative_path(spec_path: str | Path, value: str | None) -> Path | None:
    if not value:
        return None
    p = Path(value).expanduser()
    if p.is_absolute():
        return p.resolve()
    base = Path(spec_path).expanduser().resolve().parent
    return (base / p).resolve()


def _resolve_runner_exe_windows(*, spec_path: str | Path, excel_path: str | Path | None = None) -> Path | None:
    """
    Best-effort lookup for the optional Windows runner executable.
    This file is useful for distributing a Python-less "Run" experience on Windows.
    """
    runner_name = "clearml_dataset_excel_runner.exe"

    candidates: list[Path] = []
    try:
        candidates.append(Path(spec_path).expanduser().resolve().parent / runner_name)
    except Exception:
        pass
    if excel_path:
        try:
            candidates.append(Path(excel_path).expanduser().resolve().parent / runner_name)
        except Exception:
            pass
    try:
        candidates.append(Path.cwd().resolve() / runner_name)
    except Exception:
        pass
    try:
        # When running from a source checkout: <repo>/src/clearml_dataset_excel/cli.py -> <repo>/dist/<exe>
        candidates.append(Path(__file__).resolve().parents[2] / "dist" / runner_name)
    except Exception:
        pass

    for p in candidates:
        try:
            if p.exists() and p.is_file():
                return p.resolve()
        except Exception:
            continue
    return None


def _main_template(argv: Sequence[str] | None = None) -> int:
    args = _create_template_parser().parse_args(argv)
    try:
        spec = load_format_spec(args.spec)
    except (SpecError, FileNotFoundError) as e:
        print(str(e), file=sys.stderr)
        return 2

    if args.template_cmd == "validate":
        print("OK")
        return 0

    if args.template_cmd == "package":
        import shutil
        import zipfile
        from tempfile import TemporaryDirectory

        spec_src = Path(args.spec).expanduser().resolve()
        spec_filename = Path(spec.addin.spec_filename or spec_src.name).name
        spec_dir = spec_src.parent

        out_zip = args.output
        if not out_zip:
            out_zip = (spec_dir / "template_package.zip").as_posix()
        out_zip_path = Path(out_zip).expanduser().resolve()

        if out_zip_path.exists() and not bool(args.overwrite):
            print(f"Output already exists: {out_zip_path}", file=sys.stderr)
            return 1
        out_zip_path.parent.mkdir(parents=True, exist_ok=True)
        if out_zip_path.exists():
            try:
                out_zip_path.unlink()
            except Exception:
                pass

        with TemporaryDirectory(prefix="clearml_dataset_excel_template_package_") as td:
            td_path = Path(td).resolve()
            gen_dir = td_path / "gen"
            gen_dir.mkdir(parents=True, exist_ok=True)

            template_path = gen_dir / spec.template.template_filename
            try:
                if args.base_excel:
                    generate_condition_template_from_excel(
                        args.base_excel,
                        spec,
                        template_path,
                        overwrite=True,
                        clear_conditions_data=True,
                    )
                else:
                    generate_condition_template(spec, template_path, overwrite=True)
            except Exception as e:
                print(str(e), file=sys.stderr)
                return 1

            vba_path = None
            if spec.addin.enabled:
                vba_path = gen_dir / spec.addin.vba_module_filename
                write_vba_module(vba_path, meta_sheet_name=spec.template.meta_sheet)
                if spec.addin.embed_vba:
                    try:
                        vba_template = _resolve_spec_relative_path(args.spec, spec.addin.vba_template_excel)
                        embed_vba_module_into_xlsm(
                            excel_path=template_path,
                            overwrite=True,
                            template_excel=vba_template,
                            bas_path=vba_path if vba_template is not None else None,
                        )
                    except Exception as e:
                        # Best-effort: keep the .bas next to the template as a manual fallback.
                        print(str(e), file=sys.stderr)

            win_template_path = None
            win_addin_path = None
            if spec.addin.enabled and spec.addin.windows_mode == "addin":
                try:
                    win_template_path = gen_dir / spec.addin.windows_template_filename
                    generate_condition_template(spec, win_template_path, overwrite=True)
                    win_addin_path = gen_dir / spec.addin.windows_addin_filename
                    generate_windows_addin_xlam(win_addin_path, overwrite=True)
                except Exception as e:
                    print(str(e), file=sys.stderr)
                    return 1

            runner_exe = _resolve_runner_exe_windows(spec_path=args.spec, excel_path=args.base_excel) if spec.addin.enabled else None

            pkg_dir = td_path / "package"
            pkg_dir.mkdir(parents=True, exist_ok=True)

            readme_lines = [
                "clearml-dataset-excel template package",
                "",
                "mac/:",
                f"- {template_path.name}",
                f"- {spec_filename}",
                f"- {spec.addin.vba_module_filename}" if spec.addin.enabled else "- (no add-in)",
                "",
                "windows/ (addin.windows_mode=addin only):",
                f"- {spec.addin.windows_template_filename}" if win_template_path is not None else "- (not generated)",
                f"- {spec.addin.windows_addin_filename}" if win_addin_path is not None else "- (not generated)",
                f"- {spec_filename}" if win_template_path is not None else "",
                "- clearml_dataset_excel_runner.exe (optional)" if runner_exe is not None else "",
                "",
                "Usage:",
            ]

            mac_embedded = bool(spec.addin.enabled and spec.addin.embed_vba and template_path.suffix.lower() == ".xlsm")
            if mac_embedded:
                readme_lines.append("- macOS: open the .xlsm, enable macros, then use the 'ClearML' ribbon tab (Run).")
            elif spec.addin.enabled:
                readme_lines.extend(
                    [
                        "- macOS: import the .bas into the workbook (VBE -> File -> Import File), then run ClearMLDatasetExcel_Run.",
                        "  (or set addin.embed_vba: true in YAML and regenerate)",
                    ]
                )
            else:
                readme_lines.append("- macOS: open the template and run the CLI manually.")

            if win_template_path is not None and win_addin_path is not None:
                readme_lines.append("- Windows: install the .xlam, open the .xlsx, then use the 'ClearML' ribbon tab (Run).")
            else:
                readme_lines.append("- Windows: (not included) set addin.windows_mode: addin in YAML and regenerate.")

            if runner_exe is not None:
                cmd_w = (spec.addin.command_windows or spec.addin.command or "").strip()
                if cmd_w and runner_exe.name not in cmd_w:
                    readme_lines.extend(
                        [
                            "",
                            "Note:",
                            "- runner.exe is included, but addin.command_windows does not reference it.",
                            "  Update YAML to prefer runner.exe (see examples/run.yaml).",
                        ]
                    )

            readme_lines.extend(["", "Logs are written next to the workbook: clearml_dataset_excel_addin.log", ""])
            readme = "\n".join(readme_lines).strip() + "\n"
            (pkg_dir / "README.txt").write_text(readme, encoding="utf-8")

            mac_dir = pkg_dir / "mac"
            mac_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(template_path, mac_dir / template_path.name)
            shutil.copy2(spec_src, mac_dir / spec_filename)
            if vba_path is not None and vba_path.exists():
                shutil.copy2(vba_path, mac_dir / vba_path.name)

            if win_template_path is not None and win_addin_path is not None:
                win_dir = pkg_dir / "windows"
                win_dir.mkdir(parents=True, exist_ok=True)
                shutil.copy2(win_template_path, win_dir / win_template_path.name)
                shutil.copy2(win_addin_path, win_dir / win_addin_path.name)
                shutil.copy2(spec_src, win_dir / spec_filename)
                if runner_exe is not None:
                    try:
                        shutil.copy2(runner_exe, win_dir / runner_exe.name)
                    except Exception:
                        pass

            try:
                with zipfile.ZipFile(out_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
                    for p in sorted(pkg_dir.rglob("*")):
                        if p.is_file():
                            z.write(p, p.relative_to(pkg_dir).as_posix())
            except Exception as e:
                print(str(e), file=sys.stderr)
                return 1

        print(out_zip_path.as_posix())
        return 0

    out = args.output
    if not out:
        spec_dir = Path(args.spec).expanduser().resolve().parent
        out = (spec_dir / spec.template.template_filename).as_posix()

    try:
        if args.base_excel:
            generated = generate_condition_template_from_excel(
                args.base_excel,
                spec,
                out,
                overwrite=bool(args.overwrite),
                clear_conditions_data=True,
            )
        else:
            generated = generate_condition_template(spec, out, overwrite=bool(args.overwrite))
    except Exception as e:
        print(str(e), file=sys.stderr)
        return 1

    if spec.addin.enabled:
        import shutil

        out_dir = generated.parent

        spec_src = Path(args.spec).expanduser().resolve()
        spec_dst = out_dir / (spec.addin.spec_filename or spec_src.name)
        if spec_src.resolve() != spec_dst.resolve():
            shutil.copy2(spec_src, spec_dst)

        vba_dst = out_dir / spec.addin.vba_module_filename
        write_vba_module(vba_dst, meta_sheet_name=spec.template.meta_sheet)
        if spec.addin.embed_vba:
            try:
                vba_template = _resolve_spec_relative_path(args.spec, spec.addin.vba_template_excel)
                embed_vba_module_into_xlsm(
                    excel_path=generated,
                    overwrite=True,
                    template_excel=vba_template,
                    bas_path=vba_dst if vba_template is not None else None,
                )
            except Exception as e:
                # Best-effort: if embedding fails (permissions, missing Excel, etc), user can import .bas manually.
                print(str(e), file=sys.stderr)

        win_template = None
        win_addin = None
        if spec.addin.windows_mode == "addin":
            try:
                if spec.addin.windows_template_filename:
                    win_template = generate_condition_template(
                        spec,
                        out_dir / spec.addin.windows_template_filename,
                        overwrite=bool(args.overwrite),
                    )
                win_addin = generate_windows_addin_xlam(
                    out_dir / spec.addin.windows_addin_filename,
                    overwrite=bool(args.overwrite),
                )
            except Exception as e:
                print(str(e), file=sys.stderr)
                return 1

        runner_exe = _resolve_runner_exe_windows(spec_path=args.spec, excel_path=generated)
        if runner_exe is not None:
            dst = out_dir / runner_exe.name
            try:
                if dst.exists() and not bool(args.overwrite):
                    pass
                else:
                    try:
                        if dst.exists() and dst.samefile(runner_exe):
                            pass
                        else:
                            shutil.copy2(runner_exe, dst)
                    except Exception:
                        shutil.copy2(runner_exe, dst)
            except Exception:
                # Best-effort; do not fail template generation.
                pass

        print(generated.as_posix())
        print(spec_dst.as_posix())
        print(vba_dst.as_posix())
        if win_template is not None:
            print(win_template.as_posix())
        if win_addin is not None:
            print(win_addin.as_posix())
        return 0

    print(generated.as_posix())
    return 0


def _create_run_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="clearml-dataset-excel run")
    parser.add_argument("--spec", required=True, help="YAML spec path")
    parser.add_argument("--excel", required=True, help="Filled condition Excel path (.xlsx/.xlsm)")
    parser.add_argument("--sheet", default=None, help="Condition sheet name (default: from spec)")
    parser.add_argument("--output-root", default=None, help="Output root directory (default: Excel folder)")
    parser.add_argument(
        "--stage-dir",
        default=None,
        help="Stage payload directory to keep (default: auto when --no-upload; temp when uploading)",
    )
    parser.add_argument("--overwrite-stage", action="store_true", help="Overwrite --stage-dir if it exists")
    parser.add_argument("--no-upload", action="store_true", help="Process only (do not upload to ClearML)")
    parser.add_argument("--dataset-project", default=None, help="Override ClearML dataset project")
    parser.add_argument("--dataset-name", default=None, help="Override ClearML dataset name")
    parser.add_argument("--output-uri", default=None, help="Override ClearML dataset output_uri")
    parser.add_argument("--description", default=None, help="Dataset description (create only)")
    parser.add_argument("--tags", action="append", default=[], help="Dataset tags (repeatable)")
    parser.add_argument("--max-workers", type=int, default=None, help="Max worker threads for upload")
    parser.add_argument("--verbose", action="store_true", help="Verbose ClearML upload logs")
    return parser


def _create_register_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="clearml-dataset-excel register")
    parser.add_argument("--spec", required=True, help="YAML spec path")
    parser.add_argument(
        "--base-excel",
        default=None,
        help="Base .xlsm/.xlsx to copy for template generation (preserves VBA/extra sheets if present)",
    )
    parser.add_argument("--dataset-project", default=None, help="Override ClearML dataset project")
    parser.add_argument("--dataset-name", default=None, help="Override ClearML dataset name")
    parser.add_argument("--output-uri", default=None, help="Override ClearML dataset output_uri")
    parser.add_argument("--description", default=None, help="Dataset description (create only)")
    parser.add_argument("--tags", action="append", default=[], help="Dataset tags (repeatable)")
    parser.add_argument("--max-workers", type=int, default=None, help="Max worker threads for upload")
    parser.add_argument("--verbose", action="store_true", help="Verbose ClearML upload logs")
    return parser


def _main_register(argv: Sequence[str] | None = None) -> int:
    args = _create_register_parser().parse_args(argv)
    try:
        spec = load_format_spec(args.spec)
    except (SpecError, FileNotFoundError) as e:
        print(str(e), file=sys.stderr)
        return 2

    dataset_project = args.dataset_project or (spec.clearml.dataset_project if spec.clearml else None)
    dataset_name = args.dataset_name or (spec.clearml.dataset_name if spec.clearml else None)
    output_uri = args.output_uri or (spec.clearml.output_uri if spec.clearml else None)
    if not dataset_project or not dataset_name:
        print("dataset_project/dataset_name is required (set in spec.clearml or CLI override)", file=sys.stderr)
        return 2

    resolved_tags = args.tags or (spec.clearml.tags if spec.clearml else None)
    spec_for_upload = with_clearml_values(
        spec,
        dataset_project=dataset_project,
        dataset_name=dataset_name,
        output_uri=output_uri,
        tags=resolved_tags,
    )

    from tempfile import TemporaryDirectory

    with TemporaryDirectory(prefix="clearml_dataset_excel_register_") as td:
        # Write a spec copy that matches the resolved ClearML dataset (important for the downloaded template+addin).
        spec_copy_path = write_spec_yaml(spec_for_upload, Path(td) / Path(args.spec).name)

        template_path = Path(td) / spec.template.template_filename
        try:
            if args.base_excel:
                generate_condition_template_from_excel(
                    args.base_excel,
                    spec,
                    template_path,
                    overwrite=True,
                    clear_conditions_data=True,
                )
            else:
                generate_condition_template(spec, template_path, overwrite=True)
        except Exception as e:
            print(str(e), file=sys.stderr)
            return 1

        vba_path = None
        if spec.addin.enabled:
            vba_path = Path(td) / spec.addin.vba_module_filename
            write_vba_module(vba_path, meta_sheet_name=spec.template.meta_sheet)
            if spec.addin.embed_vba:
                try:
                    vba_template = _resolve_spec_relative_path(args.spec, spec.addin.vba_template_excel)
                    embed_vba_module_into_xlsm(
                        excel_path=template_path,
                        overwrite=True,
                        template_excel=vba_template,
                        bas_path=vba_path if vba_template is not None else None,
                    )
                except Exception as e:
                    # Best-effort: if embedding fails (permissions, missing Excel, etc), user can import .bas manually.
                    print(str(e), file=sys.stderr)

        win_template_path = None
        win_addin_path = None
        if spec.addin.enabled and spec.addin.windows_mode == "addin":
            try:
                if spec.addin.windows_template_filename:
                    win_template_path = Path(td) / spec.addin.windows_template_filename
                    generate_condition_template(spec, win_template_path, overwrite=True)
                win_addin_path = Path(td) / spec.addin.windows_addin_filename
                generate_windows_addin_xlam(win_addin_path, overwrite=True)
            except Exception as e:
                print(str(e), file=sys.stderr)
                return 1

        runner_exe_windows = _resolve_runner_exe_windows(spec_path=args.spec, excel_path=args.base_excel) if spec.addin.enabled else None

        stage_td = stage_template_payload(
            spec_path=spec_copy_path,
            spec=spec_for_upload,
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
                dataset_project=dataset_project,
                dataset_name=dataset_name,
                tags=resolved_tags,
                output_uri=output_uri,
                description=args.description,
                max_workers=args.max_workers,
                verbose=bool(args.verbose),
            )
        finally:
            stage_td.cleanup()

    print(ds_id)
    return 0


def _main_run(argv: Sequence[str] | None = None) -> int:
    args = _create_run_parser().parse_args(argv)
    try:
        spec = load_format_spec(args.spec)
    except (SpecError, FileNotFoundError) as e:
        print(str(e), file=sys.stderr)
        return 2

    try:
        outputs = process_condition_excel(
            spec,
            args.excel,
            sheet_name=args.sheet,
            output_root=args.output_root,
            check_files_exist=True,
        )
    except ProcessingError as e:
        print(str(e), file=sys.stderr)
        return 1

    # Generate template Excel alongside upload payload
    from tempfile import TemporaryDirectory

    with TemporaryDirectory(prefix="clearml_dataset_excel_template_") as td:
        spec_copy_path = Path(args.spec).expanduser().resolve()

        template_path = Path(td) / spec.template.template_filename
        try:
            excel_in = Path(args.excel).expanduser().resolve()
            if spec.addin.enabled and excel_in.suffix.lower() == ".xlsm":
                generate_condition_template_from_excel(excel_in, spec, template_path, overwrite=True)
            else:
                generate_condition_template(spec, template_path, overwrite=True)
        except Exception as e:
            print(str(e), file=sys.stderr)
            return 1

        vba_path = None
        if spec.addin.enabled:
            vba_path = Path(td) / spec.addin.vba_module_filename
            write_vba_module(vba_path, meta_sheet_name=spec.template.meta_sheet)
        if spec.addin.enabled and spec.addin.embed_vba:
            try:
                vba_template = _resolve_spec_relative_path(args.spec, spec.addin.vba_template_excel)
                embed_vba_module_into_xlsm(
                    excel_path=template_path,
                    overwrite=True,
                    template_excel=vba_template,
                    bas_path=vba_path if vba_template is not None else None,
                )
            except Exception as e:
                # Best-effort: if embedding fails (permissions, missing Excel, etc), user can import .bas manually.
                print(str(e), file=sys.stderr)

        win_template_path = None
        win_addin_path = None
        if spec.addin.enabled and spec.addin.windows_mode == "addin":
            try:
                if spec.addin.windows_template_filename:
                    win_template_path = Path(td) / spec.addin.windows_template_filename
                    generate_condition_template(spec, win_template_path, overwrite=True)
                win_addin_path = Path(td) / spec.addin.windows_addin_filename
                generate_windows_addin_xlam(win_addin_path, overwrite=True)
            except Exception as e:
                print(str(e), file=sys.stderr)
                return 1

        stage_dir_override = Path(args.stage_dir).expanduser().resolve() if args.stage_dir else None
        if stage_dir_override is None and args.no_upload:
            stage_dir_override = outputs.output_dir / "_clearml_stage"

        runner_exe_windows = _resolve_runner_exe_windows(spec_path=args.spec, excel_path=args.excel) if spec.addin.enabled else None

        dataset_project = None
        dataset_name = None
        output_uri = None
        resolved_tags = None
        spec_for_upload = spec
        if not args.no_upload:
            dataset_project = args.dataset_project or (spec.clearml.dataset_project if spec.clearml else None)
            dataset_name = args.dataset_name or (spec.clearml.dataset_name if spec.clearml else None)
            output_uri = args.output_uri or (spec.clearml.output_uri if spec.clearml else None)
            if not dataset_project or not dataset_name:
                print("dataset_project/dataset_name is required (set in spec.clearml or CLI override)", file=sys.stderr)
                return 2

            resolved_tags = args.tags or (spec.clearml.tags if spec.clearml else None)
            spec_for_upload = with_clearml_values(
                spec,
                dataset_project=dataset_project,
                dataset_name=dataset_name,
                output_uri=output_uri,
                tags=resolved_tags,
            )
            # Write a spec copy that matches the resolved ClearML dataset (important for the downloaded template+addin).
            spec_copy_path = write_spec_yaml(spec_for_upload, Path(td) / Path(args.spec).name)

        if stage_dir_override is not None:
            try:
                stage_dir = stage_dataset_payload_to_dir(
                    stage_dir=stage_dir_override,
                    spec_path=spec_copy_path,
                    spec=spec_for_upload,
                    condition_excel=Path(args.excel).expanduser().resolve(),
                    outputs=outputs,
                    template_excel=template_path,
                    template_excel_windows=win_template_path,
                    addin_xlam_windows=win_addin_path,
                    vba_module=vba_path,
                    runner_exe_windows=runner_exe_windows,
                    overwrite=bool(args.overwrite_stage),
                )
            except Exception as e:
                print(str(e), file=sys.stderr)
                return 1

            if args.no_upload:
                print(outputs.output_dir.as_posix())
                print(stage_dir.as_posix())
                return 0

            ds_id = upload_dataset(
                stage_dir=stage_dir,
                spec=spec_for_upload,
                dataset_project=str(dataset_project),
                dataset_name=str(dataset_name),
                tags=resolved_tags,
                output_uri=output_uri,
                description=args.description,
                max_workers=args.max_workers,
                verbose=bool(args.verbose),
            )
        else:
            stage_td = stage_dataset_payload(
                spec_path=spec_copy_path,
                spec=spec_for_upload,
                condition_excel=Path(args.excel).expanduser().resolve(),
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
                    dataset_project=str(dataset_project),
                    dataset_name=str(dataset_name),
                    tags=resolved_tags,
                    output_uri=output_uri,
                    description=args.description,
                    max_workers=args.max_workers,
                    verbose=bool(args.verbose),
                )
            finally:
                stage_td.cleanup()

    print(ds_id)
    return 0


def _create_agent_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="clearml-dataset-excel agent")
    sub = parser.add_subparsers(dest="subcommand")

    rp = sub.add_parser("reprocess", help="Reprocess a ClearML Dataset using spec from current Task configuration")
    rp.add_argument("--config-name", default="dataset_format_spec", help="Task Configuration name (default: dataset_format_spec)")
    rp.add_argument("--dataset-id", default=None, help="Base dataset id to download/version (default: infer from current Task)")
    rp.add_argument("--dataset-project", default=None, help="Output dataset project (override spec.clearml.dataset_project)")
    rp.add_argument("--dataset-name", default=None, help="Output dataset name (override spec.clearml.dataset_name)")
    rp.add_argument("--output-uri", default=None, help="Output URI (override spec.clearml.output_uri)")
    rp.add_argument("--description", default=None, help="Dataset description")
    rp.add_argument("--tags", action="append", default=[], help="Dataset tag (can be repeated)")
    rp.add_argument("--max-workers", type=int, default=None, help="Max worker threads for hashing/upload")
    rp.add_argument("--excel", dest="condition_excel", default=None, help="Condition Excel path inside the dataset (override payload.json)")
    rp.add_argument("--sheet", default=None, help="Condition sheet name (override spec.template.condition_sheet)")
    rp.add_argument("--no-upload", action="store_true", help="Process and stage only (do not upload to ClearML)")
    rp.add_argument("--output-root", default=None, help="Output root directory when --no-upload (default: current dir)")
    rp.add_argument(
        "--stage-dir",
        default=None,
        help="Stage payload directory to keep when --no-upload (default: <output_dir>/_clearml_stage)",
    )
    rp.add_argument("--overwrite-stage", action="store_true", help="Overwrite --stage-dir if it exists")
    rp.add_argument("--verbose", action="store_true", default=False, help="Verbose ClearML upload")

    return parser


def _create_addin_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="clearml-dataset-excel addin")
    sub = parser.add_subparsers(dest="subcommand", required=True)

    bld = sub.add_parser("build", help="Build a Windows Excel add-in (.xlam) containing ClearMLDatasetExcel_Run")
    bld.add_argument("--output", required=True, help="Output .xlam path")
    bld.add_argument("--overwrite", action="store_true", help="Overwrite output file if exists")

    loc = sub.add_parser("locate", help="Print the default Excel AddIns directory for this platform")
    loc.add_argument("--json", action="store_true", help="Print JSON")

    emb = sub.add_parser(
        "embed",
        help="Embed add-in VBA into an .xlsm (template copy or best-effort Excel automation)",
    )
    emb.add_argument("--excel", required=True, help="Target .xlsm path")
    emb.add_argument(
        "--bas",
        required=False,
        help="VBA module (.bas) path (optional when using bundled default or --template-excel)",
    )
    emb.add_argument(
        "--template-excel",
        default=None,
        help="Template .xlsm that already contains the add-in macro (copies vbaProject.bin; no Excel automation)",
    )
    emb.add_argument("--overwrite", action="store_true", help="Overwrite module if it already exists")

    ins = sub.add_parser("inspect", help="Inspect an .xlsm/.xlsx/.xlam for add-in meta and embedded VBA")
    ins.add_argument("--excel", required=True, help="Target .xlsm/.xlsx path")
    ins.add_argument("--meta-sheet", default="_meta", help="Meta sheet name (default: _meta)")
    ins.add_argument("--json", action="store_true", help="Print JSON")

    inst = sub.add_parser("install", help="Install a .xlam into the user's Excel AddIns directory (best-effort)")
    inst.add_argument("--xlam", required=True, help="Source .xlam path")
    inst.add_argument(
        "--dest",
        default=None,
        help="Destination directory (default: platform-specific, e.g. %APPDATA%/Microsoft/AddIns on Windows)",
    )
    inst.add_argument("--name", default=None, help="Destination file name (default: keep original)")
    inst.add_argument("--overwrite", action="store_true", help="Overwrite if destination exists")

    upd = sub.add_parser("update", help="Update an installed .xlam (backs up existing, then overwrites)")
    upd.add_argument("--xlam", required=True, help="Source .xlam path")
    upd.add_argument(
        "--dest",
        default=None,
        help="Destination directory (default: platform-specific, e.g. %APPDATA%/Microsoft/AddIns on Windows)",
    )
    upd.add_argument("--name", default=None, help="Destination file name (default: keep original)")
    upd.add_argument("--no-backup", action="store_true", help="Do not keep a .bak copy of the existing add-in")

    uninst = sub.add_parser("uninstall", help="Remove an installed .xlam from the Excel AddIns directory")
    uninst.add_argument("--path", default=None, help="Installed .xlam path to remove (overrides --dest/--name)")
    uninst.add_argument(
        "--dest",
        default=None,
        help="Destination directory (default: platform-specific, e.g. %APPDATA%/Microsoft/AddIns on Windows)",
    )
    uninst.add_argument("--name", default="clearml_dataset_excel_addin.xlam", help="Installed add-in file name")

    uq = sub.add_parser("unquarantine", help="Remove macOS quarantine xattr from an Excel file (best-effort)")
    uq.add_argument("--excel", required=True, help="Target .xlsm/.xlsx/.xlam path")
    return parser


def _create_payload_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="clearml-dataset-excel payload")
    sub = parser.add_subparsers(dest="subcommand", required=True)

    show = sub.add_parser("show", help="Show payload.json contents")
    show.add_argument("--root", required=True, help="Dataset payload root directory (contains payload.json)")

    val = sub.add_parser("validate", help="Validate payload.json and referenced files")
    val.add_argument("--root", required=True, help="Dataset payload root directory (contains payload.json)")
    val.add_argument(
        "--deep",
        action="store_true",
        help="Deep validation: re-run processing in a temp dir using payload spec/condition_excel/path_map",
    )
    return parser


def _main_payload(argv: Sequence[str] | None = None) -> int:
    args = _create_payload_parser().parse_args(argv)
    if args.subcommand not in {"show", "validate"}:
        print("Unknown payload subcommand", file=sys.stderr)
        return 2

    if args.subcommand == "show":
        try:
            payload = load_payload_meta(args.root)
        except PayloadError as e:
            print(str(e), file=sys.stderr)
            return 2
        import json

        print(json.dumps(payload.meta, ensure_ascii=False, indent=2, sort_keys=True))
        return 0

    errors = validate_payload_deep(args.root) if bool(getattr(args, "deep", False)) else validate_payload(args.root)
    if errors:
        for e in errors:
            print(e, file=sys.stderr)
        return 1
    print("OK")
    return 0


def _main_addin(argv: Sequence[str] | None = None) -> int:
    args = _create_addin_parser().parse_args(argv)
    if args.subcommand not in {"build", "locate", "embed", "inspect", "install", "update", "uninstall", "unquarantine"}:
        print("Unknown addin subcommand", file=sys.stderr)
        return 2

    def resolve_excel_addins_dir(dest: str | None) -> Path:
        import os

        if dest:
            return Path(dest).expanduser().resolve()
        if sys.platform.startswith("win"):
            appdata = os.environ.get("APPDATA")
            if not appdata:
                raise RuntimeError("APPDATA is not set; provide --dest")
            return (Path(appdata) / "Microsoft" / "AddIns").resolve()
        if sys.platform == "darwin":
            # Note: Excel for macOS supports .xlam, but this project primarily uses workbook macros on macOS.
            return (
                Path.home()
                / "Library"
                / "Group Containers"
                / "UBF8T346G9.Office"
                / "User Content"
                / "Add-Ins"
            ).resolve()
        raise RuntimeError("Unknown platform; provide --dest")

    if args.subcommand == "build":
        try:
            out = generate_windows_addin_xlam(args.output, overwrite=bool(args.overwrite))
        except Exception as e:
            print(str(e), file=sys.stderr)
            return 1
        print(out.as_posix())
        return 0

    if args.subcommand == "locate":
        try:
            dst = resolve_excel_addins_dir(None)
        except Exception as e:
            print(str(e), file=sys.stderr)
            return 2
        if bool(getattr(args, "json", False)):
            import json

            print(json.dumps({"excel_addins_dir": dst.as_posix()}, ensure_ascii=False, indent=2, sort_keys=True))
        else:
            print(dst.as_posix())
        return 0

    if args.subcommand == "inspect":
        try:
            info = inspect_addin_excel(args.excel, meta_sheet_name=str(args.meta_sheet))
        except Exception as e:
            print(str(e), file=sys.stderr)
            return 1

        if args.json:
            import json

            print(json.dumps(info, ensure_ascii=False, indent=2, sort_keys=True))
        else:
            print(f"excel: {info.get('excel')}")
            print(f"has_vba_project: {info.get('has_vba_project')}")
            print(f"has_clearml_macro: {info.get('has_clearml_macro')}")
            print(f"meta_sheet: {info.get('meta_sheet')}")
            meta = info.get("meta") if isinstance(info.get("meta"), dict) else {}
            for k in sorted(meta.keys()):
                if str(k).startswith("addin_"):
                    print(f"meta.{k}: {meta.get(k)}")
        return 0

    if args.subcommand == "install":
        import shutil

        src = Path(args.xlam).expanduser().resolve()
        if not src.exists():
            print(f".xlam not found: {src}", file=sys.stderr)
            return 2
        if src.suffix.lower() != ".xlam":
            print(f"--xlam must be .xlam: {src}", file=sys.stderr)
            return 2

        try:
            dst_dir = resolve_excel_addins_dir(args.dest)
        except Exception as e:
            print(str(e), file=sys.stderr)
            return 2

        name = str(args.name).strip() if args.name else src.name
        if not name.lower().endswith(".xlam"):
            name = name + ".xlam"

        dst_dir.mkdir(parents=True, exist_ok=True)
        dst = (dst_dir / name).resolve()
        if dst.exists() and not bool(args.overwrite):
            print(f"Destination exists (use --overwrite): {dst}", file=sys.stderr)
            return 1

        shutil.copy2(src, dst)
        print(dst.as_posix())
        if sys.platform.startswith("win"):
            print("Next: Excel -> File -> Options -> Add-ins -> Manage: Excel Add-ins -> Go... -> Browse -> enable the add-in.")
        elif sys.platform == "darwin":
            print("Next: Excel -> Tools -> Excel Add-ins... -> Browse -> enable the add-in.")
        return 0

    if args.subcommand == "update":
        import shutil
        from datetime import datetime

        src = Path(args.xlam).expanduser().resolve()
        if not src.exists():
            print(f".xlam not found: {src}", file=sys.stderr)
            return 2
        if src.suffix.lower() != ".xlam":
            print(f"--xlam must be .xlam: {src}", file=sys.stderr)
            return 2

        try:
            dst_dir = resolve_excel_addins_dir(args.dest)
        except Exception as e:
            print(str(e), file=sys.stderr)
            return 2

        name = str(args.name).strip() if args.name else src.name
        if not name.lower().endswith(".xlam"):
            name = name + ".xlam"

        dst_dir.mkdir(parents=True, exist_ok=True)
        dst = (dst_dir / name).resolve()

        backup_path = None
        if dst.exists() and not bool(getattr(args, "no_backup", False)):
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = dst.with_suffix(dst.suffix + f".bak.{ts}")
            try:
                shutil.copy2(dst, backup_path)
            except Exception:
                backup_path = None

        shutil.copy2(src, dst)
        print(dst.as_posix())
        if backup_path is not None:
            print(backup_path.as_posix())
        return 0

    if args.subcommand == "uninstall":
        target = None
        if getattr(args, "path", None):
            target = Path(args.path).expanduser().resolve()
        else:
            try:
                dst_dir = resolve_excel_addins_dir(args.dest)
            except Exception as e:
                print(str(e), file=sys.stderr)
                return 2
            name = str(args.name).strip() if args.name else "clearml_dataset_excel_addin.xlam"
            if not name.lower().endswith(".xlam"):
                name = name + ".xlam"
            target = (dst_dir / name).resolve()

        if not target.exists():
            print(f"Not found: {target}", file=sys.stderr)
            return 1
        if target.suffix.lower() != ".xlam":
            print(f"Refusing to remove non-.xlam file: {target}", file=sys.stderr)
            return 2

        try:
            target.unlink()
        except Exception as e:
            print(str(e), file=sys.stderr)
            return 1
        print(target.as_posix())
        return 0

    if args.subcommand == "unquarantine":
        from .utils import clear_macos_quarantine

        p = Path(args.excel).expanduser().resolve()
        if not p.exists():
            print(f"Excel not found: {p}", file=sys.stderr)
            return 2
        ok = clear_macos_quarantine(p)
        print(p.as_posix())
        if sys.platform == "darwin":
            print("removed" if ok else "not_present_or_failed")
        else:
            print("noop")
        return 0

    try:
        embed_vba_module_into_xlsm(
            excel_path=args.excel,
            bas_path=args.bas,
            overwrite=bool(args.overwrite),
            template_excel=args.template_excel,
        )
    except NotImplementedError as e:
        print(str(e), file=sys.stderr)
        return 2
    except Exception as e:
        print(str(e), file=sys.stderr)
        return 1

    print(Path(args.excel).expanduser().resolve().as_posix())
    return 0


def _main_agent(argv: Sequence[str] | None = None) -> int:
    argv_list = list(argv) if argv is not None else []
    if not argv_list or (argv_list and argv_list[0].startswith("-")):
        argv_list = ["reprocess", *argv_list]

    args = _create_agent_parser().parse_args(argv_list)
    if args.subcommand != "reprocess":
        print("Unknown agent subcommand", file=sys.stderr)
        return 2

    try:
        from clearml import Task
    except Exception as e:  # pragma: no cover
        print("clearml is required. Install dependencies first.", file=sys.stderr)
        return 2

    task = Task.current_task()
    if task is None:
        print("No current ClearML Task found. Run this inside a ClearML Agent task.", file=sys.stderr)
        return 2

    try:
        dataset_id = infer_or_raise_dataset_id(task, args.dataset_id)
        if args.no_upload:
            staged = stage_reprocess_dataset_from_task(
                task=task,
                dataset_id=dataset_id,
                config_name=args.config_name,
                condition_excel=args.condition_excel,
                sheet_name=args.sheet,
                output_root=args.output_root,
                stage_dir=args.stage_dir,
                overwrite_stage=bool(args.overwrite_stage),
            )
            print(staged.output_dir.as_posix())
            print(staged.stage_dir.as_posix())
            return 0

        ds_id = reprocess_dataset_from_task(
            task=task,
            dataset_id=dataset_id,
            dataset_project=args.dataset_project,
            dataset_name=args.dataset_name,
            config_name=args.config_name,
            condition_excel=args.condition_excel,
            sheet_name=args.sheet,
            output_uri=args.output_uri,
            tags=args.tags or None,
            description=args.description,
            max_workers=args.max_workers,
            verbose=bool(args.verbose),
        )
    except AgentError as e:
        print(str(e), file=sys.stderr)
        return 1

    print(ds_id)
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    argv_list = list(argv) if argv is not None else sys.argv[1:]

    if argv_list and argv_list[0] in {"manifest", "template", "run", "register", "agent", "addin", "payload"}:
        cmd = argv_list[0]
        sub_argv: Sequence[str] = argv_list[1:]
    else:
        cmd = "manifest"
        sub_argv = argv_list

    if cmd == "manifest":
        return _main_manifest(sub_argv)
    if cmd == "template":
        return _main_template(sub_argv)
    if cmd == "run":
        return _main_run(sub_argv)
    if cmd == "register":
        return _main_register(sub_argv)
    if cmd == "agent":
        return _main_agent(sub_argv)
    if cmd == "addin":
        return _main_addin(sub_argv)
    if cmd == "payload":
        return _main_payload(sub_argv)

    print(f"Unknown command: {cmd}", file=sys.stderr)
    return 2


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
