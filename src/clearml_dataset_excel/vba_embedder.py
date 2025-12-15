from __future__ import annotations

import re
import sys
import tempfile
import subprocess
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from .vba_project import vba_project_has_symbol
from .utils import clear_macos_quarantine


def _parse_vb_module_name_from_bas(bas_path: Path) -> str | None:
    try:
        head = bas_path.read_text(encoding="utf-8", errors="ignore").splitlines()[:20]
    except Exception:
        return None
    for line in head:
        m = re.match(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$', line)
        if m:
            return m.group(1)
    return None


def _vba_project_contains_symbol(xlsm_path: Path, symbol: str) -> bool:
    try:
        with zipfile.ZipFile(xlsm_path, "r") as z:
            try:
                data = z.read("xl/vbaProject.bin")
            except KeyError:
                return False
    except Exception:
        return False
    try:
        return vba_project_has_symbol(data, symbol)
    except Exception:
        return False


_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_APP_PROPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_VBA_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
_VBA_CONTENT_TYPE = "application/vnd.ms-office.vbaProject"
_WORKBOOK_CT_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
_WORKBOOK_CT_XLAM = "application/vnd.ms-excel.addin.macroEnabled.main+xml"
_CUSTOM_UI_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
_CUSTOM_UI_CONTENT_TYPE = "application/vnd.ms-office.customUI+xml"
_CUSTOM_UI_PART = "customUI/customUI.xml"


def _default_custom_ui_xml(*, on_action: str) -> bytes:
    # Use the Office 2007 customUI namespace for broad compatibility.
    # The callback must be a Public Sub with signature: (control As IRibbonControl).
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">\n'
        "  <ribbon>\n"
        "    <tabs>\n"
        '      <tab id="clearml_dataset_excel_tab" label="ClearML">\n'
        '        <group id="clearml_dataset_excel_group" label="Dataset Excel">\n'
        f'          <button id="clearml_dataset_excel_run" label="Run" size="large" imageMso="Play" onAction="{on_action}"/>\n'
        "        </group>\n"
        "      </tab>\n"
        "    </tabs>\n"
        "  </ribbon>\n"
        "</customUI>\n"
    ).encode("utf-8")


def _patch_content_types_for_vba(content_types_xml: bytes, *, workbook_content_type: str) -> bytes:
    ET.register_namespace("", _CONTENT_TYPES_NS)
    root = ET.fromstring(content_types_xml)
    ns = {"ct": _CONTENT_TYPES_NS}

    # Ensure workbook.xml uses macro-enabled content type.
    workbook_override = None
    for elem in root.findall("ct:Override", ns):
        if elem.attrib.get("PartName") == "/xl/workbook.xml":
            workbook_override = elem
            break
    if workbook_override is None:
        workbook_override = ET.SubElement(
            root,
            f"{{{_CONTENT_TYPES_NS}}}Override",
            {"PartName": "/xl/workbook.xml", "ContentType": workbook_content_type},
        )
    else:
        workbook_override.attrib["ContentType"] = workbook_content_type

    # Ensure vbaProject.bin has the VBA content type (use an Override to avoid affecting other .bin parts).
    vba_override = None
    for elem in root.findall("ct:Override", ns):
        if elem.attrib.get("PartName") == "/xl/vbaProject.bin":
            vba_override = elem
            break
    if vba_override is None:
        ET.SubElement(
            root,
            f"{{{_CONTENT_TYPES_NS}}}Override",
            {"PartName": "/xl/vbaProject.bin", "ContentType": _VBA_CONTENT_TYPE},
        )
    else:
        vba_override.attrib["ContentType"] = _VBA_CONTENT_TYPE

    return ET.tostring(root, encoding="utf-8")


def _patch_workbook_rels_for_vba(workbook_rels_xml: bytes) -> bytes:
    ET.register_namespace("", _REL_NS)
    root = ET.fromstring(workbook_rels_xml)
    ns = {"r": _REL_NS}

    rel_elems = list(root.findall("r:Relationship", ns))
    for rel in rel_elems:
        if rel.attrib.get("Type") == _VBA_REL_TYPE:
            rel.attrib["Target"] = "vbaProject.bin"
            return ET.tostring(root, encoding="utf-8")

    existing_ids = {rel.attrib.get("Id") for rel in rel_elems if rel.attrib.get("Id")}
    max_num = 0
    for rid in existing_ids:
        m = re.match(r"^rId(\d+)$", str(rid))
        if m:
            max_num = max(max_num, int(m.group(1)))

    new_id = f"rId{max_num + 1}"
    while new_id in existing_ids:
        max_num += 1
        new_id = f"rId{max_num + 1}"

    ET.SubElement(
        root,
        f"{{{_REL_NS}}}Relationship",
        {"Id": new_id, "Type": _VBA_REL_TYPE, "Target": "vbaProject.bin"},
    )
    return ET.tostring(root, encoding="utf-8")


def _patch_workbook_xml_for_vba(workbook_xml: bytes) -> bytes:
    ET.register_namespace("", _SHEET_NS)
    root = ET.fromstring(workbook_xml)
    ns = {"s": _SHEET_NS}

    # Some Excel builds (notably macOS) behave better when fileVersion exists.
    fv = root.find("s:fileVersion", ns)
    if fv is None:
        fv = ET.Element(
            f"{{{_SHEET_NS}}}fileVersion",
            {"appName": "xl", "lastEdited": "7", "lowestEdited": "7", "rupBuild": "11207"},
        )
        root.insert(0, fv)
    else:
        fv.attrib.setdefault("appName", "xl")

    wbpr = root.find("s:workbookPr", ns)
    if wbpr is None:
        wbpr = ET.Element(f"{{{_SHEET_NS}}}workbookPr")
        # Keep fileVersion at the very top when present.
        insert_idx = 1 if len(root) > 0 and root[0].tag.endswith("fileVersion") else 0
        root.insert(insert_idx, wbpr)

    # Macro workbooks typically have the VBA codename set. Ensure it exists for robust ThisWorkbook usage.
    if not wbpr.attrib.get("codeName"):
        wbpr.attrib["codeName"] = "ThisWorkbook"

    return ET.tostring(root, encoding="utf-8")


def _patch_package_rels_add_or_update(rels_xml: bytes, *, rel_type: str, target: str) -> bytes:
    ET.register_namespace("", _REL_NS)
    root = ET.fromstring(rels_xml)
    ns = {"r": _REL_NS}

    rel_elems = list(root.findall("r:Relationship", ns))
    for rel in rel_elems:
        if rel.attrib.get("Type") == rel_type:
            rel.attrib["Target"] = target
            return ET.tostring(root, encoding="utf-8")

    existing_ids = {rel.attrib.get("Id") for rel in rel_elems if rel.attrib.get("Id")}
    max_num = 0
    for rid in existing_ids:
        m = re.match(r"^rId(\d+)$", str(rid))
        if m:
            max_num = max(max_num, int(m.group(1)))

    new_id = f"rId{max_num + 1}"
    while new_id in existing_ids:
        max_num += 1
        new_id = f"rId{max_num + 1}"

    ET.SubElement(
        root,
        f"{{{_REL_NS}}}Relationship",
        {"Id": new_id, "Type": str(rel_type), "Target": str(target)},
    )
    return ET.tostring(root, encoding="utf-8")


def _patch_content_types_add_override(content_types_xml: bytes, *, part_name: str, content_type: str) -> bytes:
    ET.register_namespace("", _CONTENT_TYPES_NS)
    root = ET.fromstring(content_types_xml)
    ns = {"ct": _CONTENT_TYPES_NS}

    override = None
    for elem in root.findall("ct:Override", ns):
        if elem.attrib.get("PartName") == part_name:
            override = elem
            break
    if override is None:
        ET.SubElement(
            root,
            f"{{{_CONTENT_TYPES_NS}}}Override",
            {"PartName": str(part_name), "ContentType": str(content_type)},
        )
    else:
        override.attrib["ContentType"] = str(content_type)

    return ET.tostring(root, encoding="utf-8")


def _patch_app_xml_for_excel_compat(app_xml: bytes) -> bytes:
    """
    Best-effort: make docProps/app.xml look like an Excel-authored workbook.

    This should be cosmetic, but some Excel builds appear to be picky about workbooks created by other writers.
    """
    ET.register_namespace("", _APP_PROPS_NS)
    root = ET.fromstring(app_xml)
    ns = {"a": _APP_PROPS_NS}

    app = root.find("a:Application", ns)
    if app is None:
        app = ET.SubElement(root, f"{{{_APP_PROPS_NS}}}Application")
    app_text = (app.text or "").strip()
    app_was_openpyxl = (not app_text) or ("openpyxl" in app_text.lower())
    if app_was_openpyxl:
        app.text = "Microsoft Excel"

    av = root.find("a:AppVersion", ns)
    if av is None:
        av = ET.SubElement(root, f"{{{_APP_PROPS_NS}}}AppVersion")
    av_text = (av.text or "").strip()
    if app_was_openpyxl or (not av_text) or av_text.startswith(("3.", "0.")):
        av.text = "16.0300"

    return ET.tostring(root, encoding="utf-8")


def _patch_worksheet_xml_for_vba(worksheet_xml: bytes, *, code_name: str) -> bytes:
    ET.register_namespace("", _SHEET_NS)
    root = ET.fromstring(worksheet_xml)
    ns = {"s": _SHEET_NS}

    sheet_pr = root.find("s:sheetPr", ns)
    if sheet_pr is None:
        sheet_pr = ET.Element(f"{{{_SHEET_NS}}}sheetPr")
        root.insert(0, sheet_pr)

    if not sheet_pr.attrib.get("codeName"):
        sheet_pr.attrib["codeName"] = str(code_name)

    return ET.tostring(root, encoding="utf-8")


def _rewrite_zip_in_place(path: Path, *, replacements: dict[str, bytes]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        replaced_names = set(replacements.keys())
        for info in zin.infolist():
            name = info.filename
            if name in replaced_names:
                continue
            zout.writestr(info, zin.read(name))

        now = datetime.now().timetuple()[:6]
        for name, data in replacements.items():
            zi = zipfile.ZipInfo(filename=name, date_time=now)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, data)
    tmp.replace(path)


def _embed_vba_from_vba_project_bin(*, excel_path: Path, vba_bin: bytes, overwrite: bool) -> None:
    try:
        with zipfile.ZipFile(excel_path, "r") as z:
            names = set(z.namelist())
            has_vba_project = "xl/vbaProject.bin" in names

            ct_xml = z.read("[Content_Types].xml")
            package_rels_xml = z.read("_rels/.rels") if "_rels/.rels" in names else None
            rels_xml = z.read("xl/_rels/workbook.xml.rels")
            workbook_xml = z.read("xl/workbook.xml")
            app_xml = z.read("docProps/app.xml") if "docProps/app.xml" in names else None

            worksheet_names = sorted(
                [n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")],
                key=lambda n: int(re.findall(r"sheet(\d+)\.xml$", n)[0])
                if re.findall(r"sheet(\d+)\.xml$", n)
                else 10**9,
            )
            worksheets_xml: dict[str, bytes] = {}
            for n in worksheet_names:
                try:
                    worksheets_xml[n] = z.read(n)
                except KeyError:
                    continue
    except Exception as e:
        raise RuntimeError(f"Failed to read target Excel as zip: {excel_path}") from e

    workbook_ct = _WORKBOOK_CT_XLAM if excel_path.suffix.lower() == ".xlam" else _WORKBOOK_CT_XLSM
    patched_ct = _patch_content_types_for_vba(ct_xml, workbook_content_type=workbook_ct)
    patched_rels = _patch_workbook_rels_for_vba(rels_xml)
    patched_workbook = _patch_workbook_xml_for_vba(workbook_xml)
    patched_app = _patch_app_xml_for_excel_compat(app_xml) if app_xml is not None else None

    patched_sheets: dict[str, bytes] = {}
    # Excel typically uses Sheet1/Sheet2... as default codenames for worksheets.
    for name, xml in worksheets_xml.items():
        m = re.search(r"sheet(\d+)\.xml$", name)
        if m:
            code_name = f"Sheet{int(m.group(1))}"
        else:
            # Fallback: should be rare; keep it stable anyway.
            code_name = "Sheet1"
        patched_sheets[name] = _patch_worksheet_xml_for_vba(xml, code_name=code_name)

    replacements: dict[str, bytes] = {
        "[Content_Types].xml": patched_ct,
        "xl/_rels/workbook.xml.rels": patched_rels,
        "xl/workbook.xml": patched_workbook,
        **patched_sheets,
    }
    if patched_app is not None:
        replacements["docProps/app.xml"] = patched_app

    # Add a Ribbon UI button (one-click) when no customUI is present yet.
    try:
        has_custom_ui = any(n in names for n in {_CUSTOM_UI_PART, "customUI/customUI14.xml"})
        if not has_custom_ui and package_rels_xml:
            patched_ct = _patch_content_types_add_override(
                patched_ct, part_name="/" + _CUSTOM_UI_PART, content_type=_CUSTOM_UI_CONTENT_TYPE
            )
            patched_package_rels = _patch_package_rels_add_or_update(
                package_rels_xml,
                rel_type=_CUSTOM_UI_REL_TYPE,
                target=_CUSTOM_UI_PART,
            )
            replacements["[Content_Types].xml"] = patched_ct
            replacements["_rels/.rels"] = patched_package_rels
            replacements[_CUSTOM_UI_PART] = _default_custom_ui_xml(on_action="ClearMLDatasetExcel_Run_Ribbon")
    except Exception:
        pass

    # Only write/replace the VBA project when needed. Even when we don't replace it, we still patch
    # content types / relationships / codeNames because some writers (e.g. openpyxl keep_vba=True)
    # can drop the vbaProject override, causing Excel to treat the file as "no macros".
    if overwrite or not has_vba_project:
        replacements["xl/vbaProject.bin"] = vba_bin

    _rewrite_zip_in_place(
        excel_path,
        replacements=replacements,
    )


def _embed_vba_from_template_excel(*, excel_path: Path, template_excel: Path, overwrite: bool) -> None:
    """
    Embed VBA by copying xl/vbaProject.bin from an existing macro-enabled workbook and patching the OOXML parts.
    This avoids Excel/COM/UI automation and works cross-platform.
    """
    if not template_excel.exists():
        raise FileNotFoundError(f"VBA template Excel not found: {template_excel}")

    try:
        with zipfile.ZipFile(template_excel, "r") as donor:
            donor_names = set(donor.namelist())
            if "xl/vbaProject.bin" not in donor_names:
                raise RuntimeError(f"VBA template does not contain xl/vbaProject.bin: {template_excel}")
            vba_bin = donor.read("xl/vbaProject.bin")
    except Exception as e:
        raise RuntimeError(f"Failed to read VBA template Excel: {template_excel}") from e

    if not vba_project_has_symbol(vba_bin, "ClearMLDatasetExcel_Run"):
        raise RuntimeError(
            "VBA template's vbaProject.bin does not contain ClearMLDatasetExcel_Run.\n"
            "Provide a template that already has the ClearML Dataset Excel add-in macro embedded."
        )

    _embed_vba_from_vba_project_bin(excel_path=excel_path, vba_bin=vba_bin, overwrite=overwrite)


def _embed_vba_on_macos(*, excel_path: Path, bas_path: Path) -> None:
    """
    macOS: Best-effort embedding via Microsoft Excel + AppleScript.

    Excel's AppleScript dictionary does not expose the VBProject model, so we automate the built-in
    "VBA Insert File" dialog and select the .bas via UI scripting (System Events).
    This requires macOS Accessibility permission for the script runner (Terminal/iTerm/VSCode/etc).
    """
    script = r"""
on run argv
  with timeout of 300 seconds
    set xlsmPath to item 1 of argv
    set basPath to item 2 of argv

    tell application "Microsoft Excel"
      activate
      set scratchWb to missing value
      if (count of workbooks) is 0 then
        make new workbook
        set scratchWb to active workbook
      end if
      try
        set display alerts to false
      end try
      open POSIX file xlsmPath
    end tell

    delay 0.6

    tell application "Microsoft Excel"
      set wb to active workbook
    end tell

    tell application "Microsoft Excel"
      activate
      set d to get dialog dialog vba insert file
      show d
    end tell

    delay 0.6

    tell application "System Events"
      tell process "Microsoft Excel"
        set frontmost to true
        keystroke "G" using {command down, shift down}
        delay 0.2
        keystroke basPath
        keystroke return
        delay 0.2
        keystroke return
      end tell
    end tell

    delay 0.6

    tell application "Microsoft Excel"
      save wb
      close wb saving yes
      if scratchWb is not missing value then
        close scratchWb saving no
      end if
    end tell
  end timeout
end run
"""
    with tempfile.TemporaryDirectory(prefix="clearml_dataset_excel_applescript_") as td:
        scpt = Path(td) / "embed_bas.scpt"
        scpt.write_text(script, encoding="utf-8")
        try:
            subprocess.run(
                ["osascript", scpt.as_posix(), excel_path.as_posix(), bas_path.as_posix()],
                check=True,
                capture_output=True,
                text=True,
            )
        except subprocess.CalledProcessError as e:
            stderr = (e.stderr or "").strip()
            msg = stderr or (e.stdout or "").strip() or str(e)
            if (
                "-1743" in msg
                or "not authorized to send apple events" in msg.lower()
                or "not authorised to send apple events" in msg.lower()
            ):
                raise RuntimeError(
                    "macOS VBA embedding requires Automation permission to control Microsoft Excel.\n"
                    "Enable it in System Settings -> Privacy & Security -> Automation "
                    "for the app running this command (Terminal/iTerm/VSCode), then retry."
                ) from e
            if "-1712" in msg or "timed out" in msg.lower() or "timeout" in msg.lower():
                raise RuntimeError(
                    "macOS VBA embedding timed out while communicating with Microsoft Excel.\n"
                    "Excel may be busy or showing a modal dialog (permission prompts, file recovery, etc).\n"
                    "Bring Excel to the front, close any dialogs, ensure Automation/Accessibility permissions are granted, then retry."
                ) from e
            if (
                "-1719" in msg
                or "1002" in msg
                or "補助アクセス" in msg
                or "キー操作" in msg
                or "keystroke" in msg.lower()
                or "assistive access" in msg.lower()
            ):
                raise RuntimeError(
                    "macOS VBA embedding requires Accessibility permission for UI scripting.\n"
                    "Enable it in System Settings -> Privacy & Security -> Accessibility "
                    "for the app running this command (Terminal/iTerm/VSCode), then retry."
                ) from e
            raise RuntimeError(f"macOS VBA embedding failed: {msg}") from e


def embed_vba_module_into_xlsm(
    *,
    excel_path: str | Path,
    bas_path: str | Path | None = None,
    overwrite: bool = False,
    template_excel: str | Path | None = None,
) -> None:
    """
    Best-effort embedding of a .bas VBA module into an .xlsm file.

    If template_excel is provided, VBA is embedded by copying vbaProject.bin from the template (no Excel needed).
    If bas_path and template_excel are both omitted, a bundled default vbaProject.bin is embedded (no Excel needed).

    - Windows: uses Excel COM automation (requires Excel installed and 'Trust access to the VBA project object model').
    - macOS: uses Microsoft Excel + AppleScript UI scripting (requires Accessibility permission).
    - Other platforms: not supported.
    """
    xlsm = Path(excel_path).expanduser().resolve()
    if not xlsm.exists():
        raise FileNotFoundError(f"Excel not found: {xlsm}")
    if xlsm.suffix.lower() not in {".xlsm", ".xlam"}:
        raise ValueError(f"Target Excel must be .xlsm/.xlam for VBA embedding: {xlsm}")

    if template_excel is not None:
        template = Path(template_excel).expanduser().resolve()
        _embed_vba_from_template_excel(excel_path=xlsm, template_excel=template, overwrite=overwrite)
        clear_macos_quarantine(xlsm)
        return

    if bas_path is None:
        if not overwrite:
            # Fast path: if the workbook already contains a VBA project, just patch the OOXML metadata
            # without loading the bundled default vbaProject.bin.
            try:
                with zipfile.ZipFile(xlsm, "r") as z:
                    if "xl/vbaProject.bin" in set(z.namelist()):
                        existing_vba = z.read("xl/vbaProject.bin")

                        # Repair older templates that embedded the bundled donor VBA project.
                        # This keeps custom user macros intact while stripping known-incompatible parts
                        # (e.g., worksheet VB_Control attributes) when present.
                        try:
                            from .default_vba_project import patch_vba_project_bin_for_excel_compat

                            patched_vba = patch_vba_project_bin_for_excel_compat(existing_vba)
                        except Exception:
                            patched_vba = existing_vba

                        if patched_vba != existing_vba:
                            _embed_vba_from_vba_project_bin(excel_path=xlsm, vba_bin=patched_vba, overwrite=True)
                        else:
                            _embed_vba_from_vba_project_bin(excel_path=xlsm, vba_bin=b"", overwrite=False)
                        clear_macos_quarantine(xlsm)
                        return
            except Exception:
                pass

        from .default_vba_project import get_default_vba_project_bin

        _embed_vba_from_vba_project_bin(
            excel_path=xlsm,
            vba_bin=get_default_vba_project_bin(),
            overwrite=overwrite,
        )
        clear_macos_quarantine(xlsm)
        return

    bas = Path(bas_path).expanduser().resolve()
    if not bas.exists():
        raise FileNotFoundError(f"VBA module not found: {bas}")

    if sys.platform == "darwin":
        if _vba_project_contains_symbol(xlsm, "ClearMLDatasetExcel_Run"):
            if overwrite:
                raise RuntimeError(
                    "macOS VBA embedding: overwrite is not supported automatically when the workbook already "
                    "contains ClearMLDatasetExcel_Run. Remove the module manually or regenerate from a base .xlsm."
                )
            return
        _embed_vba_on_macos(excel_path=xlsm, bas_path=bas)
        clear_macos_quarantine(xlsm)
        return

    if not sys.platform.startswith("win"):
        raise NotImplementedError("VBA embedding is supported on Windows/macOS only.")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except Exception as e:  # pragma: no cover
        raise RuntimeError("Embedding VBA on Windows requires 'pywin32' (win32com).") from e

    module_name = _parse_vb_module_name_from_bas(bas)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(xlsm.as_posix())
        try:
            vbproj = wb.VBProject
        except Exception as e:  # pragma: no cover
            raise RuntimeError(
                "Failed to access VBProject. Enable Excel setting: 'Trust access to the VBA project object model'."
            ) from e

        if overwrite and module_name:
            try:
                for comp in vbproj.VBComponents:
                    if getattr(comp, "Name", None) == module_name:
                        vbproj.VBComponents.Remove(comp)
                        break
            except Exception:
                pass

        vbproj.VBComponents.Import(bas.as_posix())
        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
    clear_macos_quarantine(xlsm)
