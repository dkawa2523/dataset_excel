import base64
import io
import sys
import tempfile
import unittest
import zlib
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

import olefile  # noqa: E402

from clearml_dataset_excel.default_vba_project import (  # noqa: E402
    _DEFAULT_VBA_PROJECT_BIN_ZLIB_B64,
    get_default_vba_project_bin,
)
from clearml_dataset_excel.msovba import decompress_stream  # noqa: E402
from clearml_dataset_excel.vba_embedder import embed_vba_module_into_xlsm  # noqa: E402


def _extract_module_source(vba_project_bin: bytes, module_name: str) -> str:
    ole = olefile.OleFileIO(io.BytesIO(vba_project_bin))
    try:
        stream = ole.openstream(["VBA", module_name]).read()
    finally:
        ole.close()

    idx = stream.lower().find(b"\x00attribut")
    if idx < 0:
        raise AssertionError(f"attribute marker not found in module stream: {module_name}")
    start = idx - 3
    if start < 0 or stream[start] != 0x01:
        raise AssertionError(f"VBA signature not found at expected offset: {module_name}")
    src = decompress_stream(stream[start:])
    return src.decode("cp1252", errors="replace")


class TestDefaultVbaProjectPatch(unittest.TestCase):
    def test_get_default_vba_project_bin_is_patched(self) -> None:
        vba = get_default_vba_project_bin()
        self.assertGreater(len(vba), 0)

        # The donor VBA project used to mark modules as "private", which makes Excel's macro list empty.
        # The patched default removes ModulePrivate records from the VBA dir stream.
        ole = olefile.OleFileIO(io.BytesIO(vba))
        try:
            dir_comp = ole.openstream(["VBA", "dir"]).read()
        finally:
            ole.close()
        dir_raw = decompress_stream(dir_comp)
        self.assertNotIn(b"\x28\x00\x00\x00\x00\x00\x2B\x00\x00\x00\x00\x00", dir_raw)

        sheet1 = _extract_module_source(vba, "Sheet1")
        self.assertNotIn("VB_Control", sheet1)

        calc = _extract_module_source(vba, "Calculations")
        self.assertIn("ClearMLDatasetExcel_Run", calc)
        self.assertIn("ClearMLDatasetExcel_Run_Ribbon", calc)
        self.assertNotIn("MacScript", calc)
        self.assertIn("/bin/zsh -lc", calc)
        self.assertIn("cmd.exe /c", calc)
        self.assertIn("cd /d", calc)
        self.assertIn("2>&1", calc)

    def test_embed_repairs_existing_donor_vba_project(self) -> None:
        raw = zlib.decompress(base64.b64decode(_DEFAULT_VBA_PROJECT_BIN_ZLIB_B64))
        self.assertIn("VB_Control", _extract_module_source(raw, "Sheet1"))
        self.assertIn("MacScript", _extract_module_source(raw, "Calculations"))

        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            xlsm = root / "t.xlsm"

            import openpyxl

            wb = openpyxl.Workbook()
            wb.active.title = "Conditions"
            wb.save(xlsm)

            with zipfile.ZipFile(xlsm, "a") as z:
                z.writestr("xl/vbaProject.bin", raw)

            embed_vba_module_into_xlsm(
                excel_path=xlsm,
                bas_path=None,
                overwrite=False,
                template_excel=None,
            )

            with zipfile.ZipFile(xlsm, "r") as z:
                repaired = z.read("xl/vbaProject.bin")

            self.assertNotIn("VB_Control", _extract_module_source(repaired, "Sheet1"))
            self.assertIn("ClearMLDatasetExcel_Run", _extract_module_source(repaired, "Calculations"))
            self.assertNotIn("MacScript", _extract_module_source(repaired, "Calculations"))
            self.assertIn("/bin/zsh -lc", _extract_module_source(repaired, "Calculations"))
            self.assertIn("cmd.exe /c", _extract_module_source(repaired, "Calculations"))

            ole = olefile.OleFileIO(io.BytesIO(repaired))
            try:
                dir_comp = ole.openstream(["VBA", "dir"]).read()
            finally:
                ole.close()
            dir_raw = decompress_stream(dir_comp)
            self.assertNotIn(b"\x28\x00\x00\x00\x00\x00\x2B\x00\x00\x00\x00\x00", dir_raw)


if __name__ == "__main__":
    unittest.main()
