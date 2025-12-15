import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.default_vba_project import get_default_vba_project_bin  # noqa: E402
from clearml_dataset_excel.vba_embedder import embed_vba_module_into_xlsm  # noqa: E402
from clearml_dataset_excel.vba_project import vba_project_has_symbol  # noqa: E402


class TestVbaEmbedderDefaultVbaProject(unittest.TestCase):
    def test_default_vba_project_contains_macro(self) -> None:
        data = get_default_vba_project_bin()
        self.assertTrue(vba_project_has_symbol(data, "ClearMLDatasetExcel_Run"))

    def test_embed_uses_default_when_bas_and_template_omitted(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            target = base / "target.xlsm"

            import openpyxl
            import zipfile

            openpyxl.Workbook().save(target)

            embed_vba_module_into_xlsm(excel_path=target, overwrite=True)

            with zipfile.ZipFile(target, "r") as z:
                self.assertIn("xl/vbaProject.bin", set(z.namelist()))
                vba = z.read("xl/vbaProject.bin")
            self.assertTrue(vba_project_has_symbol(vba, "ClearMLDatasetExcel_Run"))


if __name__ == "__main__":
    unittest.main()

