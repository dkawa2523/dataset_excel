import sys
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.format_excel import annotate_template_with_clearml_info  # noqa: E402
from clearml_dataset_excel.vba_embedder import embed_vba_module_into_xlsm  # noqa: E402


def _add_vba_project_bin(xlsm_path: Path, data: bytes) -> None:
    import zipfile

    with zipfile.ZipFile(xlsm_path, "a") as z:
        z.writestr("xl/vbaProject.bin", data)


def _read_zip_entry(xlsm_path: Path, name: str) -> bytes:
    import zipfile

    with zipfile.ZipFile(xlsm_path, "r") as z:
        return z.read(name)


class TestFormatExcelVbaRepair(unittest.TestCase):
    def test_annotate_repairs_vba_content_types_override(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            target = base / "target.xlsm"
            donor = base / "donor.xlsm"

            import openpyxl

            wb = openpyxl.Workbook()
            wb.active.title = "Info"
            wb.save(target)

            openpyxl.Workbook().save(donor)
            donor_vba = b"dummy ClearMLDatasetExcel_Run dummy"
            _add_vba_project_bin(donor, donor_vba)

            embed_vba_module_into_xlsm(excel_path=target, template_excel=donor, overwrite=True)
            original_vba = _read_zip_entry(target, "xl/vbaProject.bin")

            annotate_template_with_clearml_info(
                target,
                dataset_project="p",
                dataset_name="n",
                dataset_id="id",
                clearml_web_url="http://example.invalid",
            )

            # Repair should not overwrite the existing VBA project.
            self.assertEqual(_read_zip_entry(target, "xl/vbaProject.bin"), original_vba)

            # vbaProject.bin must be registered as an Override (some writers drop this).
            ct_root = ET.fromstring(_read_zip_entry(target, "[Content_Types].xml"))
            ct_ns = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
            vba_ct = None
            for o in ct_root.findall("ct:Override", ct_ns):
                if o.attrib.get("PartName") == "/xl/vbaProject.bin":
                    vba_ct = o.attrib.get("ContentType")
                    break
            self.assertEqual(vba_ct, "application/vnd.ms-office.vbaProject")


if __name__ == "__main__":
    unittest.main()

