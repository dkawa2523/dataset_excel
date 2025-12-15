import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.format_excel import generate_windows_addin_xlam  # noqa: E402
from clearml_dataset_excel.vba_project import vba_project_has_symbol  # noqa: E402


class TestFormatExcelGenerateXlam(unittest.TestCase):
    def test_generate_windows_addin_xlam_has_addin_workbook_content_type(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "addin.xlam"
            generate_windows_addin_xlam(path, overwrite=True)

            with zipfile.ZipFile(path, "r") as z:
                ct_xml = z.read("[Content_Types].xml")
                rels_xml = z.read("_rels/.rels")
                ui_xml = z.read("customUI/customUI.xml")
                vba = z.read("xl/vbaProject.bin")

            self.assertTrue(vba_project_has_symbol(vba, "ClearMLDatasetExcel_Run"))
            self.assertTrue(vba_project_has_symbol(vba, "ClearMLDatasetExcel_Run_Ribbon"))

            root = ET.fromstring(ct_xml)
            ns = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
            workbook_ct = None
            for elem in root.findall("ct:Override", ns):
                if elem.attrib.get("PartName") == "/xl/workbook.xml":
                    workbook_ct = elem.attrib.get("ContentType")
                    break

            self.assertEqual(workbook_ct, "application/vnd.ms-excel.addin.macroEnabled.main+xml")

            ui_ct = None
            for elem in root.findall("ct:Override", ns):
                if elem.attrib.get("PartName") == "/customUI/customUI.xml":
                    ui_ct = elem.attrib.get("ContentType")
                    break
            self.assertEqual(ui_ct, "application/vnd.ms-office.customUI+xml")

            rel_root = ET.fromstring(rels_xml)
            rns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
            targets = [
                rel.attrib.get("Target")
                for rel in rel_root.findall("r:Relationship", rns)
                if rel.attrib.get("Type") == "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
            ]
            self.assertIn("customUI/customUI.xml", targets)

            ui_root = ET.fromstring(ui_xml)
            on_action = None
            for elem in ui_root.iter():
                if str(elem.tag).endswith("button"):
                    on_action = elem.attrib.get("onAction")
                    break
            self.assertEqual(on_action, "ClearMLDatasetExcel_Run_Ribbon")


if __name__ == "__main__":
    unittest.main()
