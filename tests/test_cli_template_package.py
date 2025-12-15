import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.cli import main  # noqa: E402
from clearml_dataset_excel.format_spec import SpecError, load_format_spec  # noqa: E402


class TestCliTemplatePackage(unittest.TestCase):
    def test_template_package_creates_zip(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "addin:",
                        "  enabled: true",
                        "  embed_vba: true",
                        "  windows_mode: addin",
                        "  spec_filename: run.yaml",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str}",
                        "    - {name: meas_path, type: path}",
                        "files:",
                        "  - id: m",
                        "    path_column: meas_path",
                        "    mapping: {axes: {t: time}, targets: [{name: f, source: value}]}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            (base / "clearml_dataset_excel_runner.exe").write_text("dummy", encoding="utf-8")

            out_zip = base / "template_package.zip"

            # Keep tests offline; package does not require ClearML.
            with patch("clearml_dataset_excel.cli.upload_dataset"):
                rc = main(
                    [
                        "template",
                        "package",
                        "--spec",
                        spec_path.as_posix(),
                        "--output",
                        out_zip.as_posix(),
                        "--overwrite",
                    ]
                )
            self.assertEqual(rc, 0)
            self.assertTrue(out_zip.exists())

            with zipfile.ZipFile(out_zip, "r") as z:
                names = set(z.namelist())

            self.assertIn("README.txt", names)
            self.assertIn("mac/condition_template.xlsm", names)
            self.assertIn("mac/run.yaml", names)
            self.assertIn("mac/clearml_dataset_excel_addin.bas", names)
            self.assertIn("windows/condition_template.xlsx", names)
            self.assertIn("windows/clearml_dataset_excel_addin.xlam", names)
            self.assertIn("windows/run.yaml", names)
            self.assertIn("windows/clearml_dataset_excel_runner.exe", names)

    def test_addin_spec_filename_must_be_basename(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "addin:",
                        "  enabled: true",
                        "  spec_filename: subdir/run.yaml",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str}",
                        "files:",
                        "  - id: m",
                        "    path_column: id",
                        "    mapping: {axes: {t: time}, targets: [{name: f, source: value}]}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            with self.assertRaises(SpecError):
                load_format_spec(spec_path)


if __name__ == "__main__":
    unittest.main()

