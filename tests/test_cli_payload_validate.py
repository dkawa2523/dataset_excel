import sys
import tempfile
import unittest
from contextlib import redirect_stderr
from io import StringIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.cli import main  # noqa: E402


class TestCliPayloadValidate(unittest.TestCase):
    def test_payload_validate_ok(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "spec").mkdir()
            (root / "template").mkdir()

            (root / "spec" / "spec.yaml").write_text("schema_version: 1\n", encoding="utf-8")
            (root / "template" / "t.xlsm").write_text("dummy", encoding="utf-8")
            (root / "payload.json").write_text(
                "\n".join(
                    [
                        "{",
                        '  "payload_version": 1,',
                        '  "spec_path": "spec/spec.yaml",',
                        '  "template_excel": "template/t.xlsm"',
                        "}",
                        "",
                    ]
                ),
                encoding="utf-8",
            )

            rc = main(["payload", "validate", "--root", root.as_posix()])
            self.assertEqual(rc, 0)

    def test_payload_validate_runner_exe_ok(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "spec").mkdir()
            (root / "template").mkdir()

            (root / "spec" / "spec.yaml").write_text("schema_version: 1\n", encoding="utf-8")
            (root / "template" / "t.xlsm").write_text("dummy", encoding="utf-8")
            (root / "template" / "clearml_dataset_excel_runner.exe").write_text("dummy", encoding="utf-8")
            (root / "payload.json").write_text(
                "\n".join(
                    [
                        "{",
                        '  "payload_version": 1,',
                        '  "spec_path": "spec/spec.yaml",',
                        '  "template_excel": "template/t.xlsm",',
                        '  "runner_exe_windows": "template/clearml_dataset_excel_runner.exe"',
                        "}",
                        "",
                    ]
                ),
                encoding="utf-8",
            )

            rc = main(["payload", "validate", "--root", root.as_posix()])
            self.assertEqual(rc, 0)

    def test_payload_validate_runner_exe_missing(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "spec").mkdir()
            (root / "template").mkdir()

            (root / "spec" / "spec.yaml").write_text("schema_version: 1\n", encoding="utf-8")
            (root / "template" / "t.xlsm").write_text("dummy", encoding="utf-8")
            (root / "payload.json").write_text(
                "\n".join(
                    [
                        "{",
                        '  "payload_version": 1,',
                        '  "spec_path": "spec/spec.yaml",',
                        '  "template_excel": "template/t.xlsm",',
                        '  "runner_exe_windows": "template/clearml_dataset_excel_runner.exe"',
                        "}",
                        "",
                    ]
                ),
                encoding="utf-8",
            )

            err = StringIO()
            with redirect_stderr(err):
                rc = main(["payload", "validate", "--root", root.as_posix()])
            self.assertEqual(rc, 1)
            self.assertIn("Missing file for runner_exe_windows", err.getvalue())

    def test_payload_validate_missing_file(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "spec").mkdir()
            (root / "template").mkdir()

            (root / "spec" / "spec.yaml").write_text("schema_version: 1\n", encoding="utf-8")
            (root / "payload.json").write_text(
                "\n".join(
                    [
                        "{",
                        '  "payload_version": 1,',
                        '  "spec_path": "spec/spec.yaml",',
                        '  "template_excel": "template/missing.xlsm"',
                        "}",
                        "",
                    ]
                ),
                encoding="utf-8",
            )

            err = StringIO()
            with redirect_stderr(err):
                rc = main(["payload", "validate", "--root", root.as_posix()])
            self.assertEqual(rc, 1)
            self.assertIn("Missing file for template_excel", err.getvalue())


if __name__ == "__main__":
    unittest.main()
