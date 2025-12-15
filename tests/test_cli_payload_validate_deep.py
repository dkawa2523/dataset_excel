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
from clearml_dataset_excel.format_excel import generate_condition_template  # noqa: E402
from clearml_dataset_excel.format_spec import load_format_spec  # noqa: E402


class TestCliPayloadValidateDeep(unittest.TestCase):
    def test_payload_validate_deep_ok(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "spec").mkdir()
            (root / "template").mkdir()
            (root / "input" / "external").mkdir(parents=True)

            spec_path = root / "spec" / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str, required: true}",
                        "    - {name: meas_path, type: path, required: true}",
                        "files:",
                        "  - id: m",
                        "    path_column: meas_path",
                        "    format: csv",
                        "    mapping: {axes: {t: time}, targets: [{name: f, source: value, type: float}]}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            spec = load_format_spec(spec_path)

            tmpl = root / "template" / "condition_template.xlsm"
            generate_condition_template(spec, tmpl, overwrite=True)

            meas_rel = "input/external/000_meas.csv"
            (root / meas_rel).write_text("time,value\n0,1\n1,2\n", encoding="utf-8")

            import openpyxl

            cond_excel_rel = "input/conditions.xlsm"
            cond_excel = root / cond_excel_rel
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Conditions"
            ws.append(["id", "meas_path"])
            ws.append(["s1", "/abs/meas.csv"])
            wb.save(cond_excel)

            (root / "payload.json").write_text(
                "\n".join(
                    [
                        "{",
                        '  "payload_version": 1,',
                        '  "spec_path": "spec/spec.yaml",',
                        '  "template_excel": "template/condition_template.xlsm",',
                        f'  "condition_excel": "{cond_excel_rel}",',
                        '  "path_map": {',
                        f'    "/abs/meas.csv": "{meas_rel}"',
                        "  }",
                        "}",
                        "",
                    ]
                ),
                encoding="utf-8",
            )

            rc = main(["payload", "validate", "--root", root.as_posix(), "--deep"])
            self.assertEqual(rc, 0)

    def test_payload_validate_deep_fails_on_bad_measurement(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "spec").mkdir()
            (root / "template").mkdir()
            (root / "input" / "external").mkdir(parents=True)

            spec_path = root / "spec" / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str, required: true}",
                        "    - {name: meas_path, type: path, required: true}",
                        "files:",
                        "  - id: m",
                        "    path_column: meas_path",
                        "    format: csv",
                        "    mapping: {axes: {t: time}, targets: [{name: f, source: value, type: float}]}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            spec = load_format_spec(spec_path)

            tmpl = root / "template" / "condition_template.xlsm"
            generate_condition_template(spec, tmpl, overwrite=True)

            meas_rel = "input/external/000_meas.csv"
            # Missing "value" column
            (root / meas_rel).write_text("time,other\n0,1\n", encoding="utf-8")

            import openpyxl

            cond_excel_rel = "input/conditions.xlsm"
            cond_excel = root / cond_excel_rel
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Conditions"
            ws.append(["id", "meas_path"])
            ws.append(["s1", "/abs/meas.csv"])
            wb.save(cond_excel)

            (root / "payload.json").write_text(
                "\n".join(
                    [
                        "{",
                        '  "payload_version": 1,',
                        '  "spec_path": "spec/spec.yaml",',
                        '  "template_excel": "template/condition_template.xlsm",',
                        f'  "condition_excel": "{cond_excel_rel}",',
                        '  "path_map": {',
                        f'    "/abs/meas.csv": "{meas_rel}"',
                        "  }",
                        "}",
                        "",
                    ]
                ),
                encoding="utf-8",
            )

            err = StringIO()
            with redirect_stderr(err):
                rc = main(["payload", "validate", "--root", root.as_posix(), "--deep"])
            self.assertEqual(rc, 1)
            self.assertIn("Deep validation failed", err.getvalue())


if __name__ == "__main__":
    unittest.main()

