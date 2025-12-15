import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.format_processor import ProcessingError, process_condition_excel  # noqa: E402
from clearml_dataset_excel.format_spec import load_format_spec  # noqa: E402


class TestFormatProcessor(unittest.TestCase):
    def test_process_condition_excel_merges_and_aggregates(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            data = base / "data"
            data.mkdir()

            meas1 = data / "meas1.csv"
            meas2 = data / "meas2.csv"
            meas1.write_text("time,value\n0,1\n1,3\n2,5\n", encoding="utf-8")
            meas2.write_text("time,value\n0,10\n1,30\n2,50\n", encoding="utf-8")

            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str, required: true}",
                        "    - {name: meas1_path, type: path}",
                        "    - {name: meas2_path, type: path}",
                        "files:",
                        "  - id: m1",
                        "    path_column: meas1_path",
                        "    format: csv",
                        "    mapping:",
                        "      axes: {t: time}",
                        "      targets:",
                        "        - {name: f1, source: value, type: float}",
                        "      aggregates:",
                        "        - {name: f1_mean, source: f1, op: mean, output_column: f1_mean}",
                        "        - {name: f1_int, source: f1, op: trapz, wrt: t, output_column: f1_int}",
                        "  - id: m2",
                        "    path_column: meas2_path",
                        "    format: csv",
                        "    mapping:",
                        "      axes: {t: time}",
                        "      targets:",
                        "        - {name: f2, source: value, type: float}",
                        "      aggregates:",
                        "        - {name: f2_mean, source: f2, op: mean, output_column: f2_mean}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            spec = load_format_spec(spec_path)

            import openpyxl

            xlsm = base / "conditions.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Conditions"
            ws.append(["id", "meas1_path", "meas2_path"])
            ws.append(["s1", meas1.as_posix(), meas2.as_posix()])
            wb.save(xlsm)

            out = process_condition_excel(spec, xlsm, output_root=base)
            self.assertTrue(out.canonical_csv.exists())
            self.assertTrue(out.conditions_csv.exists())
            self.assertTrue(out.consolidated_excel.exists())
            uploaded = {p.resolve() for p in out.uploaded_files}
            self.assertIn(xlsm.resolve(), uploaded)
            self.assertIn(meas1.resolve(), uploaded)
            self.assertIn(meas2.resolve(), uploaded)

            import pandas as pd

            cond_out = pd.read_csv(out.conditions_csv)
            self.assertIn("f1_mean", cond_out.columns)
            self.assertIn("f1_int", cond_out.columns)
            self.assertIn("f2_mean", cond_out.columns)
            self.assertEqual(cond_out.loc[0, "id"], "s1")

            canon_out = pd.read_csv(out.canonical_csv)
            self.assertIn("t", canon_out.columns)
            self.assertIn("f1", canon_out.columns)
            self.assertIn("f2", canon_out.columns)
            # merged by t => 3 rows
            self.assertEqual(len(canon_out), 3)

    def test_process_condition_excel_axes_mismatch_errors(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            data = base / "data"
            data.mkdir()
            meas1 = data / "meas1.csv"
            meas2 = data / "meas2.csv"
            meas1.write_text("time,value\n0,1\n", encoding="utf-8")
            meas2.write_text("x,value\n0,1\n", encoding="utf-8")

            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str}",
                        "    - {name: p1, type: path}",
                        "    - {name: p2, type: path}",
                        "files:",
                        "  - id: a",
                        "    path_column: p1",
                        "    mapping: {axes: {t: time}, targets: [{name: f1, source: value}]}",
                        "  - id: b",
                        "    path_column: p2",
                        "    mapping: {axes: {x: x}, targets: [{name: f2, source: value}]}",
                        "output:",
                        "  combine_mode: merge",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            spec = load_format_spec(spec_path)

            import openpyxl

            xlsm = base / "conditions.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Conditions"
            ws.append(["id", "p1", "p2"])
            ws.append(["s1", meas1.as_posix(), meas2.as_posix()])
            wb.save(xlsm)

            with self.assertRaises(ProcessingError):
                process_condition_excel(spec, xlsm, output_root=base)


if __name__ == "__main__":
    unittest.main()
