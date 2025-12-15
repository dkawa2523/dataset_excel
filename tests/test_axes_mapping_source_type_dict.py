import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.cli import main  # noqa: E402
from clearml_dataset_excel.format_spec import load_format_spec, spec_to_yaml_dict  # noqa: E402


class TestAxesMappingSourceTypeDict(unittest.TestCase):
    def test_axes_mapping_allows_source_and_type_object(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            data = base / "data"
            data.mkdir()
            (data / "meas.csv").write_text("time,value\n0,1\n1,2\n", encoding="utf-8")

            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str}",
                        "    - {name: meas_path, type: path}",
                        "files:",
                        "  - id: m",
                        "    path_column: meas_path",
                        "    mapping:",
                        "      axes:",
                        "        t: {source: time, type: float}",
                        "      targets:",
                        "        - {name: f, source: value, type: float}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            spec = load_format_spec(spec_path)
            self.assertEqual(spec.files[0].axes.t, "time")
            self.assertEqual(spec.files[0].axis_types.t, "float")

            raw = spec_to_yaml_dict(spec)
            t_axis = raw["files"][0]["mapping"]["axes"]["t"]
            self.assertIsInstance(t_axis, dict)
            self.assertEqual(t_axis.get("source"), "time")
            self.assertEqual(t_axis.get("type"), "float")

            import openpyxl

            xlsm = base / "conditions.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Conditions"
            ws.append(["id", "meas_path"])
            ws.append(["s1", "data/meas.csv"])
            wb.save(xlsm)

            buf = StringIO()
            with redirect_stdout(buf):
                rc = main(
                    [
                        "run",
                        "--spec",
                        spec_path.as_posix(),
                        "--excel",
                        xlsm.as_posix(),
                        "--output-root",
                        base.as_posix(),
                        "--no-upload",
                    ]
                )
            self.assertEqual(rc, 0)


if __name__ == "__main__":
    unittest.main()

