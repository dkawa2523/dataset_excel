import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.format_clearml import stage_dataset_payload  # noqa: E402
from clearml_dataset_excel.format_excel import generate_condition_template  # noqa: E402
from clearml_dataset_excel.format_processor import process_condition_excel  # noqa: E402
from clearml_dataset_excel.format_spec import load_format_spec  # noqa: E402


class TestStageDatasetPayloadPathMap(unittest.TestCase):
    def test_stage_dataset_payload_writes_path_map_for_external_files(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            base = root / "base"
            base.mkdir()

            meas = root / "meas.csv"
            meas.write_text("time,value\n0,1\n", encoding="utf-8")

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
                        "    mapping: {axes: {t: time}, targets: [{name: f, source: value}]}",
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
            ws.append(["id", "meas_path"])
            ws.append(["s1", meas.as_posix()])
            wb.save(xlsm)

            outputs = process_condition_excel(spec, xlsm, output_root=base)
            template_path = base / spec.template.template_filename
            generate_condition_template(spec, template_path, overwrite=True)

            stage_td = stage_dataset_payload(
                spec_path=spec_path,
                spec=spec,
                condition_excel=xlsm,
                outputs=outputs,
                template_excel=template_path,
                vba_module=None,
            )
            try:
                stage_root = Path(stage_td.name).resolve()
                meta = (stage_root / "payload.json").read_text(encoding="utf-8")
                import json

                payload = json.loads(meta)
                self.assertIn("path_map", payload)
                self.assertEqual(payload["path_map"][meas.as_posix()], "input/external/001_meas.csv")
            finally:
                stage_td.cleanup()


if __name__ == "__main__":
    unittest.main()

