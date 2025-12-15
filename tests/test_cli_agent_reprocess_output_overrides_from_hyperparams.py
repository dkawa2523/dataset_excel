import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.cli import main  # noqa: E402


class _FakeTaskWithOutputOverrides:
    def __init__(self, cfg: dict) -> None:
        self.id = "task_clone_123"
        self._cfg = cfg

    def get_configuration_object_as_dict(self, name: str):  # type: ignore[no-untyped-def]
        if name == "dataset_format_spec":
            return self._cfg
        return {}

    def get_project_name(self):  # type: ignore[no-untyped-def]
        return "P_clone"

    def get_parameter(self, key: str):  # type: ignore[no-untyped-def]
        if key == "clearml_dataset_excel/dataset_id":
            return "ds_base_1"
        if key == "clearml_dataset_excel/output_dataset_project":
            return "P_out"
        if key == "clearml_dataset_excel/output_dataset_name":
            return "N_out"
        return None


class _FakeBaseDataset:
    def __init__(self, local_copy: Path) -> None:
        self._local_copy = local_copy
        self.project = "P_base"
        self.name = "N_base"

    def get_local_copy(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return self._local_copy.as_posix()


class TestCliAgentReprocessOutputOverrides(unittest.TestCase):
    def test_agent_reprocess_output_project_name_overrides(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            input_dir = root / "input"
            (input_dir / "external").mkdir(parents=True)

            meas = input_dir / "external" / "000_meas.csv"
            meas.write_text("time,value\n0,1\n1,2\n", encoding="utf-8")

            import openpyxl

            xlsm = input_dir / "conditions.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Conditions"
            ws.append(["id", "meas_path"])
            ws.append(["s1", "/abs/meas.csv"])
            wb.save(xlsm)

            (root / "payload.json").write_text(
                '{"payload_version":1,"condition_excel":"input/conditions.xlsm","path_map":{"/abs/meas.csv":"input/external/000_meas.csv"}}\n',
                encoding="utf-8",
            )

            cfg = {
                "schema_version": 1,
                "condition": {"columns": [{"name": "id", "type": "str"}, {"name": "meas_path", "type": "path"}]},
                "files": [
                    {
                        "id": "m",
                        "path_column": "meas_path",
                        "format": "csv",
                        "mapping": {"axes": {"t": "time"}, "targets": [{"name": "f", "source": "value"}]},
                    }
                ],
            }

            fake_task = _FakeTaskWithOutputOverrides(cfg)
            base_ds = _FakeBaseDataset(root)

            def _dataset_get_side_effect(**kwargs):  # type: ignore[no-untyped-def]
                if kwargs.get("dataset_id") == "ds_base_1":
                    return base_ds
                raise AssertionError(f"Unexpected Dataset.get call: {kwargs}")

            called: dict[str, object] = {}

            def _fake_upload_dataset(*, dataset_project, dataset_name, base_dataset_id=None, spec=None, **kwargs):  # type: ignore[no-untyped-def]
                called["dataset_project"] = dataset_project
                called["dataset_name"] = dataset_name
                called["base_dataset_id"] = base_dataset_id
                if spec is not None and getattr(spec, "clearml", None) is not None:
                    called["spec_clearml_dataset_project"] = spec.clearml.dataset_project
                    called["spec_clearml_dataset_name"] = spec.clearml.dataset_name
                return "ds_new"

            with patch("clearml.Task") as task_cls, patch("clearml.Dataset") as dataset_cls, patch(
                "clearml_dataset_excel.agent.upload_dataset", side_effect=_fake_upload_dataset
            ):
                task_cls.current_task.return_value = fake_task
                dataset_cls.get.side_effect = _dataset_get_side_effect

                buf = StringIO()
                with redirect_stdout(buf):
                    rc = main(["agent", "reprocess"])

            self.assertEqual(rc, 0)
            self.assertIn("ds_new", buf.getvalue())
            self.assertEqual(called.get("dataset_project"), "P_out")
            self.assertEqual(called.get("dataset_name"), "N_out")
            self.assertIsNone(called.get("base_dataset_id"))
            self.assertEqual(called.get("spec_clearml_dataset_project"), "P_out")
            self.assertEqual(called.get("spec_clearml_dataset_name"), "N_out")


if __name__ == "__main__":
    unittest.main()

