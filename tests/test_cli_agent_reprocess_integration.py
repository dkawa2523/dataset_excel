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


class _FakeTask:
    def __init__(self, cfg: dict) -> None:
        self.id = "ds_base_1"
        self._cfg = cfg

    def get_configuration_object_as_dict(self, name: str):  # type: ignore[no-untyped-def]
        if name == "dataset_format_spec":
            return self._cfg
        return {}


class _FakeLogger:
    def report_scalar(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None

    def report_table(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None

    def report_image(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None


class _FakeDatasetTask:
    def __init__(self) -> None:
        self.connect_configuration_calls: list[dict] = []

    def connect_configuration(self, configuration, name=None, **kwargs):  # type: ignore[no-untyped-def]
        self.connect_configuration_calls.append({"configuration": configuration, "name": name, **kwargs})
        return configuration

    def upload_artifact(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return True


class _FakeBaseDataset:
    def __init__(self, local_copy: Path) -> None:
        self._local_copy = local_copy
        self.project = "P_cloned"
        self.name = "N_cloned"

    def get_local_copy(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return self._local_copy.as_posix()


class _FakeUploadDataset:
    def __init__(self) -> None:
        self.id = "ds_new_2"
        self._task = _FakeDatasetTask()
        self._logger = _FakeLogger()
        self.add_files_calls: list[dict] = []
        self.upload_calls: list[dict] = []
        self.finalize_calls: list[dict] = []

    def add_files(self, path: str, **kwargs):  # type: ignore[no-untyped-def]
        self.add_files_calls.append({"path": path, **kwargs})

    def upload(self, **kwargs):  # type: ignore[no-untyped-def]
        self.upload_calls.append(dict(kwargs))

    def finalize(self, **kwargs):  # type: ignore[no-untyped-def]
        self.finalize_calls.append(dict(kwargs))
        return True

    def get_logger(self):  # type: ignore[no-untyped-def]
        return self._logger


class TestCliAgentReprocessIntegration(unittest.TestCase):
    def test_agent_reprocess_with_mock_clearml(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            input_dir = root / "input"
            (input_dir / "external").mkdir(parents=True)

            meas1 = input_dir / "external" / "000_meas.csv"
            meas2 = input_dir / "external" / "001_meas.csv"
            meas1.write_text("time,value\n0,1\n1,2\n", encoding="utf-8")
            meas2.write_text("time,value\n0,10\n1,20\n", encoding="utf-8")

            import openpyxl

            xlsm = input_dir / "conditions.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Conditions"
            ws.append(["id", "meas_path"])
            ws.append(["s1", "/abs1/meas.csv"])
            ws.append(["s2", "/abs2/meas.csv"])
            wb.save(xlsm)

            (root / "payload.json").write_text(
                '{"condition_excel":"input/conditions.xlsm","path_map":{"/abs1/meas.csv":"input/external/000_meas.csv","/abs2/meas.csv":"input/external/001_meas.csv"}}\n',
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

            fake_task = _FakeTask(cfg)
            base_ds = _FakeBaseDataset(root)
            upload_ds = _FakeUploadDataset()

            def _dataset_get_side_effect(**kwargs):  # type: ignore[no-untyped-def]
                if kwargs.get("writable_copy"):
                    return upload_ds
                return base_ds

            with patch("clearml.Task") as task_cls, patch("clearml.Dataset") as dataset_cls:
                task_cls.current_task.return_value = fake_task
                dataset_cls.get.side_effect = _dataset_get_side_effect

                buf = StringIO()
                with redirect_stdout(buf):
                    rc = main(["agent", "reprocess"])

            self.assertEqual(rc, 0)
            self.assertIn(upload_ds.id, buf.getvalue())
            self.assertEqual(len(upload_ds.add_files_calls), 1)
            self.assertEqual(len(upload_ds.upload_calls), 1)
            self.assertEqual(len(upload_ds.finalize_calls), 1)
            self.assertEqual(len(upload_ds._task.connect_configuration_calls), 1)
            cfg_saved = upload_ds._task.connect_configuration_calls[0]["configuration"]
            self.assertIn("condition", cfg_saved)
            self.assertIn("files", cfg_saved)


if __name__ == "__main__":
    unittest.main()
