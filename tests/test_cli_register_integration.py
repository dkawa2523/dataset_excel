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


class _FakeLogger:
    def __init__(self) -> None:
        self.report_scalar_calls: list[dict] = []
        self.report_table_calls: list[dict] = []
        self.report_image_calls: list[dict] = []

    def report_scalar(self, title: str, series: str, value: float, iteration: int):  # type: ignore[no-untyped-def]
        self.report_scalar_calls.append(
            {"title": title, "series": series, "value": value, "iteration": iteration}
        )

    def report_table(self, title: str, series: str, iteration=None, table_plot=None, **kwargs):  # type: ignore[no-untyped-def]
        self.report_table_calls.append(
            {"title": title, "series": series, "iteration": iteration, "table_plot": table_plot}
        )

    def report_image(self, title: str, series: str, iteration=None, local_path=None, **kwargs):  # type: ignore[no-untyped-def]
        self.report_image_calls.append(
            {"title": title, "series": series, "iteration": iteration, "local_path": local_path}
        )


class _FakeDatasetTask:
    def __init__(self) -> None:
        self.connect_configuration_calls: list[dict] = []
        self.upload_artifact_calls: list[dict] = []

    def connect_configuration(self, configuration, name=None, **kwargs):  # type: ignore[no-untyped-def]
        self.connect_configuration_calls.append({"configuration": configuration, "name": name, **kwargs})
        return configuration

    def upload_artifact(self, name: str, artifact_object, **kwargs):  # type: ignore[no-untyped-def]
        self.upload_artifact_calls.append({"name": name, "artifact_object": artifact_object, **kwargs})
        return True


class _FakeDataset:
    def __init__(self) -> None:
        self.id = "ds_fake_register_1"
        self.add_files_calls: list[dict] = []
        self.upload_calls: list[dict] = []
        self.finalize_calls: list[dict] = []
        self._task = _FakeDatasetTask()
        self._logger = _FakeLogger()

    def add_files(self, path: str, **kwargs):  # type: ignore[no-untyped-def]
        self.add_files_calls.append({"path": path, **kwargs})

    def upload(self, **kwargs):  # type: ignore[no-untyped-def]
        self.upload_calls.append(dict(kwargs))

    def finalize(self, **kwargs):  # type: ignore[no-untyped-def]
        self.finalize_calls.append(dict(kwargs))
        return True

    def get_logger(self):  # type: ignore[no-untyped-def]
        return self._logger


class TestCliRegisterIntegration(unittest.TestCase):
    def test_register_upload_with_mock_clearml(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "addin:",
                        "  enabled: true",
                        "  target_os: auto",
                        "  spec_filename: spec.yaml",
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

            fake_dataset = _FakeDataset()
            with patch("clearml.Dataset") as dataset_cls:
                dataset_cls.get.side_effect = Exception("not found")
                dataset_cls.create.return_value = fake_dataset

                buf = StringIO()
                with redirect_stdout(buf):
                    rc = main(
                        [
                            "register",
                            "--spec",
                            spec_path.as_posix(),
                            "--dataset-project",
                            "P",
                            "--dataset-name",
                            "N",
                        ]
                    )

            self.assertEqual(rc, 0)
            self.assertIn(fake_dataset.id, buf.getvalue())
            self.assertEqual(len(fake_dataset.add_files_calls), 1)
            self.assertEqual(len(fake_dataset.upload_calls), 1)
            self.assertEqual(len(fake_dataset.finalize_calls), 1)
            self.assertEqual(len(fake_dataset._task.connect_configuration_calls), 1)
            cfg = fake_dataset._task.connect_configuration_calls[0]["configuration"]
            self.assertIn("addin", cfg)
            self.assertIn("condition", cfg)
            self.assertNotIn("condition_columns", cfg)


if __name__ == "__main__":
    unittest.main()
