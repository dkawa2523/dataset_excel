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


class _FakeDatasetTask:
    def __init__(self) -> None:
        self.set_script_calls: list[dict] = []
        self.set_packages_calls: list[object] = []
        self.connect_calls: list[dict] = []

    def connect_configuration(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None

    def connect(self, mutable, name=None, **kwargs):  # type: ignore[no-untyped-def]
        self.connect_calls.append({"mutable": mutable, "name": name, **kwargs})
        return mutable

    def upload_artifact(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return True

    def set_script(self, **kwargs):  # type: ignore[no-untyped-def]
        self.set_script_calls.append(dict(kwargs))

    def set_packages(self, packages):  # type: ignore[no-untyped-def]
        self.set_packages_calls.append(packages)


class _FakeLogger:
    def report_scalar(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None

    def report_table(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None


class _FakeDataset:
    def __init__(self) -> None:
        self.id = "ds_fake_register_1"
        self._task = _FakeDatasetTask()
        self._logger = _FakeLogger()

    def add_files(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None

    def upload(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None

    def finalize(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return True

    def get_logger(self):  # type: ignore[no-untyped-def]
        return self._logger


class TestCliRegisterExecutionScript(unittest.TestCase):
    def test_register_sets_task_script_when_execution_is_set(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "clearml:",
                        "  dataset_project: P",
                        "  dataset_name: N",
                        "  execution:",
                        "    repository: https://example.com/repo.git",
                        "    branch: main",
                        "    working_dir: .",
                        "    entry_point: clearml_agent_reprocess.py",
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
                    rc = main(["register", "--spec", spec_path.as_posix()])

            self.assertEqual(rc, 0)
            self.assertEqual(len(fake_dataset._task.set_script_calls), 1)
            call = fake_dataset._task.set_script_calls[0]
            self.assertEqual(call.get("repository"), "https://example.com/repo.git")
            self.assertEqual(call.get("branch"), "main")
            self.assertEqual(call.get("working_dir"), ".")
            self.assertEqual(call.get("entry_point"), "clearml_agent_reprocess.py")
            self.assertEqual(len(fake_dataset._task.set_packages_calls), 1)
            pkg_arg = fake_dataset._task.set_packages_calls[0]
            self.assertTrue(str(pkg_arg).endswith("requirements.txt"))
            self.assertEqual(len(fake_dataset._task.connect_calls), 2)
            by_name = {c.get("name"): c for c in fake_dataset._task.connect_calls}

            conn = by_name.get("dataset_format_spec")
            self.assertIsNotNone(conn)
            assert conn is not None
            self.assertIn("yaml", conn.get("mutable", {}))
            self.assertIn("schema_version: 1", conn["mutable"]["yaml"])

            info = by_name.get("clearml_dataset_excel")
            self.assertIsNotNone(info)
            assert info is not None
            self.assertEqual(info.get("mutable", {}).get("dataset_id"), "ds_fake_register_1")
            self.assertEqual(info.get("mutable", {}).get("dataset_project"), "P")
            self.assertEqual(info.get("mutable", {}).get("dataset_name"), "N")
            self.assertEqual(info.get("mutable", {}).get("output_dataset_project"), "P")
            self.assertEqual(info.get("mutable", {}).get("output_dataset_name"), "N")
            self.assertEqual(info.get("mutable", {}).get("output_uri"), "")
            self.assertEqual(info.get("mutable", {}).get("output_tags"), [])

    def test_register_sets_packages_from_requirements_txt_with_working_dir(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "clearml:",
                        "  dataset_project: P",
                        "  dataset_name: N",
                        "  execution:",
                        "    repository: https://example.com/repo.git",
                        "    branch: main",
                        "    working_dir: subdir",
                        "    entry_point: clearml_agent_reprocess.py",
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
                    rc = main(["register", "--spec", spec_path.as_posix()])

            self.assertEqual(rc, 0)
            self.assertEqual(len(fake_dataset._task.set_packages_calls), 1)
            self.assertTrue(str(fake_dataset._task.set_packages_calls[0]).endswith("requirements.txt"))


if __name__ == "__main__":
    unittest.main()
