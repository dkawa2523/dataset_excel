import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.format_clearml import upload_dataset  # noqa: E402
from clearml_dataset_excel.format_spec import load_format_spec  # noqa: E402


class _FakeDataset:
    def __init__(self) -> None:
        self.id = "ds_fake"
        self._task = None
        self._logger = None

    def add_files(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None

    def upload(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return None

    def finalize(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return True

    def get_logger(self):  # type: ignore[no-untyped-def]
        class _L:
            def report_scalar(self, *a, **k):  # type: ignore[no-untyped-def]
                return None

            def report_table(self, *a, **k):  # type: ignore[no-untyped-def]
                return None

            def report_image(self, *a, **k):  # type: ignore[no-untyped-def]
                return None

            def report_histogram(self, *a, **k):  # type: ignore[no-untyped-def]
                return None

        if self._logger is None:
            self._logger = _L()
        return self._logger


class TestUploadDataset(unittest.TestCase):
    def _make_stage(self, base: Path) -> tuple[Path, Path, object]:
        stage = base / "stage"
        (stage / "spec").mkdir(parents=True, exist_ok=True)
        (stage / "template").mkdir(parents=True, exist_ok=True)

        spec_path = stage / "spec" / "spec.yaml"
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
        template_path = stage / "template" / "condition_template.xlsm"
        template_path.write_text("dummy", encoding="utf-8")
        (stage / "payload.json").write_text(
            json.dumps(
                {"payload_version": 1, "spec_path": "spec/spec.yaml", "template_excel": "template/condition_template.xlsm"},
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )
        spec = load_format_spec(spec_path)
        return stage, spec_path, spec

    def test_upload_dataset_creates_file_output_uri_dir(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            stage, spec_path, spec = self._make_stage(base)
            out_dir = base / "output_not_exist"
            self.assertFalse(out_dir.exists())
            out_uri = "file://" + out_dir.as_posix()

            fake_dataset = _FakeDataset()

            with patch("clearml.Dataset") as dataset_cls:
                dataset_cls.get.side_effect = Exception("not found")

                def _create(**kwargs):  # type: ignore[no-untyped-def]
                    # Directory should already exist when ClearML checks permissions.
                    self.assertTrue(out_dir.exists())
                    self.assertEqual(kwargs.get("output_uri"), out_uri)
                    return fake_dataset

                dataset_cls.create.side_effect = _create
                ds_id = upload_dataset(
                    stage_dir=stage,
                    spec=spec,
                    dataset_project="P",
                    dataset_name="N",
                    output_uri=out_uri,
                )

            self.assertEqual(ds_id, fake_dataset.id)
            self.assertTrue(out_dir.exists())

    def test_upload_dataset_recovers_from_unfinalized_parent(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            stage, spec_path, spec = self._make_stage(base)

            fake_dataset = _FakeDataset()

            def _get(*args, **kwargs):  # type: ignore[no-untyped-def]
                if kwargs.get("writable_copy", False):
                    raise ValueError("Cannot inherit from a parent that was not finalized/closed")
                return fake_dataset

            with patch("clearml.Dataset") as dataset_cls:
                dataset_cls.get.side_effect = _get
                dataset_cls.create.side_effect = AssertionError("Dataset.create should not be called")
                ds_id = upload_dataset(
                    stage_dir=stage,
                    spec=spec,
                    dataset_project="P",
                    dataset_name="N",
                    output_uri=None,
                )

            self.assertEqual(ds_id, fake_dataset.id)


if __name__ == "__main__":
    unittest.main()

