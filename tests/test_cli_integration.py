import sys
import tempfile
import unittest
from contextlib import redirect_stderr
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.cli import main  # noqa: E402


class _FakeDataset:
    def __init__(self) -> None:
        self.id = "ds_fake_123"
        self.add_files_calls: list[dict] = []
        self.add_external_files_calls: list[dict] = []
        self.remove_files_calls: list[dict] = []
        self.upload_calls: list[dict] = []
        self.finalize_calls: list[dict] = []

    def add_files(self, path: str, **kwargs):  # type: ignore[no-untyped-def]
        self.add_files_calls.append({"path": path, **kwargs})

    def add_external_files(self, source_url: str, **kwargs):  # type: ignore[no-untyped-def]
        self.add_external_files_calls.append({"source_url": source_url, **kwargs})

    def remove_files(self, **kwargs):  # type: ignore[no-untyped-def]
        self.remove_files_calls.append(dict(kwargs))

    def upload(self, **kwargs):  # type: ignore[no-untyped-def]
        self.upload_calls.append(dict(kwargs))

    def finalize(self, **kwargs):  # type: ignore[no-untyped-def]
        self.finalize_calls.append(dict(kwargs))
        return True


class TestCliIntegration(unittest.TestCase):
    def test_run_with_mock_clearml(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            data_dir = base / "data"
            data_dir.mkdir()
            (data_dir / "a.txt").write_text("a", encoding="utf-8")
            (data_dir / "b.txt").write_text("b", encoding="utf-8")
            (data_dir / "c.jpg").write_text("c", encoding="utf-8")

            manifest = base / "manifest.csv"
            manifest.write_text(
                "path,dataset_path\n"
                f"{data_dir.as_posix()},ds\n"
                f"{(data_dir / 'c.jpg').as_posix()},images\n"
                "s3://bucket/path/to/file.txt,ext\n",
                encoding="utf-8",
            )

            fake_dataset = _FakeDataset()
            with patch("clearml.Dataset") as dataset_cls:
                dataset_cls.create.return_value = fake_dataset
                buf = StringIO()
                with redirect_stdout(buf):
                    rc = main(
                        [
                            "--manifest",
                            manifest.as_posix(),
                            "--dataset-project",
                            "P",
                            "--dataset-name",
                            "N",
                            "--include",
                            "*.txt",
                            "--exclude",
                            "*b.txt",
                            "--collision-policy",
                            "ignore",
                            "--no-auto-upload",
                        ]
                    )

            self.assertEqual(rc, 0)
            self.assertIn(fake_dataset.id, buf.getvalue())

            create_kwargs = dataset_cls.create.call_args.kwargs
            self.assertEqual(create_kwargs["dataset_project"], "P")
            self.assertEqual(create_kwargs["dataset_name"], "N")

            self.assertGreaterEqual(len(fake_dataset.add_files_calls), 2)  # manifest + directory

            resolved_data_dir = data_dir.resolve().as_posix()
            dir_calls = [c for c in fake_dataset.add_files_calls if c["path"] == resolved_data_dir]
            self.assertEqual(len(dir_calls), 1)
            self.assertEqual(dir_calls[0]["wildcard"], ["*.txt"])
            self.assertEqual(dir_calls[0]["dataset_path"], "ds")

            resolved_c_jpg = (data_dir / "c.jpg").resolve().as_posix()
            file_calls = [c for c in fake_dataset.add_files_calls if c["path"] == resolved_c_jpg]
            self.assertEqual(file_calls, [])  # filtered out by --include '*.txt'

            self.assertEqual(len(fake_dataset.add_external_files_calls), 1)
            self.assertEqual(fake_dataset.add_external_files_calls[0]["source_url"], "s3://bucket/path/to/file.txt")
            self.assertEqual(fake_dataset.add_external_files_calls[0]["dataset_path"], "ext")

            self.assertEqual(len(fake_dataset.remove_files_calls), 1)
            self.assertEqual(fake_dataset.remove_files_calls[0]["dataset_path"], "*b.txt")

            self.assertEqual(len(fake_dataset.upload_calls), 1)
            self.assertEqual(len(fake_dataset.finalize_calls), 1)
            self.assertFalse(fake_dataset.finalize_calls[0]["auto_upload"])

    def test_collision_policy_error_stops_before_clearml(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            (base / "dir1").mkdir()
            (base / "dir2").mkdir()
            (base / "dir1" / "a.txt").write_text("1", encoding="utf-8")
            (base / "dir2" / "a.txt").write_text("2", encoding="utf-8")

            manifest = base / "manifest.csv"
            manifest.write_text(
                "path\n"
                f"{(base / 'dir1' / 'a.txt').as_posix()}\n"
                f"{(base / 'dir2' / 'a.txt').as_posix()}\n",
                encoding="utf-8",
            )

            with redirect_stderr(StringIO()):
                rc = main(
                    [
                        "--manifest",
                        manifest.as_posix(),
                        "--dataset-project",
                        "P",
                        "--dataset-name",
                        "N",
                        "--collision-policy",
                        "error",
                    ]
                )
            self.assertEqual(rc, 1)


if __name__ == "__main__":
    unittest.main()
