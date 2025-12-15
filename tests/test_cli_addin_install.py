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


class TestCliAddinInstall(unittest.TestCase):
    def test_addin_install_copies_xlam(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            src = root / "src.xlam"
            src.write_bytes(b"dummy")

            dest_dir = root / "dest"
            buf = StringIO()
            with redirect_stdout(buf):
                rc = main(["addin", "install", "--xlam", src.as_posix(), "--dest", dest_dir.as_posix()])
            self.assertEqual(rc, 0)

            out_path = Path(buf.getvalue().splitlines()[0]).expanduser().resolve()
            self.assertTrue(out_path.exists())
            self.assertEqual(out_path.read_bytes(), b"dummy")

    def test_addin_install_respects_name(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            src = root / "src.xlam"
            src.write_bytes(b"dummy")

            dest_dir = root / "dest"
            buf = StringIO()
            with redirect_stdout(buf):
                rc = main(
                    [
                        "addin",
                        "install",
                        "--xlam",
                        src.as_posix(),
                        "--dest",
                        dest_dir.as_posix(),
                        "--name",
                        "clearml_dataset_excel_addin.xlam",
                    ]
                )
            self.assertEqual(rc, 0)
            out_path = Path(buf.getvalue().splitlines()[0]).expanduser().resolve()
            self.assertEqual(out_path.name, "clearml_dataset_excel_addin.xlam")


if __name__ == "__main__":
    unittest.main()

