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


class TestCliAddinUpdateUninstall(unittest.TestCase):
    def test_addin_update_creates_backup_and_overwrites(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            src1 = root / "src1.xlam"
            src2 = root / "src2.xlam"
            src1.write_bytes(b"v1")
            src2.write_bytes(b"v2")

            dest_dir = root / "dest"

            # First install
            buf = StringIO()
            with redirect_stdout(buf):
                rc = main(["addin", "install", "--xlam", src1.as_posix(), "--dest", dest_dir.as_posix(), "--name", "a.xlam"])
            self.assertEqual(rc, 0)

            # Update (should back up)
            buf = StringIO()
            with redirect_stdout(buf):
                rc = main(["addin", "update", "--xlam", src2.as_posix(), "--dest", dest_dir.as_posix(), "--name", "a.xlam"])
            self.assertEqual(rc, 0)

            lines = [l for l in buf.getvalue().splitlines() if l.strip()]
            self.assertGreaterEqual(len(lines), 1)
            dst = Path(lines[0]).resolve()
            self.assertTrue(dst.exists())
            self.assertEqual(dst.read_bytes(), b"v2")

            backups = sorted(dest_dir.glob("a.xlam.bak.*"))
            self.assertGreaterEqual(len(backups), 1)
            self.assertEqual(backups[-1].read_bytes(), b"v1")

    def test_addin_uninstall_removes_target(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            src = root / "src.xlam"
            src.write_bytes(b"v1")
            dest_dir = root / "dest"

            buf = StringIO()
            with redirect_stdout(buf):
                rc = main(["addin", "install", "--xlam", src.as_posix(), "--dest", dest_dir.as_posix(), "--name", "a.xlam"])
            self.assertEqual(rc, 0)

            installed = dest_dir / "a.xlam"
            self.assertTrue(installed.exists())

            buf = StringIO()
            with redirect_stdout(buf):
                rc = main(["addin", "uninstall", "--dest", dest_dir.as_posix(), "--name", "a.xlam"])
            self.assertEqual(rc, 0)
            self.assertFalse(installed.exists())


if __name__ == "__main__":
    unittest.main()

