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


class TestCliAddinUnquarantine(unittest.TestCase):
    def test_addin_unquarantine_noop_or_not_present(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            p = root / "t.xlsm"
            p.write_bytes(b"dummy")

            buf = StringIO()
            with redirect_stdout(buf):
                rc = main(["addin", "unquarantine", "--excel", p.as_posix()])
            self.assertEqual(rc, 0)

            lines = [l for l in buf.getvalue().splitlines() if l.strip()]
            self.assertGreaterEqual(len(lines), 2)
            self.assertEqual(Path(lines[0]).resolve(), p.resolve())
            self.assertIn(lines[1], {"removed", "not_present_or_failed", "noop"})


if __name__ == "__main__":
    unittest.main()

