import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.vba_addin import generate_vba_module  # noqa: E402


class TestVbaAddin(unittest.TestCase):
    def test_generate_vba_module_contains_keys(self) -> None:
        text = generate_vba_module(meta_sheet_name="_meta")
        self.assertIn('Private Const META_SHEET As String = "_meta"', text)
        self.assertIn("addin_enabled", text)
        self.assertIn("addin_target_os", text)
        self.assertIn("addin_spec_filename", text)
        self.assertIn("addin_command_mac", text)
        self.assertIn("addin_command_windows", text)
        self.assertIn("cd /d", text)
        self.assertIn('"${SPEC}"', text)
        self.assertIn('"${EXCEL}"', text)
        self.assertIn('"${{SPEC}}"', text)
        self.assertIn('"${{EXCEL}}"', text)

    def test_generate_vba_module_has_shell_quote(self) -> None:
        text = generate_vba_module(meta_sheet_name="_meta")
        self.assertIn("Private Function ShellQuote", text)
        self.assertIn("Replace(s, Chr(39), Chr(39) & Chr(34) & Chr(39) & Chr(34) & Chr(39))", text)
        self.assertIn("ShellQuote = Chr(39) & t & Chr(39)", text)
        self.assertIn('/bin/zsh -lc ', text)


if __name__ == "__main__":
    unittest.main()
