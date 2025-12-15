import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.config import load_yaml_config  # noqa: E402


class TestYamlConfig(unittest.TestCase):
    def test_load_yaml_config_resolves_paths(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td).resolve()
            cfg_path = base / "config.yaml"
            cfg_path.write_text(
                "\n".join(
                    [
                        "manifest: manifest.xlsx",
                        "base_dir: data",
                        "output_uri: file://./output",
                        "tags: [a, b]",
                        "use_current_task: true",
                        "dataset-project: P",
                        "dataset-name: N",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            cfg = load_yaml_config(cfg_path.as_posix())
            self.assertEqual(cfg["manifest"], (base / "manifest.xlsx").as_posix())
            self.assertEqual(cfg["base_dir"], (base / "data").as_posix())
            self.assertEqual(cfg["output_uri"], "file://" + (base / "output").as_posix())
            self.assertEqual(cfg["tags"], ["a", "b"])
            self.assertTrue(cfg["use_current_task"])
            self.assertEqual(cfg["dataset_project"], "P")
            self.assertEqual(cfg["dataset_name"], "N")


if __name__ == "__main__":
    unittest.main()

