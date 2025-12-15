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


class _FakeTaskHyperparams:
    def __init__(self, yaml_text: str) -> None:
        self.id = "ds_base_1"
        self._yaml_text = yaml_text

    def get_configuration_object_as_dict(self, name: str):  # type: ignore[no-untyped-def]
        return {}

    def get_parameter(self, key: str):  # type: ignore[no-untyped-def]
        if key == "dataset_format_spec/yaml":
            return self._yaml_text
        return None


class _FakeBaseDataset:
    def __init__(self, local_copy: Path) -> None:
        self._local_copy = local_copy

    def get_local_copy(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        return self._local_copy.as_posix()


class TestCliAgentReprocessHyperparamYamlOverrides(unittest.TestCase):
    def test_agent_reprocess_uses_hyperparam_yaml_if_present(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "spec").mkdir()
            input_dir = root / "input"
            (input_dir / "external").mkdir(parents=True)

            # Dataset includes spec yaml (should be ignored when hyperparam yaml exists)
            spec_path = root / "spec" / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "output:",
                        "  output_dirname: processed_file",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str}",
                        "    - {name: meas_path, type: path}",
                        "files:",
                        "  - id: m",
                        "    path_column: meas_path",
                        "    format: csv",
                        "    mapping: {axes: {t: time}, targets: [{name: f, source: value}]}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            meas = input_dir / "external" / "000_meas.csv"
            meas.write_text("time,value\n0,1\n1,2\n", encoding="utf-8")

            import openpyxl

            xlsm = input_dir / "conditions.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Conditions"
            ws.append(["id", "meas_path"])
            ws.append(["s1", "/abs/meas.csv"])
            wb.save(xlsm)

            (root / "payload.json").write_text(
                '{"payload_version":1,"spec_path":"spec/spec.yaml","condition_excel":"input/conditions.xlsm","path_map":{"/abs/meas.csv":"input/external/000_meas.csv"}}\n',
                encoding="utf-8",
            )

            # Hyperparameters YAML overrides output_dirname
            hp_yaml = "\n".join(
                [
                    "schema_version: 1",
                    "output:",
                    "  output_dirname: processed_hp",
                    "condition:",
                    "  columns:",
                    "    - {name: id, type: str}",
                    "    - {name: meas_path, type: path}",
                    "files:",
                    "  - id: m",
                    "    path_column: meas_path",
                    "    format: csv",
                    "    mapping: {axes: {t: time}, targets: [{name: f, source: value}]}",
                ]
            )

            out_root = root / "out"
            out_root.mkdir()

            fake_task = _FakeTaskHyperparams(hp_yaml)
            base_ds = _FakeBaseDataset(root)

            def _dataset_get_side_effect(**kwargs):  # type: ignore[no-untyped-def]
                return base_ds

            with patch("clearml.Task") as task_cls, patch("clearml.Dataset") as dataset_cls:
                task_cls.current_task.return_value = fake_task
                dataset_cls.get.side_effect = _dataset_get_side_effect

                buf = StringIO()
                with redirect_stdout(buf):
                    rc = main(["agent", "reprocess", "--no-upload", "--output-root", out_root.as_posix()])

            self.assertEqual(rc, 0)
            lines = [ln.strip() for ln in buf.getvalue().splitlines() if ln.strip()]
            self.assertGreaterEqual(len(lines), 2)
            output_dir = Path(lines[0]).expanduser().resolve()
            self.assertEqual(output_dir.name, "processed_hp")


if __name__ == "__main__":
    unittest.main()

