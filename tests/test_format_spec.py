import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from clearml_dataset_excel.format_spec import SpecError, load_format_spec  # noqa: E402
from clearml_dataset_excel.format_spec import load_format_spec_from_mapping, spec_to_yaml_dict  # noqa: E402


class TestFormatSpec(unittest.TestCase):
    def test_load_format_spec_ok(self) -> None:
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
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str, required: true}",
                        "    - {name: meas_path, type: path}",
                        "files:",
                        "  - id: meas",
                        "    path_column: meas_path",
                        "    format: csv",
                        "    mapping:",
                        "      axes: {t: time}",
                        "      targets:",
                        "        - {name: f, source: value, type: float}",
                        "      derived:",
                        "        - {name: f2, expr: \"f*2\", type: float}",
                        "      aggregates:",
                        "        - {name: f_mean, source: f, op: mean, output_column: f_mean}",
                        "output:",
                        "  include_file_path_columns: false",
                        "  combine_mode: append",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            spec = load_format_spec(spec_path)
            self.assertEqual(spec.schema_version, 1)
            self.assertIsNotNone(spec.clearml)
            assert spec.clearml is not None
            self.assertEqual(spec.clearml.dataset_project, "P")
            self.assertEqual(spec.clearml.dataset_name, "N")
            self.assertEqual(spec.template.condition_sheet, "Conditions")
            self.assertFalse(spec.addin.enabled)
            self.assertEqual(len(spec.condition_columns), 2)
            self.assertEqual(spec.files[0].file_id, "meas")
            self.assertEqual(spec.files[0].axes.t, "time")
            self.assertEqual(spec.files[0].targets[0].name, "f")
            self.assertEqual(spec.files[0].derived[0].name, "f2")
            self.assertEqual(spec.files[0].aggregates[0].op, "mean")
            self.assertFalse(spec.output.include_file_path_columns)
            self.assertEqual(spec.output.combine_mode, "append")

    def test_load_format_spec_missing_condition_column_for_path(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str}",
                        "files:",
                        "  - id: meas",
                        "    path_column: meas_path",
                        "    mapping:",
                        "      axes: {}",
                        "      targets: []",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            with self.assertRaises(SpecError):
                load_format_spec(spec_path)

    def test_load_format_spec_from_mapping_yaml_shape(self) -> None:
        raw = {
            "schema_version": 1,
            "clearml": {"dataset_project": "P", "dataset_name": "N"},
            "condition": {
                "columns": [
                    {"name": "id", "type": "str", "required": True},
                    {"name": "meas_path", "type": "path"},
                ]
            },
            "files": [
                {
                    "id": "meas",
                    "path_column": "meas_path",
                    "format": "csv",
                    "mapping": {"axes": {"t": "time"}, "targets": [{"name": "f", "source": "value"}]},
                }
            ],
        }
        spec = load_format_spec_from_mapping(raw)
        self.assertEqual(spec.schema_version, 1)
        self.assertIsNotNone(spec.clearml)
        assert spec.clearml is not None
        self.assertEqual(spec.clearml.dataset_project, "P")
        self.assertEqual(spec.files[0].file_id, "meas")
        self.assertEqual(spec.files[0].axes.t, "time")

    def test_load_format_spec_from_mapping_asdict_shape(self) -> None:
        raw = {
            "schema_version": 1,
            "clearml": {"dataset_project": "P", "dataset_name": "N"},
            "condition_columns": [
                {"name": "id", "dtype": "str", "required": True, "description": None, "enum": []},
                {"name": "meas_path", "dtype": "path", "required": False, "description": None, "enum": []},
            ],
            "files": [
                {
                    "file_id": "meas",
                    "path_column": "meas_path",
                    "format": "csv",
                    "axes": {"t": "time"},
                    "targets": [{"name": "f", "source": "value", "dtype": "float"}],
                    "derived": [],
                    "aggregates": [],
                }
            ],
        }
        spec = load_format_spec_from_mapping(raw)
        self.assertEqual(spec.schema_version, 1)
        self.assertEqual(spec.files[0].file_id, "meas")
        self.assertEqual(spec.files[0].axes.t, "time")
        self.assertEqual(spec.files[0].targets[0].name, "f")

    def test_spec_to_yaml_dict_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            spec_path = base / "spec.yaml"
            spec_path.write_text(
                "\n".join(
                    [
                        "schema_version: 1",
                        "addin:",
                        "  enabled: true",
                        "  target_os: mac",
                        "condition:",
                        "  columns:",
                        "    - {name: id, type: str, required: true}",
                        "    - {name: meas_path, type: path}",
                        "files:",
                        "  - id: meas",
                        "    path_column: meas_path",
                        "    mapping: {axes: {t: time}, targets: [{name: f, source: value}]}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            spec1 = load_format_spec(spec_path)
            raw = spec_to_yaml_dict(spec1)
            spec2 = load_format_spec_from_mapping(raw)
            self.assertEqual(spec2, spec1)


if __name__ == "__main__":
    unittest.main()
