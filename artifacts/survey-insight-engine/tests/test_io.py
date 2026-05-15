"""Tests for unified file intake."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook
import pandas as pd

from src.io import (
    _detect_scenario,
    _identify_file_roles,
    _normalise_dataframe,
    load_survey_inputs,
)
from tests.conftest import (
    COMBINED_XLSX_PATH,
    DATAMAP_FIXTURE_PATH,
    FORMAT_A_DOCX_PATH,
    FIXTURE_DIR,
    RAW_DECODER_CSV_PATH,
    WORD_RAW_CSV_PATH,
)


class Upload(BytesIO):
    def __init__(self, content: bytes, name: str) -> None:
        super().__init__(content)
        self.name = name


def upload_from_path(path: Path, name: str | None = None) -> Upload:
    return Upload(path.read_bytes(), name or path.name)


def fallback_combined_upload() -> Upload:
    workbook = Workbook()
    raw_sheet = workbook.active
    raw_sheet.title = "Alpha"
    raw_sheet.append(["record", "Q1"])
    raw_sheet.append(["F001", "1"])

    map_sheet = workbook.create_sheet("Beta")
    for row in [
        ["[Q1]: Fallback single select", None, None],
        ["Values: 1-1", None, None],
        [None, 1, "Only"],
        [None, None, None],
    ]:
        map_sheet.append(row)

    buffer = BytesIO()
    workbook.save(buffer)
    workbook.close()
    return Upload(buffer.getvalue(), "fallback.xlsx")


class TestIO(unittest.TestCase):
    def test_normalise_dataframe_removes_nullable_dtypes_and_pd_na(self) -> None:
        dataframe = pd.DataFrame(
            {
                "nullable_int": pd.Series([1, pd.NA], dtype="Int64"),
                "nullable_string": pd.Series(["A", pd.NA], dtype="string"),
                "nullable_boolean": pd.Series([True, pd.NA], dtype="boolean"),
                "nullable_float": pd.Series([1.5, pd.NA], dtype="Float64"),
            }
        )

        normalised = _normalise_dataframe(dataframe)
        dtypes = {str(dtype) for dtype in normalised.dtypes}

        self.assertNotIn("Int64", dtypes)
        self.assertNotIn("string", dtypes)
        self.assertNotIn("boolean", dtypes)
        self.assertNotIn("Float64", dtypes)
        for value in normalised.to_numpy(dtype=object).ravel():
            self.assertIsNot(value, pd.NA)

    def test_load_scenario_a_two_separate_files(self) -> None:
        data_map, raw_df, report = load_survey_inputs(
            [
                upload_from_path(RAW_DECODER_CSV_PATH, "raw_data.csv"),
                upload_from_path(DATAMAP_FIXTURE_PATH, "survey_datamap.xlsx"),
            ]
        )

        self.assertEqual(report.scenario, "A_separate_files")
        self.assertEqual(len(data_map["questions"]), report.questions_parsed)
        self.assertEqual(len(raw_df), report.raw_rows)
        self.assertIn("Q3", raw_df.columns)

    def test_load_scenario_b_combined_xlsx(self) -> None:
        data_map, raw_df, report = load_survey_inputs(
            [upload_from_path(COMBINED_XLSX_PATH, "combined.xlsx")]
        )

        self.assertEqual(report.scenario, "B_combined_xlsx")
        self.assertEqual(report.raw_data_source, "sheet:Sheet1")
        self.assertEqual(report.datamap_source, "sheet:DataMap")
        self.assertEqual(len(data_map["questions"]), 3)
        self.assertEqual(len(raw_df), 2)

    def test_load_scenario_c_word_plus_raw(self) -> None:
        data_map, raw_df, report = load_survey_inputs(
            [
                upload_from_path(FORMAT_A_DOCX_PATH, "survey.docx"),
                upload_from_path(WORD_RAW_CSV_PATH, "word_raw.csv"),
            ]
        )

        self.assertEqual(report.scenario, "C_word_datamap")
        self.assertEqual(report.datamap_source, "survey.docx")
        self.assertGreaterEqual(len(data_map["questions"]), 5)
        self.assertEqual(len(raw_df), 3)

    def test_detect_scenario_docx_present(self) -> None:
        scenario = _detect_scenario(
            [
                upload_from_path(FORMAT_A_DOCX_PATH, "survey.docx"),
                upload_from_path(WORD_RAW_CSV_PATH, "raw.csv"),
            ]
        )

        self.assertEqual(scenario, "C_word_datamap")

    def test_detect_scenario_one_xlsx_is_combined(self) -> None:
        scenario = _detect_scenario(
            [upload_from_path(COMBINED_XLSX_PATH, "combined.xlsx")]
        )

        self.assertEqual(scenario, "B_combined_xlsx")

    def test_detect_scenario_two_files_is_separate(self) -> None:
        scenario = _detect_scenario(
            [
                upload_from_path(RAW_DECODER_CSV_PATH, "raw.csv"),
                upload_from_path(DATAMAP_FIXTURE_PATH, "datamap.xlsx"),
            ]
        )

        self.assertEqual(scenario, "A_separate_files")

    def test_load_report_populated_correctly(self) -> None:
        _data_map, raw_df, report = load_survey_inputs(
            [
                upload_from_path(RAW_DECODER_CSV_PATH, "raw.csv"),
                upload_from_path(DATAMAP_FIXTURE_PATH, "datamap.xlsx"),
            ]
        )

        self.assertEqual(report.raw_rows, len(raw_df))
        self.assertEqual(report.raw_columns, len(raw_df.columns))
        self.assertIsInstance(report.parser_warnings, list)
        self.assertTrue(report.detection_notes)

    def test_file_role_identification_by_name(self) -> None:
        datamap_upload = upload_from_path(DATAMAP_FIXTURE_PATH, "client_datamap.xlsx")
        raw_upload = upload_from_path(RAW_DECODER_CSV_PATH, "client_raw.csv")

        detected_datamap, detected_raw = _identify_file_roles(
            [raw_upload, datamap_upload]
        )

        self.assertIs(detected_datamap, datamap_upload)
        self.assertIs(detected_raw, raw_upload)

    def test_scenario_b_detects_data_and_map_sheets(self) -> None:
        _data_map, _raw_df, report = load_survey_inputs(
            [upload_from_path(COMBINED_XLSX_PATH, "combined.xlsx")]
        )

        self.assertIn("Data sheet: 'Sheet1'", report.detection_notes[0])
        self.assertIn("Map sheet: 'DataMap'", report.detection_notes[0])

    def test_scenario_b_fallback_when_no_keyword_match(self) -> None:
        data_map, raw_df, report = load_survey_inputs([fallback_combined_upload()])

        self.assertEqual(report.scenario, "B_combined_xlsx")
        self.assertEqual(report.raw_data_source, "sheet:Alpha")
        self.assertEqual(report.datamap_source, "sheet:Beta")
        self.assertEqual(data_map["questions"][0]["canonical_id"], "Q1")
        self.assertEqual(len(raw_df), 1)

    def test_cleanup_temp_files_on_error(self) -> None:
        class FakeTempFile:
            def __init__(self, path: Path) -> None:
                self.path = path
                self.name = str(path)

            def write(self, content: bytes) -> int:
                self.path.write_bytes(content)
                return len(content)

            def __enter__(self) -> "FakeTempFile":
                return self

            def __exit__(self, *_args: object) -> None:
                return None

        class TempFactory:
            def __init__(self, directory: Path) -> None:
                self.directory = directory
                self.paths: list[Path] = []
                self.prefix = f"cleanup_temp_case_{id(self)}"

            def __call__(self, suffix: str, delete: bool) -> object:
                path = self.directory / f"{self.prefix}_{len(self.paths)}{suffix}"
                if path.exists():
                    path.unlink()
                self.paths.append(path)
                return FakeTempFile(path)

        factory = TempFactory(Path(tempfile.gettempdir()))
        try:
            with patch("src.io.tempfile.NamedTemporaryFile", factory), patch(
                "src.io.parse_datamap", side_effect=RuntimeError("boom")
            ):
                with self.assertRaisesRegex(RuntimeError, "boom"):
                    load_survey_inputs(
                        [
                            upload_from_path(RAW_DECODER_CSV_PATH, "raw_data.csv"),
                            upload_from_path(
                                DATAMAP_FIXTURE_PATH, "survey_datamap.xlsx"
                            ),
                        ]
                    )
        finally:
            for path in factory.paths:
                if path.exists():
                    path.unlink()

        self.assertTrue(factory.paths)
        self.assertTrue(all(not path.exists() for path in factory.paths))


if __name__ == "__main__":
    unittest.main()
