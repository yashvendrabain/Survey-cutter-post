from __future__ import annotations

from pathlib import Path
import uuid
import unittest

from src.models import DenominatorPolicy, QuestionSpec, QuestionType
from src.raw_decoder import decode_numeric_cell, decode_raw_data


def _rating_spec(
    *,
    label_to_numeric_value: dict[str, float] | None = None,
    na_label_set: frozenset[str] | None = None,
    allowed_numeric_range: tuple[float, float] = (0.0, 10.0),
) -> QuestionSpec:
    mapping = label_to_numeric_value or {
        "0 (extremely low)": 0.0,
        "1": 1.0,
        "2": 2.0,
        "3": 3.0,
        "4": 4.0,
        "5": 5.0,
        "6": 6.0,
        "7": 7.0,
        "8": 8.0,
        "9": 9.0,
        "10 (extremely high)": 10.0,
    }
    return QuestionSpec(
        question_id="Q30",
        canonical_id="Q30",
        question_text="Rate the following",
        question_type=QuestionType.GRID_RATED,
        raw_columns=("Q30r1c1",),
        option_map={},
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        grid_row_labels={"Q30r1c1": "Pre-purchase familiarity"},
        label_to_numeric_value=mapping,
        na_label_set=na_label_set
        if na_label_set is not None
        else frozenset({"This was not something I considered"}),
        allowed_numeric_range=allowed_numeric_range,
    )


class TestRawDecoderLabelCoercion(unittest.TestCase):
    def _decode(self, value, spec: QuestionSpec | None = None):
        warnings: list[dict] = []
        decoded = decode_numeric_cell(
            value,
            spec or _rating_spec(),
            warnings,
            row_idx=7,
            column_name="Q30r1c1",
        )
        return decoded, warnings

    def test_native_int_in_range_returns_float(self) -> None:
        decoded, warnings = self._decode(8)
        self.assertEqual(decoded, 8.0)
        self.assertEqual(warnings, [])

    def test_native_int_out_of_range_returns_none_with_warning(self) -> None:
        decoded, warnings = self._decode(11)
        self.assertIsNone(decoded)
        self.assertEqual(len(warnings), 1)
        self.assertIn("out_of_range", warnings[0]["action"])

    def test_extremely_high_label_returns_numeric_value(self) -> None:
        decoded, warnings = self._decode("10 (extremely high)")
        self.assertEqual(decoded, 10.0)
        self.assertEqual(warnings, [])

    def test_extremely_low_label_returns_numeric_value(self) -> None:
        decoded, warnings = self._decode("0 (extremely low)")
        self.assertEqual(decoded, 0.0)
        self.assertEqual(warnings, [])

    def test_known_na_label_returns_none_without_warning(self) -> None:
        decoded, warnings = self._decode("This was not something I considered")
        self.assertIsNone(decoded)
        self.assertEqual(warnings, [])

    def test_bare_numeric_string_returns_float(self) -> None:
        decoded, warnings = self._decode("8")
        self.assertEqual(decoded, 8.0)
        self.assertEqual(warnings, [])

    def test_unrecognized_string_returns_none_with_warning(self) -> None:
        decoded, warnings = self._decode("Cost")
        self.assertIsNone(decoded)
        self.assertEqual(len(warnings), 1)
        self.assertEqual(warnings[0]["action"], "unrecognized_value_treated_as_missing")

    def test_none_returns_none_without_warning(self) -> None:
        decoded, warnings = self._decode(None)
        self.assertIsNone(decoded)
        self.assertEqual(warnings, [])

    def test_empty_string_returns_none_without_warning(self) -> None:
        decoded, warnings = self._decode("")
        self.assertIsNone(decoded)
        self.assertEqual(warnings, [])

    def test_float_in_range_returns_float(self) -> None:
        decoded, warnings = self._decode(7.5)
        self.assertEqual(decoded, 7.5)
        self.assertEqual(warnings, [])

    def test_decimal_with_integer_scale_is_allowed_when_in_range(self) -> None:
        spec = _rating_spec(
            label_to_numeric_value={str(i): float(i) for i in range(1, 6)},
            na_label_set=frozenset(),
            allowed_numeric_range=(1.0, 5.0),
        )
        decoded, warnings = self._decode(4.5, spec)
        self.assertEqual(decoded, 4.5)
        self.assertEqual(warnings, [])

    def test_boolean_returns_none_with_warning(self) -> None:
        decoded, warnings = self._decode(True)
        self.assertIsNone(decoded)
        self.assertEqual(len(warnings), 1)
        self.assertEqual(warnings[0]["action"], "unexpected_type_bool")

    def test_decode_raw_data_applies_numeric_label_metadata_to_column(self) -> None:
        scratch_root = Path.cwd() / "outputs" / "test_tmp"
        scratch_root.mkdir(parents=True, exist_ok=True)
        raw_path = scratch_root / f"raw_{uuid.uuid4().hex}.csv"
        raw_path.write_text(
            "respondent_id,Q30r1c1\n"
            "1,10 (extremely high)\n"
            "2,8\n"
            "3,This was not something I considered\n",
            encoding="utf-8",
        )
        data_map = {
            "questions": [
                {
                    "canonical_id": "Q30",
                    "raw_id": "Q30",
                    "question_text": "Rate the following",
                    "type_hint": "values_range",
                    "value_range": (1, 12),
                    "options": [],
                    "sub_columns": [("Q30r1c1", "Winner")],
                    "parent_canonical_id": None,
                    "source_row": 1,
                    "warnings": [],
                    "label_to_numeric_value": _rating_spec().label_to_numeric_value,
                    "na_label_set": _rating_spec().na_label_set,
                    "allowed_numeric_range": _rating_spec().allowed_numeric_range,
                }
            ],
            "source_path": "datamap.xlsx",
            "sheet_name": "Sheet1",
            "total_rows_in_sheet": 1,
            "parser_warnings": [],
        }

        decoded_df, report = decode_raw_data(str(raw_path), data_map)

        self.assertEqual(decoded_df["Q30r1c1"].tolist()[:2], [10.0, 8.0])
        self.assertTrue(decoded_df["Q30r1c1"].isna().iloc[2])
        self.assertEqual(report.decoder_warnings, ())


if __name__ == "__main__":
    unittest.main()
