from __future__ import annotations

from pathlib import Path
import uuid
import unittest

from openpyxl import Workbook

from src.datamap_parser import parse_datamap


def _write_datamap(rows: list[tuple[object | None, object | None, object | None]]) -> str:
    scratch_root = Path.cwd() / "outputs" / "test_tmp"
    scratch_root.mkdir(parents=True, exist_ok=True)
    path = scratch_root / f"datamap_{uuid.uuid4().hex}.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    for row in rows:
        worksheet.append(row)
    workbook.save(path)
    workbook.close()
    return str(path)


class TestDatamapParserNumericCoercionSetup(unittest.TestCase):
    def test_annotated_numeric_labels_do_not_build_numeric_label_metadata(self) -> None:
        rows = [
            ("[Q30r1]: Pre-purchase familiarity - Please rate", None, None),
            ("Values: 1 - 12", None, None),
        ]
        labels = [
            "0 (extremely low)",
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10 (extremely high)",
            "This was not something I considered",
        ]
        for code, label in enumerate(labels, start=1):
            rows.append((None, code, label))

        data_map = parse_datamap(_write_datamap(rows), min_questions=1)
        question = data_map["questions"][0]

        self.assertEqual(question.get("label_to_numeric_value", {}), {})
        self.assertEqual(question.get("na_label_set", frozenset()), frozenset())
        self.assertIsNone(question.get("allowed_numeric_range"))

    def test_integer_labels_build_numeric_label_metadata_without_na_labels(self) -> None:
        rows = [
            ("[Q1]: Rate from one to five", None, None),
            ("Values: 1 - 5", None, None),
        ]
        for code in range(1, 6):
            rows.append((None, code, str(code)))

        data_map = parse_datamap(_write_datamap(rows), min_questions=1)
        question = data_map["questions"][0]

        self.assertEqual(question["label_to_numeric_value"], {str(i): float(i) for i in range(1, 6)})
        self.assertEqual(question["na_label_set"], frozenset())
        self.assertEqual(question["allowed_numeric_range"], (1.0, 5.0))

    def test_open_text_does_not_build_numeric_label_metadata(self) -> None:
        data_map = parse_datamap(
            _write_datamap(
                [
                    ("[QTEXT]: Comment", None, None),
                    ("Open text response", None, None),
                    (None, 1, "1"),
                    (None, 2, "2"),
                ]
            ),
            min_questions=1,
        )
        question = data_map["questions"][0]

        self.assertEqual(question.get("label_to_numeric_value", {}), {})
        self.assertEqual(question.get("na_label_set", frozenset()), frozenset())
        self.assertIsNone(question.get("allowed_numeric_range"))

    def test_sub_column_only_question_does_not_build_numeric_label_metadata(self) -> None:
        data_map = parse_datamap(
            _write_datamap(
                [
                    ("[Q_MULTI]: Select all that apply", None, None),
                    ("Values: 0 - 1", None, None),
                    (None, "[Q_MULTIr1]", "1"),
                    (None, "[Q_MULTIr2]", "2"),
                ]
            ),
            min_questions=1,
        )
        question = data_map["questions"][0]

        self.assertEqual(question.get("label_to_numeric_value", {}), {})
        self.assertEqual(question.get("na_label_set", frozenset()), frozenset())
        self.assertIsNone(question.get("allowed_numeric_range"))


if __name__ == "__main__":
    unittest.main()
