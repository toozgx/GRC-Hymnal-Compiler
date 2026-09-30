"""Regression tests using synthetic workbooks; no live sheet or Typst needed."""

import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

import openpyxl

SRC = Path(__file__).resolve().parents[1] / "src"
sys.path.insert(0, str(SRC))
from transform import load_title_page


def valid_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheets = {
        "Hymns Metadata": [
            ["ID", "Title", "Category", "Tune 1", "Tune 2", "Tune 3", "Main Tune", "Status"],
            ["TH", "Test Hymn", "Worship", "Test Tune", None, None, 1, "active"],
        ],
        "Lyrics Section": [
            ["Hymn ID", "Sequence", "Type", "Label", "Text"],
            ["TH", 1, "Stanza", 1, "First line\nSecond line"],
        ],
        "Contributors": [
            ["Hymn ID", "Sequence", "Role", "Person Name", "Note", "Tune Slot"],
            ["TH", 1, "author", "Test Author", None, None],
            ["TH", 1, "composer", "Test Composer", None, 1],
        ],
        "Categories": [["Category", "Sequence"], ["Worship", 1]],
        "Title Page": [["Key", "Value"], ["Title", "Test Hymnal"], ["Subtitle", "Test edition"]],
    }
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    return wb


class WorkbookValidationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / "workbook.xlsx"

    def validate(self, wb):
        wb.save(self.path)
        wb.close()
        return subprocess.run(
            [sys.executable, str(SRC / "validate_workbook.py"), str(self.path)],
            capture_output=True, text=True, encoding="utf-8",
            env=dict(os.environ, PYTHONUTF8="1"),
        )

    def assert_rejected(self, result, message):
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn(message, result.stdout)
        self.assertIn("Summary:", result.stdout)
        self.assertNotIn("Traceback", result.stdout + result.stderr)

    def test_valid_workbook_keeps_numeric_labels_sequences_and_tune_slots(self):
        result = self.validate(valid_workbook())
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("Summary: 0 error(s), 0 warning(s)", result.stdout)
        self.assertEqual(load_title_page(self.path), {"title": "Test Hymnal", "subtitle": "Test edition"})

    def test_text_fields_reject_numbers_with_row_and_column(self):
        fields = {
            "Hymns Metadata": {"A2": "ID", "B2": "Title", "C2": "Category", "D2": "Tune 1", "E2": "Tune 2", "F2": "Tune 3", "H2": "Status"},
            "Lyrics Section": {"A2": "Hymn ID", "C2": "Type", "E2": "Text"},
            "Contributors": {"A2": "Hymn ID", "C2": "Role", "D2": "Person Name", "E2": "Note"},
            "Categories": {"A2": "Category"},
        }
        for sheet, cells in fields.items():
            for cell, column in cells.items():
                with self.subTest(sheet=sheet, column=column):
                    wb = valid_workbook()
                    wb[sheet][cell] = 123
                    self.assert_rejected(self.validate(wb), f"{sheet} row 2, {column}: expected text")

    def test_boolean_lyric_is_not_accepted_as_text(self):
        wb = valid_workbook()
        wb["Lyrics Section"]["E2"] = True
        self.assert_rejected(self.validate(wb), "Lyrics Section row 2, Text: expected text")

    def test_duplicate_title_keys_are_rejected_even_when_blank(self):
        for key, value in [("Title", None), (" title ", "Other title"), ("SUBTITLE", "Other subtitle")]:
            with self.subTest(key=key, value=value):
                wb = valid_workbook()
                wb["Title Page"].append([key, value])
                self.assert_rejected(self.validate(wb), "Title Page row 4: duplicate key")
                with self.assertRaisesRegex(ValueError, "duplicate key"):
                    load_title_page(self.path)

    def test_required_title_page_values_must_be_nonblank_text(self):
        for cell in ("B2", "B3"):
            for value in (None, "   ", 123, False):
                with self.subTest(cell=cell, value=value):
                    wb = valid_workbook()
                    wb["Title Page"][cell] = value
                    self.assert_rejected(self.validate(wb), "Title Page")
                    with self.assertRaises(ValueError):
                        load_title_page(self.path)

    def test_title_page_keys_and_values_are_normalized_consistently(self):
        wb = valid_workbook()
        wb["Title Page"]["A2"] = " TITLE "
        wb["Title Page"]["B2"] = " Test Hymnal "
        result = self.validate(wb)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(load_title_page(self.path)["title"], "Test Hymnal")

    def test_missing_title_page_key_is_rejected(self):
        wb = valid_workbook()
        wb["Title Page"].delete_rows(2)
        self.assert_rejected(self.validate(wb), "missing or empty value for key 'title'")

    def test_numeric_title_page_key_is_reported(self):
        wb = valid_workbook()
        wb["Title Page"].append([123, "Unexpected key"])
        self.assert_rejected(self.validate(wb), "Title Page row 4, Key: expected text")

    def test_domain_checks_still_reject_missing_lyrics(self):
        wb = valid_workbook()
        wb["Lyrics Section"].delete_rows(2)
        self.assert_rejected(self.validate(wb), "Status=active but has no Lyrics Section content")

    def test_domain_checks_still_reject_missing_category(self):
        wb = valid_workbook()
        wb["Categories"].delete_rows(2)
        self.assert_rejected(self.validate(wb), "has no row in Categories sheet")

    def test_missing_column_stops_before_domain_checks(self):
        wb = valid_workbook()
        wb["Hymns Metadata"]["B1"] = "Wrong heading"
        self.assert_rejected(self.validate(wb), "missing expected column(s): ['Title']")


if __name__ == "__main__":
    unittest.main()
