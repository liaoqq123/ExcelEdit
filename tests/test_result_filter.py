"""结果内筛选的纯假对象回归，不创建 Tk。"""

from pathlib import Path
from types import SimpleNamespace
import unittest

from gui import ExcelEditApp
from json_validation import JsonValidationError, JsonValidationReport


class ResultFilterTests(unittest.TestCase):
    def test_json_errors_can_be_filtered_by_case_insensitive_cell_address(self) -> None:
        target = JsonValidationError(Path("book.xlsx"), "book.xlsx", "Data", 21, 1, "bad", "syntax")
        other = JsonValidationError(Path("book.xlsx"), "book.xlsx", "Data", 21, 2, "wrong", "syntax")
        report = JsonValidationReport(errors=[target, other])
        app = SimpleNamespace(result_source=("json", report), local_sheet_scope=None)
        for query in ("A21", "a21"):
            with self.subTest(query=query):
                app.local_filter_var = SimpleNamespace(get=lambda: query)
                kind, filtered = ExcelEditApp._filtered_result_source(app)
                self.assertEqual(kind, "json")
                self.assertEqual(filtered.errors, [target])
        self.assertEqual(report.errors, [target, other])


if __name__ == "__main__":
    unittest.main()
