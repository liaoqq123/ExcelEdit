"""核心搜索、缓存和输入校验回归测试。"""

import json
import tempfile
import threading
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import xlsxwriter

import cell_search
from cache_store import read_cache_data, update_cache_data
from cell_search import ExcelCellSearchIssue, search_excel_cells
from excel_common import EXCEL_MAX_ROWS, ExcelReadError
from file_search import search_files
from popup_gui import parse_positive_int
from worksheet_search import scan_excel_workbooks


class InputValidationTests(unittest.TestCase):
    def test_positive_integer_validation_rejects_non_finite_and_out_of_range_values(self) -> None:
        for value in ("Infinity", "NaN", "-2行", "6.5行", "abc6"):
            with self.subTest(value=value), self.assertRaises(ExcelReadError):
                parse_positive_int(value, "引用行", EXCEL_MAX_ROWS)

        with self.assertRaises(ExcelReadError):
            parse_positive_int(str(EXCEL_MAX_ROWS + 1), "引用行", EXCEL_MAX_ROWS)
        self.assertEqual(parse_positive_int("第 6 行", "引用行", EXCEL_MAX_ROWS), 6)


class CellSearchTests(unittest.TestCase):
    def test_prefilter_ignores_unrelated_inline_and_formula_strings(self) -> None:
        xml = (
            b'<worksheet><sheetData><row r="1">'
            b'<c r="A1" t="inlineStr"><is><t>haystack</t></is></c>'
            b'<c r="B1" t="str"><f>CONCAT("a","b")</f><v>haystack</v></c>'
            b'</row></sheetData></worksheet>'
        )
        self.assertFalse(
            cell_search._prefilter_buffer_has_hint(
                xml.lower(),
                set(),
                cell_search._direct_keyword_hints("needle"),
                "needle",
                set(),
            )
        )

    def test_prefilter_keeps_rich_inline_text_fallback(self) -> None:
        xml = (
            b'<worksheet><sheetData><row r="1">'
            b'<c r="A1" t="inlineStr"><is>'
            b'<r><t>nee</t></r><r><t>dle</t></r>'
            b'</is></c></row></sheetData></worksheet>'
        )
        self.assertTrue(
            cell_search._prefilter_buffer_has_hint(
                xml.lower(),
                set(),
                cell_search._direct_keyword_hints("needle"),
                "needle",
                set(),
            )
        )

    def test_shared_string_id_must_belong_to_a_shared_string_cell(self) -> None:
        xml = (
            b'<worksheet><sheetData><row r="1">'
            b'<c r="A1"><v>7</v></c>'
            b'<c r="B1" t="s"><v>8</v></c>'
            b'</row></sheetData></worksheet>'
        )
        self.assertFalse(
            cell_search._prefilter_buffer_has_hint(
                xml.lower(),
                {b"7"},
                cell_search._direct_keyword_hints("needle"),
                "needle",
                set(),
            )
        )

    def test_shared_string_prefilter_skips_xml_parse_when_keyword_is_absent(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            workbook = xlsxwriter.Workbook(root / "shared.xlsx")
            workbook.add_worksheet("Data").write("A1", "haystack")
            workbook.close()

            with (
                patch("cell_search.ElementTree.iterparse", wraps=cell_search.ElementTree.iterparse) as iterparse,
                cell_search.zipfile.ZipFile(root / "shared.xlsx") as archive,
            ):
                matches = cell_search._read_matching_shared_strings(archive, "needle")

            self.assertEqual(matches, {})
            self.assertEqual(iterparse.call_count, 0)

    def test_shared_string_decoder_handles_rich_text_and_xml_entities(self) -> None:
        item_xml = b"<r><t>Tom &amp; </t></r><r><t>Jerry</t></r>"
        self.assertEqual(cell_search._decode_shared_string_item(item_xml), "Tom & Jerry")

    def test_prefilter_parses_only_the_sheet_with_a_shared_string_hit(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            workbook = xlsxwriter.Workbook(root / "sheets.xlsx")
            workbook.add_worksheet("NoMatch").write("A1", "haystack")
            workbook.add_worksheet("Hit").write("A1", "needle")
            workbook.close()

            with patch(
                "cell_search._search_sheet_xml_cells",
                wraps=cell_search._search_sheet_xml_cells,
            ) as parse_sheet:
                matches = search_excel_cells(root, "needle")

            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0].sheet_name, "Hit")
            self.assertEqual(parse_sheet.call_count, 1)

    def test_large_sheet_prefilter_uses_chunked_scan_without_missing_hits(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            workbook = xlsxwriter.Workbook(root / "chunked.xlsx")
            worksheet = workbook.add_worksheet("Data")
            for row_index in range(100):
                worksheet.write(row_index, 0, "needle" if row_index == 99 else "haystack")
            workbook.close()

            with (
                patch.object(cell_search, "FAST_PREFILTER_MEMORY_LIMIT", 1),
                patch.object(cell_search, "PREFILTER_CHUNK_SIZE", 64),
                patch.object(cell_search, "PREFILTER_OVERLAP_SIZE", 32),
            ):
                matches = search_excel_cells(root, "needle")

            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0].row_index, 100)

    def test_bad_workbook_is_skipped_and_visible_formats_are_searchable(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            workbook_path = root / "good.xlsx"
            workbook = xlsxwriter.Workbook(workbook_path)
            worksheet = workbook.add_worksheet("Data")
            date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
            percent_format = workbook.add_format({"num_format": "0.0%"})
            currency_format = workbook.add_format({"num_format": "$#,##0.00"})
            worksheet.write_datetime("A1", datetime(2026, 7, 14), date_format)
            worksheet.write_number("A2", 0.125, percent_format)
            worksheet.write_number("A3", 1234.5, currency_format)
            worksheet.write("A4", "needle")
            workbook.close()
            (root / "bad.xlsx").write_bytes(b"not an Excel archive")

            issues: list[ExcelCellSearchIssue] = []
            self.assertEqual(len(search_excel_cells(root, "2026-07-14", issues=issues)), 1)
            self.assertEqual(len(search_excel_cells(root, "12.5%")), 1)
            self.assertEqual(len(search_excel_cells(root, "$1,234.50")), 1)
            self.assertEqual(len(issues), 1)
            self.assertEqual(issues[0].path.name, "bad.xlsx")

    def test_result_limit_is_enforced_inside_a_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            workbook = xlsxwriter.Workbook(root / "many.xlsx")
            worksheet = workbook.add_worksheet("Data")
            for row_index in range(10):
                worksheet.write(row_index, 0, "match")
            workbook.close()

            self.assertEqual(len(search_excel_cells(root, "match", max_results=3)), 3)

    def test_exact_cell_search_only_matches_the_complete_value(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            workbook = xlsxwriter.Workbook(root / "exact.xlsx")
            worksheet = workbook.add_worksheet("Data")
            worksheet.write("A1", "needle")
            worksheet.write("A2", "needle extra")
            worksheet.write("A3", "NEEDLE")
            workbook.close()

            matches = search_excel_cells(root, "needle", exact_match=True)

            self.assertEqual([match.row_index for match in matches], [1, 3])


class FileSearchTests(unittest.TestCase):
    def test_file_result_limit_is_enforced(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            for index in range(10):
                (root / f"file-{index}.txt").write_text("", encoding="utf-8")
            self.assertEqual(len(search_files(root, max_results=4)), 4)

    def test_file_suffix_filter_accepts_common_input_formats(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            for file_name in ("report.XLSX", "notes.txt", "archive.tar.gz", "image.png"):
                (root / file_name).write_text("", encoding="utf-8")

            results = search_files(root, suffixes="xlsx, *.tar.gz")

            self.assertEqual([result.file_name for result in results], ["archive.tar.gz", "report.XLSX"])

    def test_exact_file_search_only_matches_the_complete_file_name(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            (root / "report.txt").write_text("", encoding="utf-8")
            (root / "report-copy.txt").write_text("", encoding="utf-8")

            results = search_files(root, "report.txt", exact_match=True)

            self.assertEqual([result.file_name for result in results], ["report.txt"])

    def test_file_search_can_include_matching_folders(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            (root / "Target").mkdir()
            (root / "Target Extra").mkdir()
            (root / "Target" / "inside.txt").write_text("", encoding="utf-8")

            results = search_files(root, "target", exact_match=True, include_folders=True)

            self.assertEqual([result.file_name for result in results], ["Target"])
            self.assertTrue(results[0].is_directory)


class WorksheetSearchTests(unittest.TestCase):
    def test_exact_worksheet_search_matches_complete_sheet_or_workbook_name(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            exact_workbook = xlsxwriter.Workbook(root / "exact-book.xlsx")
            exact_workbook.add_worksheet("Target")
            exact_workbook.close()
            similar_workbook = xlsxwriter.Workbook(root / "similar-book.xlsx")
            similar_workbook.add_worksheet("Target Extra")
            similar_workbook.close()

            sheet_results = scan_excel_workbooks(root, "target", exact_match=True)
            workbook_results = scan_excel_workbooks(root, "exact-book", exact_match=True)

            self.assertEqual([result.workbook_name for result in sheet_results], ["exact-book.xlsx"])
            self.assertEqual([result.workbook_name for result in workbook_results], ["exact-book.xlsx"])


class CacheStoreTests(unittest.TestCase):
    def test_legacy_migration_and_concurrent_atomic_updates(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            cache_file = root / "user" / "cache_data.json"
            legacy_file = root / "legacy.json"
            legacy_file.write_text(json.dumps({"legacy": True}), encoding="utf-8")

            self.assertEqual(read_cache_data(cache_file, legacy_file), {"legacy": True})

            threads = [
                threading.Thread(
                    target=update_cache_data,
                    args=({f"key-{index}": index}, cache_file, legacy_file),
                )
                for index in range(12)
            ]
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join()

            payload = json.loads(cache_file.read_text(encoding="utf-8"))
            self.assertTrue(payload["legacy"])
            for index in range(12):
                self.assertEqual(payload[f"key-{index}"], index)
            self.assertFalse(list(cache_file.parent.glob("*.tmp")))


if __name__ == "__main__":
    unittest.main()
