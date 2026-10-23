"""核心搜索、缓存和输入校验回归测试。"""

import json
import tempfile
import threading
import unittest
import zipfile
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import xlsxwriter

import cell_search
from cache_store import read_cache_data, update_cache_data
from cell_search import ExcelCellSearchIssue, search_excel_cells
from excel_common import EXCEL_MAX_ROWS, ExcelReadError, is_excel_file
from file_search import search_files
from json_validation import validate_excel_json_columns
from input_validation import parse_positive_int
from worksheet_search import scan_excel_workbooks


class ExcelFileEligibilityTests(unittest.TestCase):
    def test_non_excel_and_lock_files_do_not_query_file_status(self) -> None:
        with patch.object(Path, "is_file", side_effect=AssertionError("Unrelated files must not query disk status")):
            for name in ("icon.png", "scene.prefab", "code.py", "README", "~$book.xlsx", "~$book.XLSM"):
                with self.subTest(name=name):
                    self.assertFalse(is_excel_file(Path(name)))

    def test_excel_candidates_still_require_an_existing_regular_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = root / "book.XLSX"
            workbook.write_bytes(b"fixture")
            named_folder = root / "folder.xlsx"
            named_folder.mkdir()
            self.assertTrue(is_excel_file(workbook))
            self.assertFalse(is_excel_file(named_folder))
            self.assertFalse(is_excel_file(root / "missing.xls"))


class InputValidationTests(unittest.TestCase):
    def test_positive_integer_validation_rejects_non_finite_and_out_of_range_values(self) -> None:
        for value in ("Infinity", "NaN", "-2行", "6.5行", "abc6"):
            with self.subTest(value=value), self.assertRaises(ExcelReadError):
                parse_positive_int(value, "引用行", EXCEL_MAX_ROWS)

        with self.assertRaises(ExcelReadError):
            parse_positive_int(str(EXCEL_MAX_ROWS + 1), "引用行", EXCEL_MAX_ROWS)
        self.assertEqual(parse_positive_int("第 6 行", "引用行", EXCEL_MAX_ROWS), 6)


class CellSearchTests(unittest.TestCase):
    def test_metadata_formula_entities_and_chinese_do_not_force_full_sheet_parse(self) -> None:
        xml = (
            '<worksheet><sheetData><row r="1">'
            '<c r="A1" t="str"><f>"说明"&amp;"其他"</f><v>haystack</v></c>'
            '</row></sheetData><hyperlinks><hyperlink location="说明 &amp; Other"/></hyperlinks></worksheet>'
        ).encode("utf-8")
        self.assertFalse(cell_search._prefilter_buffer_has_hint(
            xml.lower(), set(), cell_search._direct_keyword_hints("needle"), "needle", set(),
        ))

    def test_negative_chinese_shared_string_search_does_not_decode_every_entry(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "chinese.xlsx"
            workbook = xlsxwriter.Workbook(path)
            sheet = workbook.add_worksheet("Data")
            for row in range(1000):
                sheet.write(row, 0, f"普通配置说明{row}")
            workbook.close()
            with zipfile.ZipFile(path) as archive, patch(
                "cell_search._decode_shared_string_item", wraps=cell_search._decode_shared_string_item,
            ) as decode:
                matches = cell_search._read_matching_shared_strings(archive, "不存在的关键词")
            self.assertEqual(matches, {})
            self.assertEqual(decode.call_count, 0)

    def test_negative_inline_unicode_and_alphanumeric_search_skip_full_parse(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "inline.xlsx"
            workbook = xlsxwriter.Workbook(path, {"constant_memory": True})
            sheet = workbook.add_worksheet("Data")
            sheet.write("A1", "其他中文说明")
            sheet.write_datetime("A2", datetime(2026, 7, 14), workbook.add_format({"num_format": "yyyy-mm-dd"}))
            workbook.close()
            with patch("cell_search._search_sheet_xml_cells", wraps=cell_search._search_sheet_xml_cells) as parse, patch(
                "cell_search._read_workbook_format_info", wraps=cell_search._read_workbook_format_info,
            ) as formats:
                self.assertEqual(search_excel_cells(directory, "角色1001001"), [])
            self.assertEqual(parse.call_count, 0)
            self.assertEqual(formats.call_count, 0)

    def test_negative_numeric_query_checks_formatted_values_before_full_parse(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = xlsxwriter.Workbook(Path(directory) / "formatted.xlsx")
            sheet = workbook.add_worksheet("Data")
            sheet.write_datetime("A1", datetime(2026, 7, 14), workbook.add_format({"num_format": "yyyy-mm-dd"}))
            sheet.write_number("A2", 12.34, workbook.add_format({"num_format": "0%"}))
            workbook.close()
            with patch("cell_search._search_sheet_xml_cells", wraps=cell_search._search_sheet_xml_cells) as parse:
                self.assertEqual(search_excel_cells(directory, "999999"), [])
            self.assertEqual(parse.call_count, 0)
            hits = search_excel_cells(directory, "1234")
            self.assertEqual([hit.row_index for hit in hits], [2])
            self.assertEqual(hits[0].value, "1234%")

    def test_early_excel_dates_and_rounding_still_match_display_values(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = xlsxwriter.Workbook(Path(directory) / "rounding.xlsx")
            sheet = workbook.add_worksheet("Data")
            sheet.write_number("A1", 1, workbook.add_format({"num_format": "yyyy-mm-dd"}))
            sheet.write_number("A2", 999999.999, workbook.add_format({"num_format": "0.00"}))
            workbook.close()
            self.assertEqual([hit.row_index for hit in search_excel_cells(directory, "1900")], [1])
            self.assertEqual([hit.row_index for hit in search_excel_cells(directory, "1000000")], [2])

    def test_negative_search_with_enabled_data_filters_keeps_fast_skip(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = xlsxwriter.Workbook(Path(directory) / "filtered.xlsx")
            sheet = workbook.add_worksheet("Data")
            sheet.write("A1", "字段说明")
            for row in range(1, 200):
                sheet.write(row, 0, f"其他中文配置{row}")
            sheet.write_formula("B2", '="说明"&"文字"', None, "说明文字")
            workbook.close()
            config = cell_search.DataFilterConfig(enable_header_filter=True, header_row_count=1)
            with patch("cell_search._search_sheet_xml_cells", wraps=cell_search._search_sheet_xml_cells) as parse:
                self.assertEqual(search_excel_cells(directory, "不存在的技能展示", data_filter_config=config), [])
            self.assertEqual(parse.call_count, 0)

    def test_partial_display_values_are_searchable(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = xlsxwriter.Workbook(Path(directory) / "display.xlsx")
            sheet = workbook.add_worksheet("Data")
            sheet.write_datetime("A1", datetime(2026, 7, 14), workbook.add_format({"num_format": "yyyy-mm-dd"}))
            sheet.write_boolean("B1", True)
            sheet.write_boolean("C1", False)
            workbook.close()
            for keyword, column in (("2026", 1), ("tru", 2), ("als", 3)):
                with self.subTest(keyword=keyword):
                    hits = search_excel_cells(directory, keyword)
                    self.assertEqual([hit.column_index for hit in hits], [column])
                    self.assertEqual(search_excel_cells(directory, keyword, exact_match=True), [])

    def test_entities_and_unicode_case_survive_prefilter(self) -> None:
        for inline in (False, True):
            with self.subTest(inline=inline), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "text.xlsx"
                workbook = xlsxwriter.Workbook(path, {"constant_memory": inline})
                sheet = workbook.add_worksheet("Data")
                sheet.write("A1", "铁剑")
                sheet.write("A2", "Äpfel")
                workbook.close()
                with zipfile.ZipFile(path) as archive:
                    parts = {info.filename: archive.read(info.filename) for info in archive.infolist()}
                with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
                    for name, content in parts.items():
                        archive.writestr(name, content.replace("铁剑".encode(), b"&#38081;&#21073;"))
                self.assertEqual([hit.row_index for hit in search_excel_cells(directory, "铁剑")], [1])
                self.assertEqual([hit.row_index for hit in search_excel_cells(directory, "äpfel")], [2])

    def test_shared_reference_handles_empty_cells_whitespace_and_long_boundaries(self) -> None:
        xml = b'<c r="A1"/><c r="B1" t=\'s\'><v> 007 </v></c>'
        self.assertTrue(cell_search._has_typed_shared_string_reference(xml, {b"7"}))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "long.xlsx"
            workbook = xlsxwriter.Workbook(path)
            workbook.add_worksheet("Data").write("A1", "needle")
            workbook.close()
            with zipfile.ZipFile(path) as archive:
                parts = {info.filename: archive.read(info.filename) for info in archive.infolist()}
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
                for name, content in parts.items():
                    if name == "xl/worksheets/sheet1.xml":
                        content = content.replace(b'<v>0</v>', b'<v>' + b' ' * 10000 + b'0</v>')
                    archive.writestr(name, content)
            with patch.object(cell_search, "FAST_PREFILTER_MEMORY_LIMIT", 1024), patch.object(cell_search, "PREFILTER_CHUNK_SIZE", 31):
                self.assertEqual(len(search_excel_cells(directory, "needle")), 1)

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


class JsonValidationTests(unittest.TestCase):
    def test_validator_checks_identified_columns_and_skips_dollar_sheets(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            workbook = xlsxwriter.Workbook(root / "json-data.xlsx")
            enabled_sheet = workbook.add_worksheet("Enabled")
            enabled_sheet.write("A4", '{"id": 1}')
            enabled_sheet.write("A5", '{"id": }')
            enabled_sheet.write("A6", '["one", "two"]')
            enabled_sheet.write("A7", '{"id": 2}')
            enabled_sheet.write("B4", "ordinary text")
            enabled_sheet.write("B6", "ordinary text")
            enabled_sheet.write("C4", "{not valid json")
            enabled_sheet.write("C6", "{still not valid json")
            disabled_sheet = workbook.add_worksheet("$Disabled")
            disabled_sheet.write("A4", '{"ignored": }')
            disabled_sheet.write("A6", '{"ignored": 1}')
            workbook.close()

            report = validate_excel_json_columns(root)

        self.assertEqual(report.workbook_count, 1)
        self.assertEqual(report.sheet_count, 1)
        self.assertEqual(report.json_column_count, 1)
        self.assertEqual(report.checked_cell_count, 4)
        self.assertEqual(len(report.errors), 1)
        error = report.errors[0]
        self.assertEqual((error.sheet_name, error.row_index, error.column_index), ("Enabled", 5, 1))
        self.assertEqual(error.value, '{"id": }')
        self.assertIn("JSON", error.message)


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
    def setUp(self) -> None:
        for patcher in (
            patch("worksheet_search.read_cache_data", return_value={}),
            patch("worksheet_search.update_cache_data", return_value=True),
            patch("worksheet_search._sheet_name_cache", None),
            patch("worksheet_search._sheet_name_cache_dirty", False),
            patch("worksheet_search._load_legacy_sheet_name_cache", return_value={}),
        ):
            patcher.start()
            self.addCleanup(patcher.stop)

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
