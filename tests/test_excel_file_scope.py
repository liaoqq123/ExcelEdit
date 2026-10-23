"""Workbook and cell searches only read supported Excel files; no GUI or Office."""

from contextlib import nullcontext
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import cell_search
from excel_common import iter_excel_files
import worksheet_search


class ExcelEnumerationMetadataTests(unittest.TestCase):
    def test_unrelated_files_and_lock_files_skip_file_queries_without_blocking_recursion(self):
        root = Path("fixture-root").absolute()
        nested = root / "folder.xlsx"

        def entry(parent, name, *, directory=False, regular=False, candidate=False):
            return SimpleNamespace(
                name=name,
                path=str(parent / name),
                is_dir=Mock(return_value=directory),
                is_file=(
                    Mock(return_value=regular) if candidate else
                    Mock(side_effect=AssertionError(f"Unrelated entry queried as a file: {name}"))
                ),
            )

        unrelated = [
            entry(root, name) for name in (
                "icon.png", "rows.csv", "settings.json", "book.xlsx.bak",
                "~$locked.xlsx", "~$locked.XLSM", "README",
            )
        ]
        candidates = [
            entry(root, "BOOK.XLSX", regular=True, candidate=True),
            entry(root, "binary.xLsB", regular=True, candidate=True),
            # A symlink/non-regular candidate must still fail is_file(False).
            entry(root, "linked.xlsx", regular=False, candidate=True),
        ]
        directory = entry(root, nested.name, directory=True)
        child = entry(nested, "inside.XLS", regular=True, candidate=True)
        by_folder = {
            root: [*unrelated, *candidates, directory],
            nested: [child],
        }

        def scan(folder):
            return nullcontext(iter(by_folder[Path(folder)]))

        with (
            patch("excel_common.validate_excel_folder", return_value=root),
            patch("excel_common.os.scandir", side_effect=scan) as scanner,
            patch.object(Path, "is_file", side_effect=AssertionError("DirEntry metadata must be reused")),
        ):
            found = set(iter_excel_files(root))

        self.assertEqual(found, {root / "BOOK.XLSX", root / "binary.xLsB", nested / "inside.XLS"})
        self.assertEqual({Path(call.args[0]) for call in scanner.call_args_list}, {root, nested})
        for item in [*unrelated, *candidates, directory, child]:
            item.is_dir.assert_called_once_with(follow_symlinks=False)
        for item in [*unrelated, directory]:
            item.is_file.assert_not_called()
        for item in [*candidates, child]:
            item.is_file.assert_called_once_with(follow_symlinks=False)


class ExcelSearchFileScopeTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.expected = set()
        for name in (
            "config.XLSX", "macros.xLsM", "template.XLTX", "macro-template.xltm",
            "legacy.XLS", "binary.xLsB",
            "directory.xlsx/inside.xlsm",
            "resources.png/textures/nested.XLSB",
            "~$resources.xlsm/actual.xltx",
        ):
            target = self.root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(b"reader is replaced by a test double")
            self.expected.add(target)

        for name in (
            "icon.png", "rows.csv", "settings.json", "notes.txt", "README",
            "book.xlsx.bak", "book.xlsx.json", "~$ignored.XLSX", "~$legacy.xls",
            "resources.png/textures/image.png", "directory.xlsx/nested.csv",
        ):
            target = self.root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(b"unrelated resource")

    def test_recursive_enumeration_keeps_all_six_excel_types_and_ignores_other_files(self):
        self.assertEqual(set(iter_excel_files(self.root)), self.expected)

    def test_cell_search_passes_only_excel_files_to_its_reader(self):
        read_paths = []

        def read_cells(path, *_args):
            read_paths.append(path)
            return []

        with patch("cell_search._search_excel_file_cells", side_effect=read_cells):
            results = cell_search.search_excel_cells(self.root, "needle", max_results=None)

        self.assertEqual(results, [])
        self.assertEqual(set(read_paths), self.expected)
        self.assertEqual(len(read_paths), len(self.expected))

    def test_worksheet_search_passes_only_excel_files_to_its_reader_without_user_cache_access(self):
        read_paths = []

        def read_sheets(path):
            read_paths.append(path)
            return ["Target"]

        with (
            patch("worksheet_search._get_cached_sheet_names", side_effect=read_sheets),
            patch("worksheet_search._flush_sheet_name_cache") as flush,
            patch("worksheet_search.read_cache_data", side_effect=AssertionError("No user cache reads")),
            patch("worksheet_search.update_cache_data", side_effect=AssertionError("No user cache writes")),
        ):
            results = worksheet_search.scan_excel_workbooks(self.root, "target")

        self.assertEqual(set(read_paths), self.expected)
        self.assertEqual(len(read_paths), len(self.expected))
        self.assertEqual({item.path for item in results}, self.expected)
        self.assertTrue(all(item.sheet_names == ["Target"] and item.error is None for item in results))
        flush.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
