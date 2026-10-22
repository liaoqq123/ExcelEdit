"""只读预览的读取边界、资源回收与异步选择回归测试。"""

from contextlib import contextmanager
from pathlib import Path
import tempfile
from threading import Event, get_ident
import time
from tkinter import TclError
import unittest
from unittest.mock import patch

import customtkinter as ctk
from PIL import Image
import xlsxwriter

from result_preview import (
    CellPreviewData, FilePreviewData, PreviewCancelled, PreviewRow, ResultPreviewPane,
    _read_cell_preview, _read_file_preview,
)


class PreviewReaderTests(unittest.TestCase):
    def test_xlsx_reads_only_five_surrounding_rows_and_columns_and_header(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "preview.xlsx"
            workbook = xlsxwriter.Workbook(path)
            sheet = workbook.add_worksheet("技能")
            for column in range(8):
                sheet.write(0, column, f"Field{column + 1}")
            for row in range(1, 20):
                for column in range(8):
                    sheet.write(row, column, f"R{row + 1}C{column + 1}")
            workbook.close()
            data = _read_cell_preview(path, "技能", 12, 5, Event())
            self.assertEqual(data.first_column, 3)
            self.assertEqual(data.field_name, "Field5")
            self.assertEqual([row.index for row in data.rows], [10, 11, 12, 13, 14])
            self.assertTrue(all(len(row.values) == 5 for row in data.rows))
            self.assertEqual(data.rows[2].values[2], "R12C5")
            # Opening for write succeeds because the read-only ZIP has been closed.
            path.rename(path.with_suffix(".renamed.xlsx"))

    def test_xlsx_cancellation_closes_iterator_and_workbook(self) -> None:
        cancel = Event()
        closed = []

        def rows(**kwargs):
            try:
                yield ("Name",)
                cancel.set()
                yield ("value",)
            finally:
                closed.append("rows")

        class Workbook:
            def __getitem__(self, name):
                return type("Worksheet", (), {"iter_rows": staticmethod(rows)})()

            def close(self):
                closed.append("workbook")

        with patch("result_preview.load_workbook", return_value=Workbook()):
            with self.assertRaises(PreviewCancelled):
                _read_cell_preview(Path("demo.xlsx"), "Data", 2, 1, cancel)
        self.assertEqual(closed, ["rows", "workbook"])

    def test_cancelled_request_never_opens_file(self) -> None:
        cancel = Event()
        cancel.set()
        with patch("result_preview.load_workbook") as reader:
            with self.assertRaises(PreviewCancelled):
                _read_cell_preview(Path("not-present.xlsx"), "Data", 1, 1, cancel)
        reader.assert_not_called()

    def test_reader_rejects_invalid_coordinates_before_opening(self) -> None:
        with patch("result_preview.load_workbook") as reader:
            for row, column in [(0, 1), (1, 0), (1_048_577, 1), (1, 16_385)]:
                with self.subTest(row=row, column=column), self.assertRaises(ValueError):
                    _read_cell_preview(Path("demo.xlsx"), "Data", row, column, Event())
        reader.assert_not_called()

    def test_xls_resource_is_released_when_sheet_is_missing(self) -> None:
        with patch("xlrd.open_workbook") as reader:
            workbook = reader.return_value
            workbook.sheet_by_name.side_effect = KeyError("missing")
            with self.assertRaises(KeyError):
                _read_cell_preview(Path("demo.xls"), "missing", 2, 1, Event())
            workbook.release_resources.assert_called_once()

    def test_xlsb_sparse_rows_preserve_real_column_positions(self) -> None:
        from types import SimpleNamespace

        @contextmanager
        def sheet_context():
            def rows(sparse=True):
                yield [SimpleNamespace(r=0, c=3, v="FieldD"), SimpleNamespace(r=0, c=5, v="FieldF")]
                yield [SimpleNamespace(r=9, c=3, v="selected"), SimpleNamespace(r=9, c=5, v=42)]
                yield [SimpleNamespace(r=999, c=3, v="far outside preview")]
            yield SimpleNamespace(rows=rows)

        @contextmanager
        def workbook_context(path):
            yield SimpleNamespace(get_sheet=lambda name: sheet_context())

        with patch("pyxlsb.open_workbook", workbook_context):
            data = _read_cell_preview(Path("demo.xlsb"), "Data", 10, 4, Event())
        self.assertEqual(data.first_column, 2)
        self.assertEqual(data.field_name, "FieldD")
        self.assertEqual(data.rows[0].index, 10)
        self.assertEqual(data.rows[0].values, (None, None, "selected", None, 42))

    def test_file_preview_decodes_bounded_image_and_releases_source(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "preview.png"
            Image.new("RGB", (1800, 1200), "#3468e4").save(path)
            data = _read_file_preview(path, Event())
            self.assertEqual(data.dimensions, (1800, 1200))
            self.assertLessEqual(data.image.width, 580)
            self.assertLessEqual(data.image.height, 420)
            self.assertGreater(data.size, 0)
            self.assertFalse(data.is_directory)
            self.assertEqual(data.image.getpixel((0, 0)), (52, 104, 228, 255))
            path.unlink()
            # The returned pixels are independent of the now deleted source file.
            self.assertEqual(data.image.getpixel((0, 0)), (52, 104, 228, 255))

    def test_corrupt_image_still_has_file_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "broken.png"
            path.write_text("invalid")
            data = _read_file_preview(path, Event())
            self.assertIsNone(data.image)
            self.assertGreater(data.size, 0)
            self.assertIn("无法预览此图片", data.image_error)


class PreviewPaneTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        try:
            cls.root = ctk.CTk()
        except TclError as exc:
            raise unittest.SkipTest(f"当前环境没有可用图形界面：{exc}") from exc
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls) -> None:
        if hasattr(cls, "root"):
            for job in cls.root.tk.call("after", "info"):
                cls.root.tk.call("after", "cancel", job)
            cls.root.destroy()

    def setUp(self) -> None:
        self.pane = ResultPreviewPane(self.root, lambda: None, lambda: None, lambda: None)

    def tearDown(self) -> None:
        self.pane.destroy()

    def pump_until(self, condition, seconds=3) -> None:
        deadline = time.monotonic() + seconds
        while not condition() and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertTrue(condition())

    def test_debounce_reads_only_final_selection_on_background_thread(self) -> None:
        calls = []
        main_thread = get_ident()

        def read(path, sheet, row, column, cancel):
            calls.append((path, get_ident()))
            return CellPreviewData(1, ("Name",), (PreviewRow(row, ("value",)),), "Name")

        with patch("result_preview._read_cell_preview", side_effect=read):
            self.pane.show_cell(Path("old.xlsx"), "Data", 12, 1, "first")
            self.pane.show_cell(Path("new.xlsx"), "Data", 20, 1, "last")
            self.pump_until(lambda: "Name" in self.pane.field_label.cget("text"))
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][0], Path("new.xlsx"))
        self.assertNotEqual(calls[0][1], main_thread)
        self.assertEqual(self.pane.content_box.get("1.0", "end-1c"), "last")
        self.assertEqual(self.pane.nearby_tree.item(self.pane.nearby_tree.get_children()[0], "text"), "20")

    def test_switch_during_read_discards_old_result_and_clear_discards_pending_result(self) -> None:
        started, release = Event(), Event()

        def read(path, cancel):
            if path.name == "old.txt":
                started.set()
                release.wait(2)
                return FilePreviewData(False, 123, "OLD")
            return FilePreviewData(False, 456, "NEW")

        with patch("result_preview._read_file_preview", side_effect=read):
            self.pane.show_file(Path("old.txt"))
            self.pump_until(started.is_set)
            self.pane.show_file(Path("new.txt"))
            release.set()
            self.pump_until(lambda: "NEW" in self.pane.file_metadata_label.cget("text"))
            self.assertNotIn("OLD", self.pane.file_metadata_label.cget("text"))
            old_generation = self.pane._generation
            self.pane.clear()
            self.pane._responses.put_nowait((old_generation, "file", FilePreviewData(False, 999, "STALE")))
            self.pump_until(self.pane._responses.empty)
        self.assertIsNone(self.pane._selection_key)
        self.assertEqual(self.pane.open_button.cget("state"), "disabled")
        self.assertNotIn("STALE", self.pane.file_metadata_label.cget("text"))


if __name__ == "__main__":
    unittest.main()
