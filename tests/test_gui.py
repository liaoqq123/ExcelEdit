"""需要图形环境的 Tk 回归测试。"""

import unittest
from pathlib import Path
from threading import Event
from tkinter import TclError
from unittest.mock import patch

from cell_search import ExcelCellMatch
from file_search import FileSearchResult
from gui import ExcelEditApp, format_excel_cell_address
from worksheet_search import DEFAULT_REFERENCE_LOOKUP_CONFIG, ExcelSheetReference, ExcelWorkbookInfo


class GuiRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        try:
            cls.app = ExcelEditApp()
        except TclError as exc:
            raise unittest.SkipTest(f"当前环境没有可用图形界面：{exc}") from exc
        cls.app.withdraw()

    @classmethod
    def tearDownClass(cls) -> None:
        if hasattr(cls, "app") and not cls.app.is_closing:
            cls.app._save_cache_from_inputs = lambda: None
            cls.app.close_app()

    def tearDown(self) -> None:
        if not self.app.is_closing:
            self.app._clear_results()

    def test_action_button_layout_uses_widget_configuration(self) -> None:
        self.app._render_results(
            [ExcelWorkbookInfo(Path("demo.xlsx"), "demo.xlsx", ["Sheet1"])]
        )
        self.assertEqual(self.app.result_tree.winfo_manager(), "grid")
        self.assertEqual(self.app.file_thumbnail_frame.winfo_manager(), "")
        self.assertEqual(self.app.result_tree.heading("workbook", "anchor"), "w")
        self.assertEqual(self.app.result_tree.column("workbook", "anchor"), "w")
        self.app.update_idletasks()
        self.app._layout_action_buttons()
        button = next(iter(self.app.action_buttons.values()))
        self.assertGreaterEqual(int(button.cget("width")), 86)
        self.assertGreaterEqual(int(button.cget("height")), 20)
        source_item = next(iter(self.app.action_buttons))
        action_index = self.app.current_result_columns.index("action")
        values = self.app.result_tree.item(source_item, "values")
        self.assertEqual(values[action_index], "")
        self.assertTrue(self.app._action_button_is_available(source_item))

    def test_stale_reference_callback_is_discarded(self) -> None:
        self.app._render_results(
            [ExcelWorkbookInfo(Path("demo.xlsx"), "demo.xlsx", ["Sheet1"])]
        )
        source_item = next(iter(self.app.action_buttons))
        generation = self.app.result_generation
        cancel_event = Event()
        self.app.reference_cancel_events[source_item] = cancel_event

        self.app._clear_results()
        self.app._references_finished(
            source_item,
            Path("demo.xlsx"),
            [ExcelSheetReference("Target", "Id")],
            DEFAULT_REFERENCE_LOOKUP_CONFIG,
            generation,
            cancel_event,
        )
        self.assertFalse(self.app.result_tree.get_children())

    def test_folder_dropdown_uses_popup_grab_instead_of_plain_post(self) -> None:
        self.app.folder_history = [r"D:\data\one", r"D:\data\two"]
        self.app._refresh_folder_history_menu()
        dropdown_menu = self.app.folder_entry._dropdown_menu

        with (
            patch.object(dropdown_menu, "tk_popup") as popup,
            patch.object(dropdown_menu, "open") as plain_open,
        ):
            self.app._open_folder_dropdown_menu()

        popup.assert_called_once()
        plain_open.assert_not_called()
        self.assertFalse(self.app.folder_entry._close_on_next_click)

    def test_file_results_can_switch_between_grouped_thumbnail_and_table_views(self) -> None:
        first_folder = Path(r"D:\data\reports")
        second_folder = Path(r"D:\data\exports")
        results = [
            FileSearchResult(first_folder / "one.xlsx", "one.xlsx", str(first_folder / "one.xlsx")),
            FileSearchResult(first_folder / "two.xlsx", "two.xlsx", str(first_folder / "two.xlsx")),
            FileSearchResult(second_folder / "three.xlsx", "three.xlsx", str(second_folder / "three.xlsx")),
        ]

        self.app._render_file_results(results)

        self.assertEqual(self.app.result_tree.winfo_manager(), "")
        self.assertEqual(self.app.file_thumbnail_frame.winfo_manager(), "grid")
        self.assertEqual(self.app.file_result_view_controls.winfo_manager(), "grid")
        self.assertEqual(len(self.app.file_thumbnail_cards), 3)
        self.assertEqual(self.app.file_result_count_label.cget("text"), "文件检索结果：3 项")
        self.assertEqual(self.app.result_group_controls.winfo_manager(), "")
        self.assertEqual(self.app._get_file_thumbnail_style(results[0])[2], "Excel / 表格文件")
        self.assertEqual(len(self.app.file_thumbnail_frame.winfo_children()), 4)

        self.app.search_mode_var.set("文件检索")
        self.app.file_result_view_var.set("表格")
        self.app._on_file_result_view_changed("表格")
        self.assertEqual(self.app.result_tree.winfo_manager(), "grid")
        self.assertEqual(self.app.file_thumbnail_frame.winfo_manager(), "")
        self.assertEqual(self.app.result_tree.heading("#0", "text"), "文件夹")
        self.assertEqual(self.app.result_tree.heading("file_name", "anchor"), "w")
        self.assertEqual(self.app.result_tree.column("file_name", "anchor"), "w")
        folder_items = self.app.result_tree.get_children()
        self.assertEqual(len(folder_items), 2)
        self.assertEqual(len(self.app.result_tree.get_children(folder_items[0])), 2)

        self.app._clear_results()
        self.assertEqual(self.app.result_group_controls.winfo_manager(), "")
        self.assertEqual(self.app.file_thumbnail_results, [])

    def test_cell_results_use_workbook_sheet_directories_and_excel_addresses(self) -> None:
        workbook_path = Path(r"D:\data\activity.xlsx")
        self.app._render_cell_matches(
            [
                ExcelCellMatch(workbook_path, "activity.xlsx", "任务", 21, 1, "first"),
                ExcelCellMatch(workbook_path, "activity.xlsx", "任务", 7, 28, "second"),
                ExcelCellMatch(workbook_path, "activity.xlsx", "奖励", 3, 2, "third"),
            ]
        )

        self.assertEqual(self.app.result_tree.heading("#0", "text"), "检索目录")
        self.assertEqual(self.app.current_result_columns, ("value", "action"))
        workbook_items = self.app.result_tree.get_children()
        self.assertEqual(len(workbook_items), 1)
        self.assertEqual(self.app.result_tree.item(workbook_items[0], "tags"), ("cell-workbook",))
        self.assertIn(workbook_items[0], self.app.directory_buttons)
        sheet_items = self.app.result_tree.get_children(workbook_items[0])
        self.assertEqual([self.app.result_tree.item(item_id, "text") for item_id in sheet_items], ["任务", "奖励"])
        self.assertEqual(self.app.result_tree.item(sheet_items[0], "tags"), ("cell-sheet",))
        task_cells = self.app.result_tree.get_children(sheet_items[0])
        self.assertEqual(
            [self.app.result_tree.item(item_id, "text") for item_id in task_cells],
            ["A21", "AB7"],
        )
        self.assertEqual(format_excel_cell_address(8, 27), "AA8")

    def test_search_settings_are_accessible_and_applied(self) -> None:
        with patch.object(self.app, "_save_cache_from_inputs"):
            self.assertEqual(self.app.search_settings_button.cget("text"), "检索设置")
            with patch.object(self.app.search_settings_dialog, "open") as open_dialog:
                self.app.open_search_settings()
            open_dialog.assert_called_once_with("工作表检索", False, False, False, "", False, False)

            self.app._apply_search_settings(True, True, True, "xlsx, pdf", True, True)
            self.assertTrue(self.app.only_enabled_sheets_var.get())
            self.assertTrue(self.app.only_enabled_data_var.get())
            self.assertTrue(self.app.only_specified_suffixes_var.get())
            self.assertEqual(self.app.file_suffixes_var.get(), "xlsx, pdf")
            self.assertTrue(self.app.search_folders_var.get())
            self.assertTrue(self.app.exact_match_var.get())

            with patch.object(self.app.search_settings_dialog, "refresh") as refresh:
                self.app._on_search_mode_changed("文件检索")
            refresh.assert_called_once_with("文件检索")


if __name__ == "__main__":
    unittest.main()
