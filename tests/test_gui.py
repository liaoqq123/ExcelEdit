"""需要图形环境的 Tk 回归测试。"""

import unittest
import tempfile
import time
from pathlib import Path
from types import SimpleNamespace
from threading import Event
from tkinter import TclError
from unittest.mock import patch

from cell_search import ExcelCellMatch
from cache_store import read_cache_data, update_cache_data
from file_search import FileSearchResult
from gui import ExcelEditApp, format_excel_cell_address, load_cache_data as load_gui_cache_data, save_cache_data as save_gui_cache_data
from json_validation import JsonValidationError, JsonValidationReport
from worksheet_search import DEFAULT_REFERENCE_LOOKUP_CONFIG, WORKSHEET_SHEET_CACHE_KEY, ExcelSheetReference, ExcelWorkbookInfo
from search_status import SearchTracker
from PIL import Image


class GuiRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.cache_reader = patch("gui.load_cache_data", return_value={})
        cls.cache_writer = patch("gui.save_cache_data", return_value=True)
        cls.cache_reader.start()
        cls.cache_writer.start()
        cls.addClassCleanup(cls.cache_reader.stop)
        cls.addClassCleanup(cls.cache_writer.stop)
        try:
            cls.app = ExcelEditApp()
        except TclError as exc:
            raise unittest.SkipTest(f"当前环境没有可用图形界面：{exc}") from exc
        cls.app.withdraw()

    def setUp(self) -> None:
        self.app._load_cached_inputs()
        self.app.search_mode_var.set("工作表检索")
        self.app.file_result_view_var.set("缩略图")
        self.app.search_tracker = None
        self.app.search_cancel_event = None
        self.app._set_search_controls_running(False)
        self.app.folder_path_var.set("")
        self.app.keyword_entry.delete(0, "end")
        self.app.failure_details_button.grid_remove()

    @classmethod
    def tearDownClass(cls) -> None:
        if hasattr(cls, "app") and not cls.app.is_closing:
            cls.app._save_cache_from_inputs = lambda: None
            cls.app.close_app()

    def tearDown(self) -> None:
        if not self.app.is_closing:
            self.app._clear_results()
            self.app.withdraw()

    def test_large_file_view_only_draws_visible_cards_and_can_reach_last_result(self) -> None:
        results = [FileSearchResult(Path(f"E:/data/{i}.txt"), f"{i}.txt", f"E:/data/{i}.txt") for i in range(20000)]
        self.app._render_file_results(results)
        view = self.app.file_thumbnail_frame
        self.assertLess(len(view.cards), 64)
        self.assertLess(len(view.canvas.find_all()), 600)
        self.assertEqual(len(view.results), 20000)
        view.canvas.yview_moveto(1)
        view._draw_visible()
        self.assertIn(results[-1], view.cards)
        self.assertLess(len(view.cards), 64)
        with patch.object(view, "open_result") as open_result:
            view._on_key(SimpleNamespace(keysym="End"))
            view._on_key(SimpleNamespace(keysym="Return"))
            open_result.assert_called_once_with(results[-1])

    def test_async_image_loading_and_stale_image_discard(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "preview.png"
            Image.new("RGB", (128, 96), "red").save(path)
            self.app._render_file_results([FileSearchResult(path, path.name, str(path))])
            view = self.app.file_thumbnail_frame
            old_generation = view.generation
            deadline = time.monotonic() + 3
            while path not in view._cache and time.monotonic() < deadline:
                self.app.update()
                time.sleep(0.01)
            self.assertIsNotNone(view._cache.get(path))
            self.assertLessEqual(max(view._cache[path].size), 256)
            view.clear()
            view._loaded.put_nowait((old_generation, path, Image.new("RGBA", (1, 1))))
            deadline = time.monotonic() + 1
            while not view._loaded.empty() and time.monotonic() < deadline:
                self.app.update()
                time.sleep(0.01)
            self.assertNotIn(path, view._cache)
            self.assertEqual(view.cards, [])

    def test_new_view_cancels_pending_table_batches(self) -> None:
        results = [FileSearchResult(Path(f"E:/data/{i}.txt"), f"{i}.txt", "") for i in range(2000)]
        self.app.file_result_view_var.set("表格")
        self.app._render_file_results(results)
        self.assertIsNotNone(self.app.tree_render_job)
        self.app._clear_results()
        self.app.update()
        self.assertEqual(self.app.result_tree.get_children(), ())
        self.assertIsNone(self.app.tree_render_job)

    def test_old_progress_is_discarded_and_limit_is_explained(self) -> None:
        old = self.app._begin_search()
        old_tracker = self.app.search_tracker
        self.app._post_search_callback(old, lambda: self.app.status_label.configure(text="stale"))
        current = self.app._begin_search()
        self.app.status_label.configure(text="new")
        if self.app.ui_poll_job is not None:
            self.app.after_cancel(self.app.ui_poll_job)
        self.app._poll_ui_callbacks()
        self.assertEqual(self.app.status_label.cget("text"), "new")
        self.app.search_tracker.finish("limit")
        self.assertIn("结果可能不完整", self.app._completion_status("完成"))
        self.assertIs(self.app.search_cancel_event, current)
        self.assertIsNot(self.app.search_tracker, old_tracker)
        self.app.search_cancel_event = None
        self.app._set_search_controls_running(False)

    def test_reference_buttons_are_limited_to_visible_rows(self) -> None:
        self.app._render_results([ExcelWorkbookInfo(Path("demo.xlsx"), "demo.xlsx", [f"Sheet{i}" for i in range(1000)])])
        self.app.deiconify()
        deadline = time.monotonic() + 3
        while self.app.tree_render_job is not None and time.monotonic() < deadline:
            self.app.update()
        self.app.update_idletasks()
        self.app._layout_action_buttons()
        self.assertEqual(len(self.app.action_item_ids), 1000)
        self.assertGreater(len(self.app.action_buttons), 0)
        self.assertLess(len(self.app.action_buttons), 30)
        self.app.result_tree.yview_moveto(1)
        self.app._layout_action_buttons()
        self.assertIn(self.app.result_tree.get_children()[-1], self.app.action_buttons)

    def test_action_button_layout_uses_widget_configuration(self) -> None:
        self.app._render_results(
            [ExcelWorkbookInfo(Path("demo.xlsx"), "demo.xlsx", ["Sheet1"])]
        )
        self.assertEqual(self.app.result_tree.winfo_manager(), "grid")
        self.assertEqual(self.app.file_thumbnail_frame.winfo_manager(), "")
        self.assertEqual(str(self.app.result_tree.heading("workbook", "anchor")), "w")
        self.assertEqual(str(self.app.result_tree.column("workbook", "anchor")), "w")
        self.app.update_idletasks()
        with patch.object(self.app, "_get_visible_result_item_ids", return_value=list(self.app.action_item_ids)), patch.object(self.app.result_tree, "bbox", return_value=(0, 0, 132, 30)):
            self.app._layout_action_buttons()
        button = next(iter(self.app.action_buttons.values()))
        self.assertGreaterEqual(int(button.cget("width")), 86)
        self.assertGreaterEqual(int(button.cget("height")), 20)
        source_item = next(iter(self.app.action_item_ids))
        action_index = self.app.current_result_columns.index("action")
        values = self.app.result_tree.item(source_item, "values")
        self.assertEqual(values[action_index], "")
        self.assertTrue(self.app._action_button_is_available(source_item))

    def test_stale_reference_callback_is_discarded(self) -> None:
        self.app._render_results(
            [ExcelWorkbookInfo(Path("demo.xlsx"), "demo.xlsx", ["Sheet1"])]
        )
        source_item = next(iter(self.app.action_item_ids))
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
        self.assertFalse(self.app.folder_dropdown_active)

    def test_folder_dropdown_defers_history_refresh_until_popup_closes(self) -> None:
        first_path = r"D:\data\one"
        second_path = r"D:\data\two"
        self.app.folder_history = [first_path, second_path]
        self.app._refresh_folder_history_menu()
        dropdown_menu = self.app.folder_entry._dropdown_menu
        refresh_states: list[bool] = []
        original_configure = self.app.folder_entry.configure

        def record_configure(*args: object, **kwargs: object) -> object:
            if "values" in kwargs:
                refresh_states.append(self.app.folder_dropdown_active)
            return original_configure(*args, **kwargs)

        def select_while_popup_is_active(*_args: object) -> None:
            self.assertTrue(self.app.folder_dropdown_active)
            self.app._select_folder_from_history(second_path)
            self.assertTrue(self.app.folder_history_refresh_pending)
            self.assertFalse(refresh_states)

        with (
            patch.object(dropdown_menu, "tk_popup", side_effect=select_while_popup_is_active),
            patch.object(self.app.folder_entry, "configure", side_effect=record_configure),
            patch("gui.save_cache_data", return_value=True),
        ):
            self.app._open_folder_dropdown_menu()

        self.assertEqual(self.app.folder_history[0], second_path)
        self.assertEqual(refresh_states, [False])
        self.assertFalse(self.app.folder_dropdown_active)
        self.assertFalse(self.app.folder_history_refresh_pending)

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
        self.assertEqual(sum(title is not None for _, title, _ in self.app.file_thumbnail_frame.rows), 2)

        self.app.search_mode_var.set("文件检索")
        self.app.file_result_view_var.set("表格")
        self.app._on_file_result_view_changed("表格")
        self.assertEqual(self.app.result_tree.winfo_manager(), "grid")
        self.assertEqual(self.app.file_thumbnail_frame.winfo_manager(), "")
        self.assertEqual(self.app.result_tree.heading("#0", "text"), "文件夹")
        self.assertEqual(str(self.app.result_tree.heading("file_name", "anchor")), "w")
        self.assertEqual(str(self.app.result_tree.column("file_name", "anchor")), "w")
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

        self.assertEqual(self.app.result_tree.heading("#0", "text"), "工作簿 / 工作表 / 位置")
        self.assertEqual(self.app.current_result_columns, ("value",))
        workbook_items = self.app.result_tree.get_children()
        self.assertEqual(len(workbook_items), 1)
        self.assertEqual(self.app.result_tree.item(workbook_items[0], "tags"), ("cell-workbook",))
        self.assertEqual(self.app.item_targets[workbook_items[0]].file_path, workbook_path)
        self.assertIsNone(self.app.item_targets[workbook_items[0]].sheet_name)
        sheet_items = self.app.result_tree.get_children(workbook_items[0])
        self.assertEqual([self.app.result_tree.item(item_id, "text") for item_id in sheet_items], ["任务", "奖励"])
        self.assertEqual(self.app.result_tree.item(sheet_items[0], "tags"), ("cell-sheet",))
        task_cells = self.app.result_tree.get_children(sheet_items[0])
        self.assertEqual(
            [self.app.result_tree.item(item_id, "text") for item_id in task_cells],
            ["A21", "AB7"],
        )
        self.assertEqual(format_excel_cell_address(8, 27), "AA8")

    def test_copy_cell_keeps_multiline_content_and_escapes_sheet_address(self) -> None:
        path = Path(r"D:\data\中文配置.xlsx")
        original = "首行\n第二行\t保留制表符\r\n=1+1"
        self.app._render_cell_matches([
            ExcelCellMatch(path, path.name, "角色's表", 12, 28, original),
        ])
        cell_item = next(item for item, target in self.app.item_targets.items() if target.row_index)
        self.app.result_tree.selection_set(cell_item)
        with patch.object(self.app, "_update_selected_preview"):
            self.app._on_result_selected()
        with patch.object(self.app, "clipboard_clear") as clear, patch.object(self.app, "clipboard_append") as append:
            self.assertEqual(self.app._copy_selected(), "break")
            append.assert_called_once_with(original)
            clear.assert_called_once()
            append.reset_mock()
            self.app._copy_selected("address")
            append.assert_called_once_with("'角色''s表'!AB12")
            append.reset_mock()
            self.app._copy_selected("path")
            append.assert_called_once_with(str(path))

    def test_local_filter_preserves_full_source_and_restores_grouped_cells(self) -> None:
        first = Path(r"D:\data\第一配置.xlsx")
        second = Path(r"D:\data\第二配置.xlsx")
        matches = [
            ExcelCellMatch(first, first.name, "技能", 12, 2, "Alpha\n完整内容"),
            ExcelCellMatch(first, first.name, "奖励", 9, 3, "Beta"),
            ExcelCellMatch(second, second.name, "技能", 8, 4, "Gamma"),
        ]
        self.app._render_cell_matches(matches)
        self.app.local_filter_var.set("alpha")
        self.app._apply_local_filter()

        self.assertEqual(self.app.result_source, ("cells", matches))
        workbooks = self.app.result_tree.get_children()
        self.assertEqual(len(workbooks), 1)
        sheets = self.app.result_tree.get_children(workbooks[0])
        self.assertEqual([self.app.result_tree.item(item, "text") for item in sheets], ["技能"])
        cells = self.app.result_tree.get_children(sheets[0])
        self.assertEqual([self.app.result_tree.item(item, "text") for item in cells], ["B12"])
        self.assertEqual(self.app.item_raw_values[cells[0]], "Alpha\n完整内容")

        self.app.local_filter_var.set("")
        self.app._apply_local_filter()
        self.assertEqual(self.app.result_source, ("cells", matches))
        workbooks = self.app.result_tree.get_children()
        self.assertEqual(len(workbooks), 2)
        visible_sheets = [sheet for workbook in workbooks for sheet in self.app.result_tree.get_children(workbook)]
        visible_cells = [cell for sheet in visible_sheets for cell in self.app.result_tree.get_children(sheet)]
        self.assertEqual(len(visible_sheets), 3)
        self.assertEqual(len(visible_cells), 3)
        self.assertEqual(set(self.app.item_raw_values.values()), {match.value for match in matches})

    def test_sheet_scope_uses_workbook_identity_and_combines_with_text_filter(self) -> None:
        first = Path(r"D:\data\one.xlsx")
        second = Path(r"D:\data\two.xlsx")
        matches = [
            ExcelCellMatch(first, first.name, "技能", 3, 1, "Alpha"),
            ExcelCellMatch(first, first.name, "技能", 4, 1, "Beta"),
            ExcelCellMatch(first, first.name, "奖励", 5, 1, "Alpha"),
            ExcelCellMatch(second, second.name, "技能", 6, 1, "Alpha"),
        ]
        self.app._render_cell_matches(matches)
        chosen = next(target for target in self.app.item_targets.values()
                      if target.file_path == first and target.sheet_name == "技能" and target.row_index == 3)
        self.app.local_filter_var.set("alpha")
        self.app._scope_to_sheet(chosen)
        visible_cells = [target for target in self.app.item_targets.values() if target.row_index]
        self.assertEqual(visible_cells, [chosen])
        self.assertEqual(self.app.result_source, ("cells", matches))

        self.app._reset_local_filter()
        self.app._apply_local_filter()
        self.assertIsNone(self.app.local_sheet_scope)
        self.assertEqual(len([target for target in self.app.item_targets.values() if target.row_index]), 4)

    def test_file_view_switch_uses_displayed_results_after_live_search_mode_changes(self) -> None:
        folder = Path(r"D:\data\images")
        results = [FileSearchResult(folder / "图.png", "图.png", str(folder / "图.png"))]
        self.app._render_file_results(results)
        self.app._on_search_mode_changed("单元格检索")
        self.app.file_result_view_var.set("表格")
        self.app._on_file_result_view_changed("表格")
        self.assertEqual(self.app.result_tree.winfo_manager(), "grid")
        self.assertEqual(self.app.file_thumbnail_frame.winfo_manager(), "")
        folder_item = self.app.result_tree.get_children()[0]
        file_item = self.app.result_tree.get_children(folder_item)[0]
        self.assertEqual(self.app.item_targets[file_item].file_path, results[0].path)
        self.assertFalse(self.app.item_targets[file_item].open_containing_folder)

        self.app.file_result_view_var.set("缩略图")
        self.app._on_file_result_view_changed("缩略图")
        self.assertEqual(self.app.file_thumbnail_frame.winfo_manager(), "grid")
        self.assertEqual(self.app.file_thumbnail_frame.results, results)
        self.assertEqual(self.app.result_source, ("files", results))

    def test_removed_selection_cannot_copy_previous_cell(self) -> None:
        path = Path(r"D:\data\demo.xlsx")
        self.app._render_cell_matches([ExcelCellMatch(path, path.name, "Sheet1", 1, 1, "previous")])
        cell_item = next(item for item, target in self.app.item_targets.items() if target.row_index)
        self.app.result_tree.selection_set(cell_item)
        with patch.object(self.app, "_update_selected_preview"):
            self.app._on_result_selected()
        self.app.result_tree.selection_remove(cell_item)
        self.app._on_result_selected()
        with patch.object(self.app, "clipboard_clear") as clear, patch.object(self.app, "clipboard_append") as append:
            self.app._copy_selected()
            clear.assert_not_called()
            append.assert_not_called()
        self.assertIsNone(self.app.selected_target)

    def test_cleared_results_cannot_copy_previous_file(self) -> None:
        path = Path(r"D:\data\demo.txt")
        result = FileSearchResult(path, path.name, str(path))
        with patch.object(self.app, "_update_selected_preview"):
            self.app._on_file_selected(result)
        self.app._clear_results()
        with patch.object(self.app, "clipboard_clear") as clear, patch.object(self.app, "clipboard_append") as append:
            self.app._copy_selected("path")
            clear.assert_not_called()
            append.assert_not_called()
        self.assertIsNone(self.app.selected_target)

    def test_cancelled_workers_keep_partial_results_for_each_search_mode(self) -> None:
        path = Path(r"D:\data\partial.xlsx")
        cells = [ExcelCellMatch(path, path.name, "技能", 7, 3, "已找到")]
        files = [FileSearchResult(path, path.name, str(path))]
        sheets = [ExcelWorkbookInfo(path, path.name, ["技能"])]
        report = JsonValidationReport(
            errors=[JsonValidationError(path, path.name, "技能", 7, 3, "{错误}", "无效 JSON")],
            checked_cell_count=8,
        )
        cases = [
            ("cells", "gui.search_excel_cells", cells,
             lambda event: self.app._search_cells_in_background(path.parent, "已找到", None, None, False, event)),
            ("files", "gui.search_files", files,
             lambda event: self.app._search_files_in_background(path.parent, "partial", (), False, False, event)),
            ("sheets", "gui.scan_excel_workbooks", sheets,
             lambda event: self.app._scan_in_background(path.parent, "技能", False, event)),
            ("json", "gui.validate_excel_json_columns", report,
             lambda event: self.app._validate_json_in_background(path.parent, event)),
        ]
        for kind, function_name, payload, run_worker in cases:
            with self.subTest(mode=kind):
                self.app._clear_results()
                event = self.app._begin_search()

                def cancel_after_partial_result(*args, **kwargs):
                    kwargs["cancel_event"].set()
                    return payload

                with patch(function_name, side_effect=cancel_after_partial_result):
                    run_worker(event)
                if self.app.ui_poll_job is not None:
                    self.app.after_cancel(self.app.ui_poll_job)
                self.app._poll_ui_callbacks()

                self.assertIsNone(self.app.search_cancel_event)
                self.assertEqual(self.app.result_source, (kind, payload))
                self.assertEqual(self.app.search_tracker.snapshot().stop_reason, "cancelled")
                self.assertIn("已取消", self.app.status_label.cget("text"))
                self.assertIn("结果可能不完整", self.app.status_label.cget("text"))

    def test_cached_view_preferences_are_applied_and_saved_with_inputs(self) -> None:
        preferences = {
            "pinned_folders": [r"D:\data\one", r"D:\data\two"],
            "preview_enabled": False,
            "result_density": "紧凑",
            "file_result_view": "表格",
        }
        with patch("gui.load_cache_data", return_value=preferences):
            self.app._load_cached_inputs()
        self.assertEqual(self.app.pinned_folders, preferences["pinned_folders"])
        self.assertFalse(self.app.preview_enabled)
        self.assertEqual(self.app.preview_pane.winfo_manager(), "")
        self.assertEqual(self.app.density_var.get(), "紧凑")
        self.assertEqual(self.app.file_result_view_var.get(), "表格")
        with patch("gui.save_cache_data", return_value=True) as save:
            self.app._save_cache_from_inputs()
        saved = save.call_args.args[0]
        for key, value in preferences.items():
            self.assertEqual(saved[key], value)

    def test_search_settings_are_accessible_and_applied(self) -> None:
        with patch.object(self.app, "_save_cache_from_inputs"):
            self.assertEqual(self.app.search_settings_button.cget("text"), "＋ 更多条件")
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


class CachePreferenceTests(unittest.TestCase):
    def test_view_preference_round_trip_preserves_shared_sheet_index_cache(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            cache_file = Path(directory) / "cache.json"
            index_cache = {"existing.xlsx": {"sheets": ["技能"]}}
            update_cache_data({WORKSHEET_SHEET_CACHE_KEY: index_cache}, cache_file, None)
            preferences = {
                "pinned_folders": [r"D:\data\one", r"D:\data\one", r"D:\data\two"],
                "preview_enabled": False,
                "result_density": "紧凑",
                "file_result_view": "表格",
            }
            with patch("gui.update_shared_cache_data", side_effect=lambda updates: update_cache_data(updates, cache_file, None)):
                self.assertTrue(save_gui_cache_data(preferences))
            with patch("gui.read_cache_file", side_effect=lambda: read_cache_data(cache_file, None)):
                loaded = load_gui_cache_data()

            self.assertEqual(loaded["pinned_folders"], [r"D:\data\one", r"D:\data\two"])
            self.assertFalse(loaded["preview_enabled"])
            self.assertEqual(loaded["result_density"], "紧凑")
            self.assertEqual(loaded["file_result_view"], "表格")
            self.assertEqual(read_cache_data(cache_file, None)[WORKSHEET_SHEET_CACHE_KEY], index_cache)


if __name__ == "__main__":
    unittest.main()
