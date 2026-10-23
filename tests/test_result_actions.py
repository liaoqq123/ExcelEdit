"""统一打开操作的无窗口检查，仅使用假控件，不创建 Tk。"""

import os
from pathlib import Path
from types import MethodType, SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from file_search import FileSearchResult
from gui import ExcelEditApp, ResultTarget, get_open_action_label
from desktop_open import OpenResult


def fake_app(target=None):
    app = SimpleNamespace(
        selected_target=target, opening_targets=set(),
        selection_open_button=Mock(), selection_directory_button=Mock(),
        status_label=Mock(), selection_label=Mock(), preview_pane=Mock(),
    )
    app._sync_result_actions = MethodType(ExcelEditApp._sync_result_actions, app)
    return app


class ResultActionTests(unittest.TestCase):
    def test_labels_are_consistent_and_require_no_file_access(self):
        cases = [
            (None, "打开"),
            (ResultTarget(Path("book.XLSX")), "打开 Excel"),
            (ResultTarget(Path("book.xlsb"), "技能"), "在 Excel 中定位"),
            (ResultTarget(Path("book.xlsx"), "技能", 8, 3), "在 Excel 中定位"),
            (ResultTarget(Path("image.png")), "打开文件"),
            (ResultTarget(Path("子文件夹"), is_directory=True), "进入文件夹"),
            (ResultTarget(Path("文件夹.xlsx"), is_directory=True), "进入文件夹"),
        ]
        with patch.object(Path, "stat", side_effect=AssertionError("选择时不得读取磁盘")):
            for target, expected in cases:
                with self.subTest(target=target):
                    self.assertEqual(get_open_action_label(target), expected)

    def test_no_selection_disables_both_open_actions(self):
        app = fake_app()
        app._sync_result_actions()
        app.selection_open_button.configure.assert_called_once_with(text="打开", state="disabled")
        app.selection_directory_button.configure.assert_called_once_with(text="打开所在目录", state="disabled")

    def test_selection_enables_primary_and_directory_actions(self):
        app = fake_app(ResultTarget(Path("book.xlsx")))
        app._sync_result_actions()
        app.selection_open_button.configure.assert_called_once_with(text="打开 Excel", state="normal")
        app.selection_directory_button.configure.assert_called_once_with(text="打开所在目录", state="normal")

    def test_running_primary_is_disabled_without_blocking_directory(self):
        app = fake_app(ResultTarget(Path("data/book.xlsx")))
        app.opening_targets.add(ResultTarget(Path(os.path.abspath("data/book.xlsx"))))
        app._sync_result_actions()
        app.selection_open_button.configure.assert_called_once_with(text="打开中…", state="disabled")
        app.selection_directory_button.configure.assert_called_once_with(text="打开所在目录", state="normal")

    def test_running_directory_is_disabled_without_blocking_primary(self):
        app = fake_app(ResultTarget(Path("子文件夹"), is_directory=True))
        app.opening_targets.add(ResultTarget(Path(os.path.abspath("子文件夹")), open_containing_folder=True))
        app._sync_result_actions()
        app.selection_open_button.configure.assert_called_once_with(text="进入文件夹", state="normal")
        app.selection_directory_button.configure.assert_called_once_with(text="定位中…", state="disabled")

    def test_other_target_does_not_leave_new_selection_busy(self):
        app = fake_app(ResultTarget(Path("new.xlsx")))
        app.opening_targets.add(ResultTarget(Path(os.path.abspath("old.xlsx"))))
        app._sync_result_actions()
        self.assertEqual(app.selection_open_button.configure.call_args.kwargs["state"], "normal")

    def test_folder_primary_enters_folder_and_directory_reveals_it(self):
        target = ResultTarget(Path("子文件夹"), is_directory=True)
        app = fake_app(target)
        app._open_target = Mock()
        ExcelEditApp._open_selected_target(app)
        ExcelEditApp._open_selected_directory(app)
        self.assertEqual(app._open_target.call_args_list[0].args, (target,))
        self.assertEqual(app._open_target.call_args_list[1].args, (ResultTarget(target.file_path, open_containing_folder=True),))

    def test_gallery_selection_preserves_folder_metadata(self):
        app = fake_app()
        app._update_selected_preview = Mock()
        result = FileSearchResult(Path("folder.xlsx"), "folder.xlsx", "folder.xlsx", is_directory=True)
        ExcelEditApp._on_file_selected(app, result)
        self.assertTrue(app.selected_target.is_directory)
        self.assertEqual(app.selection_open_button.configure.call_args.kwargs["text"], "进入文件夹")

    def test_gallery_enter_selects_target_before_opening_without_duplicate_preview(self):
        app = fake_app()
        app._update_selected_preview = Mock()
        app._on_file_selected = MethodType(ExcelEditApp._on_file_selected, app)
        app._open_target = Mock()
        result = FileSearchResult(Path("folder.xlsx"), "folder.xlsx", "folder.xlsx", is_directory=True)
        ExcelEditApp._open_file_result(app, result)
        self.assertEqual(app.selected_target, ResultTarget(result.path, is_directory=True))
        self.assertEqual(app.selection_open_button.configure.call_args.kwargs["text"], "进入文件夹")
        app._open_target.assert_called_once_with(app.selected_target)
        ExcelEditApp._open_file_result(app, result)
        app._update_selected_preview.assert_called_once()

    def test_clearing_selection_resets_actions_and_stale_target(self):
        app = fake_app(ResultTarget(Path("old.xlsx")))
        app.result_tree = Mock()
        app.result_tree.selection.return_value = ()
        ExcelEditApp._on_result_selected(app)
        self.assertIsNone(app.selected_target)
        self.assertEqual(app.selection_open_button.configure.call_args.kwargs["state"], "disabled")
        self.assertEqual(app.selection_directory_button.configure.call_args.kwargs["state"], "disabled")
        app.preview_pane.clear.assert_called_once()

    def test_completion_restores_action_without_changing_selection(self):
        target = ResultTarget(Path(os.path.abspath("book.xlsx")))
        app = fake_app(target)
        app.opening_targets.add(target)
        ExcelEditApp._open_target_finished(app, target, OpenResult(True, True, True, 123))
        self.assertIs(app.selected_target, target)
        self.assertEqual(app.selection_open_button.configure.call_args.kwargs, {"text": "打开 Excel", "state": "normal"})

    def test_enter_opens_group_target_like_primary_button(self):
        app = fake_app()
        app.result_tree = Mock()
        app.result_tree.selection.return_value = ("workbook-group",)
        app.item_targets = {"workbook-group": ResultTarget(Path("book.xlsx"))}
        app._open_item = Mock()
        app._open_result_item = MethodType(ExcelEditApp._open_result_item, app)
        ExcelEditApp.open_selected_item(app)
        app._open_item.assert_called_once_with("workbook-group")

    def test_double_click_disclosure_arrow_does_not_open_file(self):
        app = fake_app()
        app.result_tree = Mock()
        app.result_tree.identify_element.return_value = "Treeitem.indicator"
        app._open_result_item = Mock()
        self.assertEqual(ExcelEditApp._handle_result_double_click(app, SimpleNamespace(x=5, y=20)), "break")
        app._open_result_item.assert_not_called()

    def test_context_menu_uses_same_folder_action_names(self):
        app = fake_app(ResultTarget(Path("子文件夹"), is_directory=True))
        app._copy_selected = Mock()
        app._open_selected_target = Mock()
        app._open_selected_directory = Mock()
        with patch("gui.Menu") as menu_class:
            ExcelEditApp._popup_result_menu(app, SimpleNamespace(x_root=10, y_root=10))
        labels = [call.kwargs["label"] for call in menu_class.return_value.add_command.call_args_list]
        self.assertEqual(labels.count("进入文件夹"), 1)
        self.assertEqual(labels.count("打开所在目录"), 1)


if __name__ == "__main__":
    unittest.main()
