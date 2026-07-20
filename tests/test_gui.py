"""需要图形环境的 Tk 回归测试。"""

import unittest
from pathlib import Path
from threading import Event
from tkinter import TclError
from unittest.mock import patch

from gui import ExcelEditApp
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


if __name__ == "__main__":
    unittest.main()
