"""Gallery interaction and background image loading contracts."""

from pathlib import Path
from tempfile import TemporaryDirectory
from time import monotonic, sleep
from types import SimpleNamespace
import unittest
from unittest.mock import Mock

import customtkinter as ctk
from PIL import Image

from file_result_view import CACHE_LIMIT, THUMBNAIL_SIZE, FileResultView
from file_search import FileSearchResult


class FileResultViewTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.window = ctk.CTk()
        cls.window.withdraw()

    @classmethod
    def tearDownClass(cls) -> None:
        for job in cls.window.tk.call("after", "info"):
            cls.window.tk.call("after", "cancel", job)
        cls.window.destroy()

    def setUp(self) -> None:
        self.open_result = Mock()
        self.on_select = Mock()
        self.on_copy = Mock()
        self.on_context_menu = Mock()
        self.view = FileResultView(
            self.window, self.open_result, lambda _result: ("□", "#345678", "文件"),
            on_select=self.on_select, on_copy=self.on_copy, on_context_menu=self.on_context_menu,
        )
        self.view.pack(fill="both", expand=True)

    def tearDown(self) -> None:
        self.view.destroy()

    def test_selection_context_copy_and_open_share_the_same_result(self) -> None:
        results = [FileSearchResult(Path(f"E:/gallery/{index}.png"), f"{index}.png", "") for index in range(3)]
        self.view.set_results(results)
        self.assertEqual(self.view.get_selected_result(), results[0])
        self.view._on_key(SimpleNamespace(keysym="Right"))
        self.on_select.assert_called_once_with(results[1])
        self.assertEqual(self.view.get_selected_result(), results[1])
        self.view._on_key(SimpleNamespace(keysym="c", state=4))
        self.on_copy.assert_called_once_with(results[1])
        self.view._on_key(SimpleNamespace(keysym="Return"))
        self.open_result.assert_called_once_with(results[1])
        # The whole tile selects; right-click uses the same target as preview.
        self.view._draw_visible()
        event = SimpleNamespace(x=25, y=65, x_root=25, y_root=65)
        self.view._on_context_menu(event)
        self.on_context_menu.assert_called_once_with(results[0], event)
        self.assertEqual(self.view.get_selected_result(), results[0])
        self.view._on_double_click(event)
        self.assertEqual(self.open_result.call_args.args, (results[0],))
        self.view.clear()
        self.assertIsNone(self.view.get_selected_result())

    def test_visible_images_load_original_dimensions_and_file_size_in_background(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "image.png"
            Image.new("RGB", (640, 480), "#ef8354").save(path)
            self.view.set_results([FileSearchResult(path, path.name, str(path))])
            deadline = monotonic() + 3
            while path not in self.view._cache and monotonic() < deadline:
                self.window.update()
                sleep(0.01)
            image = self.view._cache[path]
            self.assertIsNotNone(image)
            self.assertLessEqual(image.width, THUMBNAIL_SIZE[0])
            self.assertLessEqual(image.height, THUMBNAIL_SIZE[1])
            self.assertEqual(self.view._metadata[path], (path.stat().st_size, (640, 480)))
            self.assertLessEqual(len(self.view._cache), CACHE_LIMIT)
            # A queued result from the previous view cannot restore cleared data.
            generation = self.view.generation
            self.view.clear()
            self.view._loaded.put_nowait((generation, path, image, (0, (1, 1))))
            deadline = monotonic() + 1
            while not self.view._loaded.empty() and monotonic() < deadline:
                self.window.update()
                sleep(0.01)
            self.assertNotIn(path, self.view._cache)
            self.assertNotIn(path, self.view._metadata)

    def test_resizing_reflows_cards_without_losing_keyboard_selection(self) -> None:
        results = [FileSearchResult(Path(f"E:/gallery/{index}.txt"), f"{index}.txt", "") for index in range(20_000)]
        self.view.set_results(results)
        self.view._on_resize(SimpleNamespace(width=350))
        self.assertEqual(self.view.canvas.canvasy(0), 0)
        self.view._on_key(SimpleNamespace(keysym="End"))
        self.view._on_resize(SimpleNamespace(width=520))
        self.assertEqual(self.view._layout_columns, 3)
        self.assertEqual(self.view.get_selected_result(), results[-1])
        self.view._draw_visible()
        self.assertIn(results[-1], self.view.cards)
        self.assertLess(len(self.view.cards), 64)
        self.assertLess(len(self.view.canvas.find_all()), 600)
        self.view._on_key(SimpleNamespace(keysym="Up"))
        self.assertEqual(self.view.get_selected_result(), results[-4])


if __name__ == "__main__":
    unittest.main()
