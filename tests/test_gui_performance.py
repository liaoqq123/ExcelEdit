"""纯 Python 界面调度回归：只调用绑定到假对象的方法，不创建 Tk 或窗口。"""

import unittest
from threading import Event
from tkinter import TclError
from types import MethodType, SimpleNamespace
from unittest.mock import Mock, patch

from gui import ExcelEditApp


class GuiSchedulingTests(unittest.TestCase):
    def test_descendant_configure_events_do_not_schedule_background_work(self) -> None:
        app = SimpleNamespace(
            is_closing=False,
            _schedule_background_image_layout=Mock(),
            _schedule_translucent_strip_layout=Mock(),
        )
        # 模拟进度标签、预览控件和网格线连续改变位置或尺寸。
        for _ in range(100):
            ExcelEditApp._on_app_configure(app, SimpleNamespace(widget=object(), width=200, height=30))
        app._schedule_background_image_layout.assert_not_called()
        app._schedule_translucent_strip_layout.assert_not_called()

    def test_root_resize_notifications_coalesce_to_one_job_per_layer(self) -> None:
        app = SimpleNamespace(
            is_closing=False,
            background_layout_job=None,
            translucent_strip_layout_job=None,
            after=Mock(side_effect=["background-job", "strip-job"]),
            _layout_background_image=Mock(),
            _layout_translucent_strips=Mock(),
        )
        app._schedule_background_image_layout = MethodType(ExcelEditApp._schedule_background_image_layout, app)
        app._schedule_translucent_strip_layout = MethodType(ExcelEditApp._schedule_translucent_strip_layout, app)
        for _ in range(100):
            ExcelEditApp._on_app_configure(app, SimpleNamespace(widget=app, width=1200, height=800))
        self.assertEqual(app.after.call_count, 2)
        self.assertEqual(app.background_layout_job, "background-job")
        self.assertEqual(app.translucent_strip_layout_job, "strip-job")

    def test_empty_background_layout_does_not_reconfigure_widgets(self) -> None:
        app = SimpleNamespace(
            is_closing=False,
            background_layout_job=None,
            background_image_label=Mock(),
            background_source_image=None,
            background_fitted_image=None,
            background_render_size=None,
            background_display_image=None,
            _schedule_translucent_strip_layout=Mock(),
        )
        for _ in range(100):
            ExcelEditApp._layout_background_image(app)
        app.background_image_label.configure.assert_not_called()
        app._schedule_translucent_strip_layout.assert_not_called()

    def test_removing_background_detaches_display_image_only_once(self) -> None:
        app = SimpleNamespace(
            is_closing=False,
            background_layout_job=None,
            background_image_label=Mock(),
            background_source_image=None,
            background_fitted_image=object(),
            background_render_size=(1200, 800),
            background_display_image=object(),
            _schedule_translucent_strip_layout=Mock(),
        )
        ExcelEditApp._layout_background_image(app)
        ExcelEditApp._layout_background_image(app)
        app.background_image_label.configure.assert_called_once_with(image="")
        app._schedule_translucent_strip_layout.assert_called_once()
        self.assertIsNone(app.background_display_image)

    def test_unchanged_background_dimensions_do_not_resize_or_replace_image(self) -> None:
        app = SimpleNamespace(
            is_closing=False,
            background_layout_job=None,
            background_image_label=Mock(),
            background_source_image=object(),
            background_render_size=(1200, 800),
            winfo_width=Mock(return_value=1200),
            winfo_height=Mock(return_value=800),
        )
        with patch("gui.ImageOps.fit") as fit:
            ExcelEditApp._layout_background_image(app)
        fit.assert_not_called()
        app.background_image_label.configure.assert_not_called()

    def test_unchanged_strip_layout_does_not_reconfigure_or_place_overlay(self) -> None:
        image = object()
        frame = SimpleNamespace(
            winfo_x=lambda: 0, winfo_y=lambda: 72,
            winfo_width=lambda: 1200, winfo_height=lambda: 52,
        )
        overlay = Mock()
        app = SimpleNamespace(
            is_closing=False,
            translucent_strip_layout_job=None,
            translucent_strip_frames={"toolbar": frame},
            translucent_strip_overlays={"toolbar": overlay},
            translucent_strip_layout_keys={"toolbar": (0, 72, 1200, 52, id(image))},
            background_fitted_image=image,
        )
        for _ in range(100):
            ExcelEditApp._layout_translucent_strips(app)
        overlay.configure.assert_not_called()
        overlay.place.assert_not_called()
        overlay.lower.assert_not_called()

    def test_empty_strip_layout_detaches_existing_image_once_then_stays_idle(self) -> None:
        frame = SimpleNamespace(
            winfo_x=lambda: 0, winfo_y=lambda: 72,
            winfo_width=lambda: 1200, winfo_height=lambda: 52,
        )
        overlay = Mock()
        overlay.winfo_manager.return_value = "place"
        overlay.place_forget.side_effect = lambda: setattr(overlay.winfo_manager, "return_value", "")
        app = SimpleNamespace(
            is_closing=False,
            translucent_strip_layout_job=None,
            translucent_strip_frames={"toolbar": frame},
            translucent_strip_overlays={"toolbar": overlay},
            translucent_strip_images={"toolbar": object()},
            translucent_strip_layout_keys={"toolbar": object()},
            background_fitted_image=None,
        )
        for _ in range(100):
            ExcelEditApp._layout_translucent_strips(app)
        overlay.configure.assert_called_once_with(image="")
        overlay.place_forget.assert_called_once()
        self.assertEqual(app.translucent_strip_images, {})
        self.assertEqual(app.translucent_strip_layout_keys, {})

    def test_clearing_background_keeps_image_alive_until_widget_detaches_it(self) -> None:
        previous_image = object()
        app = SimpleNamespace(
            is_closing=False,
            background_image_path="old-background.png",
            background_source_image=object(),
            background_fitted_image=object(),
            background_render_size=(1200, 800),
            background_display_image=previous_image,
            background_layout_job=None,
            background_image_label=Mock(),
            translucent_strip_layout_keys={"toolbar": object()},
            _schedule_background_image_layout=Mock(),
            _schedule_translucent_strip_layout=Mock(),
        )
        self.assertTrue(ExcelEditApp._set_background_image_path(app, "", show_error=False))
        self.assertIs(app.background_display_image, previous_image)
        app.background_image_label.configure.assert_not_called()

        def detach_image(**kwargs):
            self.assertEqual(kwargs, {"image": ""})
            self.assertIs(app.background_display_image, previous_image)

        app.background_image_label.configure.side_effect = detach_image
        ExcelEditApp._layout_background_image(app)
        app.background_image_label.configure.assert_called_once_with(image="")
        self.assertIsNone(app.background_display_image)


class GuiCloseSchedulingTests(unittest.TestCase):
    @staticmethod
    def _fake_app():
        preview = SimpleNamespace(clear=Mock(side_effect=TclError("image pyimage3 does not exist")))
        app = SimpleNamespace(
            is_closing=False,
            _cancel_tree_render=Mock(),
            search_cancel_event=Event(),
            reference_cancel_events={"reference": Event()},
            ui_poll_job="poll-job",
            after_cancel=Mock(),
            search_settings_dialog=SimpleNamespace(close=Mock()),
            _save_cache_from_inputs=Mock(),
            preview_pane=preview,
            tk=SimpleNamespace(call=Mock(side_effect=lambda *args: ("layout-job", "preview-job") if args == ("after", "info") else "")),
            destroy=Mock(),
        )
        return app

    def test_close_cancels_work_and_callbacks_then_destroys_without_preview_redraw(self) -> None:
        app = self._fake_app()
        reference_event = app.reference_cancel_events["reference"]
        ExcelEditApp.close_app(app)

        self.assertTrue(app.is_closing)
        self.assertTrue(app.search_cancel_event.is_set())
        self.assertTrue(reference_event.is_set())
        self.assertEqual(app.reference_cancel_events, {})
        self.assertIsNone(app.ui_poll_job)
        app._cancel_tree_render.assert_called_once()
        app.after_cancel.assert_called_once_with("poll-job")
        app.search_settings_dialog.close.assert_called_once()
        app._save_cache_from_inputs.assert_called_once()
        app.tk.call.assert_any_call("after", "cancel", "layout-job")
        app.tk.call.assert_any_call("after", "cancel", "preview-job")
        app.preview_pane.clear.assert_not_called()
        app.destroy.assert_called_once()

    def test_cache_save_error_still_cancels_callbacks_and_destroys(self) -> None:
        app = self._fake_app()
        app._save_cache_from_inputs.side_effect = RuntimeError("cache write failed")
        with self.assertRaisesRegex(RuntimeError, "cache write failed"):
            ExcelEditApp.close_app(app)
        app.tk.call.assert_any_call("after", "cancel", "layout-job")
        app.tk.call.assert_any_call("after", "cancel", "preview-job")
        app.preview_pane.clear.assert_not_called()
        app.destroy.assert_called_once()

    def test_callback_cleanup_error_still_destroys(self) -> None:
        app = self._fake_app()
        app.tk.call.side_effect = TclError("interpreter is shutting down")
        with self.assertRaises(TclError):
            ExcelEditApp.close_app(app)
        app.destroy.assert_called_once()

    def test_close_is_idempotent(self) -> None:
        app = self._fake_app()
        ExcelEditApp.close_app(app)
        ExcelEditApp.close_app(app)
        app._cancel_tree_render.assert_called_once()
        app._save_cache_from_inputs.assert_called_once()
        app.destroy.assert_called_once()


if __name__ == "__main__":
    unittest.main()
