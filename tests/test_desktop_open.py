"""打开操作的纯模拟测试：不调用 Shell、PowerShell、Excel、焦点 API 或 Tk。"""

import base64
import json
from pathlib import Path
import subprocess
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import desktop_open
from desktop_open import OpenResult
from gui import ExcelEditApp, ResultTarget


RUN_POWERSHELL = desktop_open._run_powershell


class DesktopOpenTests(unittest.TestCase):
    def setUp(self) -> None:
        forbidden_tk = patch("tkinter._tkinter.create", side_effect=AssertionError("测试禁止创建 Tk"))
        forbidden_tk.start()
        self.addCleanup(forbidden_tk.stop)
        forbidden_activation = patch("window_activation.bring_window_to_front", side_effect=AssertionError("打开完成后禁止再次激活窗口"))
        forbidden_activation.start()
        self.addCleanup(forbidden_activation.stop)
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.folder = self.root / "当前目录" / "子 文件夹,中文"
        self.folder.mkdir(parents=True)
        self.excel = self.folder / "中文, 空格's.XLSX"
        self.excel.write_bytes(b"fixture")
        self.document = self.folder / "普通, 文件.txt"
        self.document.write_text("fixture", encoding="utf-8")
        self.calls = {}
        for name in ("reveal_item", "launch_path", "_run_powershell", "is_window_foreground", "find_process_window", "sleep"):
            patched = patch.object(desktop_open, name, side_effect=AssertionError(f"测试未模拟 {name}"))
            self.calls[name] = patched.start()
            self.addCleanup(patched.stop)
        for name in ("run", "Popen"):
            patched = patch.object(desktop_open.subprocess, name, side_effect=AssertionError(f"测试禁止真实 subprocess.{name}"))
            self.calls[name] = patched.start()
            self.addCleanup(patched.stop)
        platform = patch.object(desktop_open.sys, "platform", "win32")
        platform.start()
        self.addCleanup(platform.stop)
        for module in ("windows_shell", "window_activation"):
            # 注入函数之外的意外原生入口也立即失败。
            import importlib
            imported = importlib.import_module(module)
            if hasattr(imported, "_native_api"):
                native = patch.object(imported, "_native_api", side_effect=AssertionError("测试禁止真实原生 API"))
                native.start()
                self.addCleanup(native.stop)

    def allow(self, name, result=None):
        function = self.calls[name]
        function.side_effect = None
        function.return_value = result
        return function

    def explorer_response(self, *, located=True, foreground=True, hwnd=4321):
        self.allow("_run_powershell", {"hwnd": hwnd, "located": located, "error": ""})
        self.allow("is_window_foreground", foreground)

    def test_subfolder_reveal_selects_full_target_inside_real_parent(self) -> None:
        self.allow("reveal_item")
        self.explorer_response()
        result = desktop_open.open_folder(self.folder, reveal=True)
        self.calls["reveal_item"].assert_called_once_with(self.folder)
        self.calls["launch_path"].assert_not_called()
        payload = self.calls["_run_powershell"].call_args.args[1]
        self.assertEqual(payload["folder"], str(self.folder.parent))
        self.assertEqual(payload["target"], str(self.folder))
        self.assertTrue(result.opened and result.located and result.foreground)
        self.calls["is_window_foreground"].assert_called_once_with(4321)

    def test_signed_com_handle_is_normalized_before_foreground_activation(self) -> None:
        self.allow("reveal_item")
        self.explorer_response(hwnd=-2147483648)
        result = desktop_open.open_folder(self.folder, reveal=True)
        self.assertEqual(result.hwnd, 2147483648)
        self.calls["is_window_foreground"].assert_called_once_with(2147483648)

    def test_excel_signed_handle_is_normalized_and_partial_location_keeps_window(self) -> None:
        self.allow("launch_path")
        self.allow("_run_powershell", {"opened": True, "located": False, "hwnd": -2147483648, "error": "工作表不存在"})
        self.allow("is_window_foreground", True)
        result = desktop_open.open_excel(self.excel, "已删除的表")
        self.assertEqual(result.hwnd, 2147483648)
        self.assertTrue(result.opened and result.foreground)
        self.assertFalse(result.located)
        self.calls["launch_path"].assert_called_once_with(self.excel)

    def test_com_handle_validation_preserves_positive_pointer_values(self) -> None:
        self.assertEqual(desktop_open._com_hwnd(2147483648), 2147483648)
        for invalid in (None, True, 0, "not-a-handle", -(1 << 32), 1 << 64):
            with self.subTest(invalid=invalid):
                self.assertEqual(desktop_open._com_hwnd(invalid), 0)

    def test_normal_directory_open_enters_target_without_revealing_it(self) -> None:
        self.allow("launch_path")
        self.explorer_response()
        result = desktop_open.open_target(self.folder)
        self.calls["launch_path"].assert_called_once_with(self.folder)
        self.calls["reveal_item"].assert_not_called()
        payload = self.calls["_run_powershell"].call_args.args[1]
        self.assertEqual((payload["folder"], payload["target"]), (str(self.folder), ""))
        self.assertTrue(result.located)

    def test_file_reveal_preserves_spaces_commas_unicode_and_apostrophe(self) -> None:
        self.allow("reveal_item")
        self.explorer_response()
        result = desktop_open.open_target(self.excel, reveal=True)
        self.calls["reveal_item"].assert_called_once_with(self.excel)
        script, payload = self.calls["_run_powershell"].call_args.args[:2]
        self.assertEqual(payload["folder"], str(self.folder))
        self.assertEqual(payload["target"], str(self.excel))
        self.assertNotIn(str(self.excel), script)
        self.assertTrue(result.located)

    def test_drive_root_is_entered_without_selecting_an_impossible_parent(self) -> None:
        drive = Path(self.root.anchor)
        self.allow("launch_path")
        self.explorer_response()
        result = desktop_open.open_folder(drive, reveal=True)
        self.calls["launch_path"].assert_called_once_with(drive)
        self.calls["reveal_item"].assert_not_called()
        payload = self.calls["_run_powershell"].call_args.args[1]
        self.assertEqual((payload["folder"], payload["target"]), (str(drive), ""))
        self.assertTrue(result.opened and result.located)

    def test_missing_item_opens_existing_parent_without_claiming_location(self) -> None:
        stale = self.folder / "removed.xlsx"
        self.allow("launch_path")
        self.explorer_response()
        result = desktop_open.open_folder(stale, reveal=True)
        self.calls["launch_path"].assert_called_once_with(self.folder)
        self.calls["reveal_item"].assert_not_called()
        self.assertTrue(result.opened)
        self.assertFalse(result.located)
        self.assertIn("移动或删除", result.message)

    def test_missing_parent_fails_before_any_open_request(self) -> None:
        with self.assertRaises(FileNotFoundError):
            desktop_open.open_folder(self.root / "removed-parent" / "missing.xlsx", reveal=True)
        self.calls["launch_path"].assert_not_called()
        self.calls["reveal_item"].assert_not_called()
        self.calls["_run_powershell"].assert_not_called()

    def test_native_selection_failure_falls_back_to_opening_correct_parent(self) -> None:
        self.calls["reveal_item"].side_effect = OSError("PIDL selection failed")
        self.allow("launch_path")
        self.explorer_response()
        result = desktop_open.open_folder(self.folder, reveal=True)
        self.calls["launch_path"].assert_called_once_with(self.folder.parent)
        self.assertEqual(self.calls["_run_powershell"].call_args.args[1]["target"], str(self.folder))
        self.assertTrue(result.opened and result.located)

    def test_unknown_explorer_window_does_not_focus_an_unrelated_window(self) -> None:
        self.allow("reveal_item")
        self.explorer_response(located=False, hwnd=0)
        result = desktop_open.open_folder(self.excel, reveal=True)
        self.assertTrue(result.opened)
        self.assertFalse(result.located)
        self.assertFalse(result.foreground)
        self.calls["is_window_foreground"].assert_not_called()
        self.calls["find_process_window"].assert_not_called()

    def test_excel_without_sheet_opens_normally_without_com_or_powershell(self) -> None:
        self.allow("launch_path")
        result = desktop_open.open_target(self.excel)
        self.calls["launch_path"].assert_called_once_with(self.excel)
        self.calls["_run_powershell"].assert_not_called()
        self.calls["find_process_window"].assert_not_called()
        self.calls["is_window_foreground"].assert_not_called()
        self.assertEqual(result, OpenResult(True, True, False, 0, ""))

    def test_excel_navigation_starts_with_exactly_one_normal_open_request(self) -> None:
        self.allow("launch_path")
        self.allow("_run_powershell", {"opened": True, "located": True, "hwnd": 9876, "error": ""})
        self.allow("is_window_foreground", True)
        events = []
        self.calls["launch_path"].side_effect = lambda path: events.append("open")
        self.calls["_run_powershell"].side_effect = lambda *args: events.append("navigate") or {"located": True, "hwnd": 9876}
        result = desktop_open.open_target(self.excel, "技能", 7, 3)
        script, payload = self.calls["_run_powershell"].call_args.args[:2]
        self.assertEqual(payload, {"path": str(self.excel), "sheet": "技能", "row": 7, "column": 3, "wait": desktop_open.EXCEL_LOOKUP_WAIT_SECONDS})
        self.assertNotIn(str(self.excel), script)
        self.assertEqual(events, ["open", "navigate"])
        self.calls["launch_path"].assert_called_once_with(self.excel)
        self.calls["find_process_window"].assert_not_called()
        self.calls["is_window_foreground"].assert_called_once_with(9876)
        self.assertEqual(result, OpenResult(True, True, True, 9876, ""))

    def test_excel_selection_failure_retains_open_workbook_without_duplicate_launch(self) -> None:
        self.allow("launch_path")
        self.allow("_run_powershell", {"opened": True, "located": False, "hwnd": 9876, "error": "工作表已删除"})
        self.allow("is_window_foreground", True)
        result = desktop_open.open_excel(self.excel, "不存在的表", 7, 3)
        self.assertTrue(result.opened)
        self.assertFalse(result.located)
        self.assertTrue(result.foreground)
        self.assertIn("工作表已删除", result.message)
        self.calls["launch_path"].assert_called_once_with(self.excel)
        self.assertEqual(self.calls["_run_powershell"].call_args.args[1]["sheet"], "不存在的表")

    def test_denied_foreground_keeps_opened_and_located_state(self) -> None:
        self.allow("launch_path")
        self.allow("_run_powershell", {"opened": True, "located": True, "hwnd": 9876, "error": ""})
        self.allow("is_window_foreground", False)
        result = desktop_open.open_excel(self.excel, "技能", 8, 4)
        self.assertEqual(result, OpenResult(True, True, False, 9876, ""))

    def test_excel_timeout_releases_serialization_lock_without_association_retry(self) -> None:
        self.allow("launch_path")
        lock = SimpleNamespace(acquire=Mock(return_value=True), release=Mock())
        self.calls["_run_powershell"].side_effect = subprocess.TimeoutExpired("powershell", 20)
        with patch.object(desktop_open, "_EXCEL_LOCK", lock):
            result = desktop_open.open_excel(self.excel, "技能", 8, 4)
        lock.acquire.assert_called_once_with(timeout=desktop_open.OPEN_TIMEOUT_SECONDS)
        lock.release.assert_called_once()
        self.assertTrue(result.opened)
        self.assertFalse(result.located)
        self.assertIn("超时", result.message)
        self.calls["launch_path"].assert_called_once_with(self.excel)

    def test_busy_excel_navigation_does_not_prevent_system_open(self) -> None:
        self.allow("launch_path")
        lock = SimpleNamespace(acquire=Mock(return_value=False), release=Mock())
        with patch.object(desktop_open, "_EXCEL_LOCK", lock):
            result = desktop_open.open_excel(self.excel, "技能")
        self.assertTrue(result.opened)
        self.assertFalse(result.located)
        self.assertIn("其他 Excel 定位", result.message)
        lock.release.assert_not_called()
        self.calls["_run_powershell"].assert_not_called()
        self.calls["launch_path"].assert_called_once_with(self.excel)

    def test_plain_excel_open_is_not_blocked_by_another_navigation(self) -> None:
        self.allow("launch_path")
        lock = SimpleNamespace(acquire=Mock(return_value=False), release=Mock())
        with patch.object(desktop_open, "_EXCEL_LOCK", lock):
            result = desktop_open.open_excel(self.excel)
        self.assertTrue(result.opened)
        lock.acquire.assert_not_called()
        self.calls["launch_path"].assert_called_once_with(self.excel)
        self.calls["_run_powershell"].assert_not_called()

    def test_missing_or_broken_navigation_helper_does_not_prevent_normal_open(self) -> None:
        self.allow("launch_path")
        for error in (FileNotFoundError("PowerShell missing"), PermissionError("blocked"), OSError("bad response")):
            with self.subTest(error=error):
                self.calls["launch_path"].reset_mock()
                self.calls["_run_powershell"].side_effect = error
                result = desktop_open.open_excel(self.excel, "技能", 8, 4)
                self.assertTrue(result.opened)
                self.assertFalse(result.located)
                self.assertIn("自动定位不可用", result.message)
                self.calls["launch_path"].assert_called_once_with(self.excel)

    def test_unconnected_or_invisible_com_workbook_does_not_trigger_duplicate_open(self) -> None:
        self.allow("launch_path")
        self.allow("_run_powershell", {"opened": False, "located": False, "hwnd": 0, "error": "没有可见目标工作簿"})
        result = desktop_open.open_excel(self.excel, "技能", 8, 4)
        self.assertTrue(result.opened)
        self.assertFalse(result.located)
        self.assertEqual(result.hwnd, 0)
        self.assertIn("没有可见目标工作簿", result.message)
        self.calls["launch_path"].assert_called_once_with(self.excel)
        self.calls["is_window_foreground"].assert_not_called()

    def test_navigation_cannot_report_success_without_a_confirmed_window(self) -> None:
        self.allow("launch_path")
        self.allow("_run_powershell", {"located": True, "hwnd": 0})
        result = desktop_open.open_excel(self.excel, "技能", 8, 4)
        self.assertTrue(result.opened)
        self.assertFalse(result.located)
        self.assertTrue(result.message)

    def test_minimized_excel_after_navigation_is_observed_without_restoring_it(self) -> None:
        self.allow("launch_path")
        self.allow("_run_powershell", {"located": False, "hwnd": 9876, "error": "目标窗口已最小化，已跳过自动定位"})
        self.allow("is_window_foreground", False)
        result = desktop_open.open_excel(self.excel, "技能", 8, 4)
        self.assertTrue(result.opened)
        self.assertFalse(result.located)
        self.assertFalse(result.foreground)
        self.assertIn("已最小化", result.message)
        self.calls["launch_path"].assert_called_once_with(self.excel)
        self.calls["is_window_foreground"].assert_called_once_with(9876)

    def test_failed_system_open_does_not_attempt_navigation_or_acquire_lock(self) -> None:
        lock = SimpleNamespace(acquire=Mock(return_value=True), release=Mock())
        self.calls["launch_path"].side_effect = OSError("关联打开失败")
        with patch.object(desktop_open, "_EXCEL_LOCK", lock):
            with self.assertRaisesRegex(OSError, "关联打开失败"):
                desktop_open.open_excel(self.excel, "技能")
        lock.acquire.assert_not_called()
        lock.release.assert_not_called()
        self.calls["_run_powershell"].assert_not_called()

    def test_invalid_cell_coordinates_do_not_start_automation(self) -> None:
        for row, column in ((0, 1), (1, 16385), (1048577, 1), (None, 1)):
            with self.subTest(row=row, column=column), self.assertRaises(ValueError):
                desktop_open.open_excel(self.excel, "技能", row, column)
        self.calls["_run_powershell"].assert_not_called()

    def test_non_excel_open_only_observes_window_from_returned_process(self) -> None:
        self.allow("launch_path", 2468)
        self.allow("find_process_window", 1357)
        self.allow("is_window_foreground", True)
        result = desktop_open.open_target(self.document)
        self.calls["launch_path"].assert_called_once_with(self.document)
        self.calls["find_process_window"].assert_called_once_with(2468)
        self.calls["is_window_foreground"].assert_called_once_with(1357)
        self.calls["_run_powershell"].assert_not_called()
        self.assertEqual(result.hwnd, 1357)

    def test_explorer_payload_preserves_unc_path_without_network_access(self) -> None:
        folder = Path(r"\\server\share\目录, 空格")
        target = folder / "子目录"
        self.allow("_run_powershell", {"hwnd": 1, "located": True})
        desktop_open._explorer_window(folder, target)
        self.assertEqual(self.calls["_run_powershell"].call_args.args[1], {
            "folder": str(folder), "target": str(target), "wait": desktop_open.EXPLORER_WAIT_SECONDS,
        })

    def test_powershell_uses_encoded_fixed_script_json_stdin_and_hidden_process(self) -> None:
        payload = {"path": str(self.excel), "sheet": "角色's表; $文本", "row": 7}
        self.allow("run", SimpleNamespace(returncode=0, stdout='{"opened":true,"hwnd":9876}', stderr=""))
        result = RUN_POWERSHELL("Write-Output $request.path", payload, timeout=4)
        args, kwargs = self.calls["run"].call_args.args[0], self.calls["run"].call_args.kwargs
        self.assertIn("-EncodedCommand", args)
        decoded = base64.b64decode(args[args.index("-EncodedCommand") + 1]).decode("utf-16-le")
        self.assertIn("ConvertFrom-Json", decoded)
        self.assertNotIn(str(self.excel), decoded)
        self.assertNotIn(payload["sheet"], decoded)
        self.assertEqual(json.loads(kwargs["input"]), payload)
        self.assertEqual(kwargs["creationflags"], getattr(subprocess, "CREATE_NO_WINDOW", 0))
        self.assertEqual(kwargs["encoding"], "utf-8")
        self.assertEqual(kwargs["timeout"], 4)
        self.assertEqual(result, {"opened": True, "hwnd": 9876})

    def test_powershell_bad_output_is_an_explicit_failure(self) -> None:
        for output in ("not JSON", "[]", "null"):
            with self.subTest(output=output):
                self.allow("run", SimpleNamespace(returncode=0, stdout=output, stderr=""))
                with self.assertRaises(OSError):
                    RUN_POWERSHELL("fixed script", {})

    def test_gui_completion_does_not_claim_denied_foreground(self) -> None:
        target = ResultTarget(self.excel, "技能", 7, 3)
        app = SimpleNamespace(opening_targets={target}, status_label=SimpleNamespace(configure=Mock()), _sync_result_actions=Mock())
        ExcelEditApp._open_target_finished(app, target, OpenResult(True, True, False, 9876))
        status = app.status_label.configure.call_args.kwargs["text"]
        self.assertIn("窗口当前未在前台", status)
        self.assertNotIn("并切到前台", status)
        self.assertNotIn(target, app.opening_targets)

    def test_gui_reports_location_and_foreground_failures_independently(self) -> None:
        target = ResultTarget(self.excel, "已删除的表")
        app = SimpleNamespace(opening_targets={target}, status_label=SimpleNamespace(configure=Mock()), _sync_result_actions=Mock())
        ExcelEditApp._open_target_finished(app, target, OpenResult(True, False, False, 9876))
        status = app.status_label.configure.call_args.kwargs["text"]
        self.assertIn("窗口当前未在前台", status)
        self.assertIn("未能完成目标定位", status)

    def test_gui_worker_posts_precise_reveal_result_without_touching_tk(self) -> None:
        target = ResultTarget(self.folder, open_containing_folder=True)
        result = OpenResult(True, True, True, 4321)
        queued = []
        app = SimpleNamespace(_post_to_ui=queued.append, _open_target_finished=Mock(), _open_target_failed=Mock())
        with patch("gui.open_desktop_target", return_value=result) as opener:
            ExcelEditApp._open_target_in_background(app, target)
        opener.assert_called_once_with(self.folder, None, None, None, reveal=True)
        app._open_target_finished.assert_not_called()
        self.assertEqual(len(queued), 1)
        queued[0]()
        app._open_target_finished.assert_called_once_with(target, result)
        app._open_target_failed.assert_not_called()


if __name__ == "__main__":
    unittest.main()
