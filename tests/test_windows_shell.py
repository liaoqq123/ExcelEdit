"""仅使用假函数表验证 Shell 接口；禁止加载真实 DLL 或启动任何窗口。"""

import ctypes
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from windows_shell import SHELLEXECUTEINFOW, launch_path, reveal_item


class WindowsShellTests(unittest.TestCase):
    def setUp(self) -> None:
        self.native = patch("windows_shell._native_api", side_effect=AssertionError("测试禁止访问真实 Shell API"))
        self.native.start()
        self.addCleanup(self.native.stop)
        self.api = SimpleNamespace(**{name: Mock() for name in (
            "CoInitializeEx", "CoUninitialize", "CoTaskMemFree", "SHParseDisplayName",
            "SHOpenFolderAndSelectItems", "ShellExecuteExW", "GetProcessId", "CloseHandle", "GetLastError",
        )})
        self.api.CoInitializeEx.return_value = 0
        self.api.SHParseDisplayName.side_effect = self._parse_item
        self.api.SHOpenFolderAndSelectItems.return_value = 0
        self.api.ShellExecuteExW.side_effect = self._launch
        self.api.GetProcessId.return_value = 789
        self.api.GetLastError.return_value = 5

    @staticmethod
    def _parse_item(path, context, output, attributes, output_attributes):
        ctypes.cast(output, ctypes.POINTER(ctypes.c_void_p))[0] = 0x12345678
        return 0

    @staticmethod
    def _launch(output):
        info = ctypes.cast(output, ctypes.POINTER(SHELLEXECUTEINFOW)).contents
        info.hProcess = 0x76543210
        return 1

    def test_reveal_uses_unicode_full_item_pidl_for_files_and_subfolders(self) -> None:
        for name in (r"C:\当前目录\子文件夹", r"C:\含 空格,逗号\中文.xlsx", r"\\server\share\目录\文件.xlsx"):
            with self.subTest(path=name):
                self.api.SHParseDisplayName.reset_mock()
                self.api.SHOpenFolderAndSelectItems.reset_mock()
                reveal_item(Path(name), api=self.api)
                self.assertEqual(self.api.SHParseDisplayName.call_args.args[0], str(Path(name)))
                args = self.api.SHOpenFolderAndSelectItems.call_args.args
                self.assertEqual(args[0].value, 0x12345678)
                self.assertEqual(args[1:], (0, None, 0))
        self.assertEqual(self.api.CoUninitialize.call_count, 3)
        self.assertEqual(self.api.CoTaskMemFree.call_count, 3)

    def test_shell_hresult_failure_frees_pidl_and_uninitializes_com(self) -> None:
        # 检查未被转换成 Python 负数的高位 HRESULT 同样被识别为失败。
        self.api.SHOpenFolderAndSelectItems.return_value = 0x80070005
        with self.assertRaisesRegex(OSError, "80070005"):
            reveal_item(Path(r"C:\data\folder"), api=self.api)
        self.assertEqual(self.api.CoTaskMemFree.call_args.args[0].value, 0x12345678)
        self.api.CoUninitialize.assert_called_once()

    def test_parse_failure_frees_any_partial_pidl_and_skips_selection(self) -> None:
        def fail_parse(*args):
            self._parse_item(*args)
            return -2147024894

        self.api.SHParseDisplayName.side_effect = fail_parse
        with self.assertRaises(OSError):
            reveal_item(Path(r"C:\data\missing.xlsx"), api=self.api)
        self.api.SHOpenFolderAndSelectItems.assert_not_called()
        self.api.CoTaskMemFree.assert_called_once()
        self.api.CoUninitialize.assert_called_once()

    def test_com_s_false_is_balanced_but_failed_initialization_is_not(self) -> None:
        self.api.CoInitializeEx.return_value = 1
        reveal_item(Path(r"C:\data\folder"), api=self.api)
        self.api.CoUninitialize.assert_called_once()
        self.api.CoUninitialize.reset_mock()
        self.api.SHParseDisplayName.reset_mock()
        self.api.CoInitializeEx.return_value = 0x80010106
        with self.assertRaisesRegex(OSError, "80010106"):
            reveal_item(Path(r"C:\data\folder"), api=self.api)
        self.api.CoUninitialize.assert_not_called()
        self.api.SHParseDisplayName.assert_not_called()

    def test_launch_preserves_literal_unicode_path_and_closes_process_handle(self) -> None:
        path = Path(r"C:\含 空格,逗号\中文文件夹")
        observed = {}

        def capture_launch(output):
            info = ctypes.cast(output, ctypes.POINTER(SHELLEXECUTEINFOW)).contents
            observed.update(size=info.cbSize, flags=info.fMask, verb=info.lpVerb,
                            path=info.lpFile, show=info.nShow, parameters=info.lpParameters)
            return self._launch(output)

        self.api.ShellExecuteExW.side_effect = capture_launch
        self.assertEqual(launch_path(path, api=self.api), 789)
        self.assertEqual(observed, {"size": ctypes.sizeof(SHELLEXECUTEINFOW), "flags": 0x540,
                                   "verb": "open", "path": str(path), "show": 1, "parameters": None})
        self.api.GetProcessId.assert_called_once_with(0x76543210)
        self.api.CloseHandle.assert_called_once_with(0x76543210)
        self.api.CoUninitialize.assert_called_once()

    def test_launch_without_process_handle_is_successful_without_pid(self) -> None:
        self.api.ShellExecuteExW.side_effect = None
        self.api.ShellExecuteExW.return_value = 1
        self.assertIsNone(launch_path(Path(r"C:\data\folder"), api=self.api))
        self.api.GetProcessId.assert_not_called()
        self.api.CloseHandle.assert_not_called()
        self.api.CoUninitialize.assert_called_once()

    def test_launch_failure_keeps_error_and_releases_allocated_handle(self) -> None:
        def fail_launch(output):
            self._launch(output)
            return 0

        self.api.ShellExecuteExW.side_effect = fail_launch
        with self.assertRaises(OSError) as context:
            launch_path(Path(r"C:\data\file.xlsx"), api=self.api)
        self.assertEqual(context.exception.errno, 5)
        self.api.GetProcessId.assert_not_called()
        self.api.CloseHandle.assert_called_once_with(0x76543210)
        self.api.CoUninitialize.assert_called_once()

    def test_pid_query_exception_still_releases_handle(self) -> None:
        self.api.GetProcessId.side_effect = RuntimeError("PID query failed")
        with self.assertRaisesRegex(RuntimeError, "PID query failed"):
            launch_path(Path(r"C:\data\file.xlsx"), api=self.api)
        self.api.CloseHandle.assert_called_once_with(0x76543210)
        self.api.CoUninitialize.assert_called_once()

    def test_shell_structure_has_windows_pointer_safe_layout(self) -> None:
        self.assertEqual(ctypes.sizeof(SHELLEXECUTEINFOW), 112 if ctypes.sizeof(ctypes.c_void_p) == 8 else 60)
        self.assertEqual(SHELLEXECUTEINFOW.cbSize.offset, 0)
        self.assertEqual(SHELLEXECUTEINFOW.fMask.offset, 4)
        self.assertEqual(SHELLEXECUTEINFOW.hwnd.offset, 8)
        self.assertEqual(SHELLEXECUTEINFOW.hProcess.offset + ctypes.sizeof(ctypes.c_void_p), ctypes.sizeof(SHELLEXECUTEINFOW))


if __name__ == "__main__":
    unittest.main()
