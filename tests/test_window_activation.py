"""仅 Fake API 测试；不加载 user32，不打开窗口或改变真实焦点。"""

import ctypes
import unittest
from unittest.mock import patch

import window_activation as activation


TARGET = 0x123456789 if ctypes.sizeof(ctypes.c_void_p) == 8 else 0x1234567
OTHER = 88


class FakeAPI:
    def __init__(self, *, activate_after=1, minimized=False, topmost=False):
        self.calls = []
        self.valid = True
        self.identity = (11, 23)
        self.minimized = minimized
        self.topmost = topmost
        self.foreground = OTHER
        self.activate_after = activate_after
        self.foreground_attempts = 0
        self.restore_succeeds = True
        self.restore_after_waits = 0
        self.restore_pending = False
        self.wait_count = 0
        self.raise_on_activation = False
        self.recycle_after_attempt = None

    def is_window(self, hwnd):
        self.calls.append(("valid", hwnd))
        return self.valid

    def window_identity(self, hwnd):
        self.calls.append(("identity", hwnd))
        if self.recycle_after_attempt and self.foreground_attempts >= self.recycle_after_attempt:
            return (99, 101)
        return self.identity

    def is_minimized(self, hwnd):
        self.calls.append(("minimized", hwnd))
        return self.minimized

    def restore(self, hwnd):
        self.calls.append(("restore", hwnd))
        if self.restore_succeeds:
            if self.restore_after_waits:
                self.restore_pending = True
            else:
                self.minimized = False
        return self.restore_succeeds

    def foreground_window(self):
        self.calls.append(("foreground",))
        return self.foreground

    def set_foreground(self, hwnd):
        self.calls.append(("activate", hwnd))
        self.foreground_attempts += 1
        if self.raise_on_activation:
            raise OSError("activation denied")
        if self.activate_after is not None and self.foreground_attempts >= self.activate_after:
            self.foreground = hwnd
        # Deliberately return success regardless: the caller must verify focus.
        return True

    def is_topmost(self, hwnd):
        self.calls.append(("pinned", hwnd))
        raise AssertionError("Normal activation must not access pinning APIs")

    def set_topmost(self, hwnd, enabled):
        self.calls.append(("topmost", hwnd, enabled))
        raise AssertionError("Normal activation must not change pinning")

    def wait(self, seconds):
        self.calls.append(("wait", seconds))
        self.wait_count += 1
        if self.restore_pending and self.wait_count >= self.restore_after_waits:
            self.minimized = False


class WindowActivationTests(unittest.TestCase):
    def test_invalid_handles_never_access_api(self):
        api = FakeAPI()
        for hwnd in (None, False, True, "42", 0, -1, 1 << (ctypes.sizeof(ctypes.c_void_p) * 8)):
            with self.subTest(hwnd=hwnd):
                self.assertFalse(activation.bring_window_to_front(hwnd, api=api))
        self.assertEqual(api.calls, [])

    def test_restores_minimized_target_before_activation(self):
        api = FakeAPI(minimized=True)
        self.assertTrue(activation.bring_window_to_front(TARGET, api=api))
        self.assertFalse(api.minimized)
        self.assertLess(api.calls.index(("restore", TARGET)), api.calls.index(("activate", TARGET)))
        self.assertEqual(api.calls.count(("restore", TARGET)), 1)
        self.assertEqual(api.foreground_attempts, 1)
        self.assertFalse(any(call[0] in {"pinned", "topmost"} for call in api.calls))

    def test_already_foreground_target_needs_no_focus_or_z_order_change(self):
        api = FakeAPI()
        api.foreground = TARGET
        self.assertTrue(activation.bring_window_to_front(TARGET, api=api))
        self.assertFalse(any(call[0] in {"restore", "activate", "pinned", "topmost", "wait"} for call in api.calls))

    def test_success_flag_without_actual_foreground_is_not_reported_as_success(self):
        api = FakeAPI(activate_after=None)
        self.assertFalse(activation.bring_window_to_front(TARGET, api=api))
        self.assertEqual(api.foreground_attempts, 1)
        self.assertFalse(api.topmost)
        self.assertFalse(any(call[0] in {"pinned", "topmost", "wait"} for call in api.calls))

    def test_normal_window_activates_without_resizing_restoring_or_waiting(self):
        api = FakeAPI()
        self.assertTrue(activation.bring_window_to_front(TARGET, api=api))
        self.assertEqual(api.foreground_attempts, 1)
        self.assertFalse(any(call[0] in {"restore", "pinned", "topmost", "wait"} for call in api.calls))

    def test_denied_activation_does_not_retry_or_force_window_forward(self):
        api = FakeAPI(activate_after=2)
        self.assertFalse(activation.bring_window_to_front(TARGET, api=api))
        self.assertEqual(api.foreground_attempts, 1)
        self.assertEqual(api.foreground, OTHER)
        self.assertFalse(any(call[0] in {"pinned", "topmost", "wait"} for call in api.calls))

    def test_preexisting_topmost_state_is_preserved_without_accessing_pinning_api(self):
        for topmost in (False, True):
            for activate_after in (None, 1):
                with self.subTest(topmost=topmost, activate_after=activate_after):
                    api = FakeAPI(activate_after=activate_after, topmost=topmost)
                    self.assertEqual(activation.bring_window_to_front(TARGET, api=api), activate_after == 1)
                    self.assertEqual(api.topmost, topmost)
                    self.assertFalse(any(call[0] in {"pinned", "topmost"} for call in api.calls))

    def test_activation_exception_returns_false_without_retry(self):
        api = FakeAPI()
        api.raise_on_activation = True
        self.assertFalse(activation.bring_window_to_front(TARGET, api=api))
        self.assertEqual(api.foreground_attempts, 1)
        self.assertFalse(any(call[0] in {"pinned", "topmost", "wait"} for call in api.calls))

    def test_async_minimized_restore_is_polled_without_repeating_requests(self):
        api = FakeAPI(minimized=True)
        api.restore_after_waits = 2
        self.assertTrue(activation.bring_window_to_front(TARGET, api=api))
        self.assertEqual(api.calls.count(("restore", TARGET)), 1)
        self.assertEqual(api.wait_count, 2)
        self.assertEqual(api.foreground_attempts, 1)

    def test_failed_restore_does_not_activate_or_repeat_restore(self):
        api = FakeAPI(minimized=True)
        api.restore_succeeds = False
        self.assertFalse(activation.bring_window_to_front(TARGET, api=api))
        self.assertEqual(api.calls.count(("restore", TARGET)), 1)
        self.assertEqual(api.foreground_attempts, 0)
        self.assertEqual(api.wait_count, 0)

    def test_unresponsive_minimized_window_has_bounded_read_only_wait(self):
        api = FakeAPI(minimized=True)
        api.restore_after_waits = activation.RESTORE_CHECKS + 1
        self.assertFalse(activation.bring_window_to_front(TARGET, api=api))
        self.assertEqual(api.calls.count(("restore", TARGET)), 1)
        self.assertEqual(api.foreground_attempts, 0)
        self.assertEqual(api.wait_count, activation.RESTORE_CHECKS)

    def test_recycled_target_handle_aborts_without_activating_new_window(self):
        api = FakeAPI()
        api.recycle_after_attempt = 1
        self.assertFalse(activation.bring_window_to_front(TARGET, api=api))
        self.assertEqual(api.foreground_attempts, 1)
        self.assertFalse(any(call[0] in {"pinned", "topmost"} for call in api.calls))
        self.assertTrue(all(call[1] == TARGET for call in api.calls if call[0] in {"activate", "restore"}))

    def test_recycled_handle_during_restore_is_not_activated(self):
        api = FakeAPI(minimized=True)
        def recycle_on_wait(_seconds):
            api.identity = (99, 101)
        api.wait = recycle_on_wait
        self.assertFalse(activation.bring_window_to_front(TARGET, api=api))
        self.assertEqual(api.foreground_attempts, 0)

    def test_invalid_or_unidentified_window_is_ignored(self):
        for invalid_identity in (False, True):
            api = FakeAPI()
            if invalid_identity:
                api.identity = (0, 0)
            else:
                api.valid = False
            self.assertFalse(activation.bring_window_to_front(TARGET, api=api))
            self.assertEqual(api.foreground_attempts, 0)

    def test_unsupported_platform_returns_false_without_loading_user32(self):
        with patch("window_activation.os.name", "posix"), patch("window_activation.ctypes.WinDLL", create=True) as loader:
            self.assertFalse(activation.bring_window_to_front(TARGET))
        loader.assert_not_called()

    def test_ctypes_signature_binding_is_pointer_safe_and_calls_no_native_function(self):
        functions = {}

        class FakeFunction:
            def __call__(self, *args):
                raise AssertionError("Native function execution is prohibited in this test")

        class FakeDLL:
            def __getattr__(self, name):
                return functions.setdefault(name, FakeFunction())

        with patch("window_activation.os.name", "nt"), patch("window_activation.ctypes.WinDLL", return_value=FakeDLL(), create=True) as loader:
            activation._WindowsAPI()
        loader.assert_called_once_with("user32", use_last_error=True)
        for name in ("IsWindow", "IsWindowVisible", "IsIconic", "ShowWindowAsync", "SetForegroundWindow", "GetWindowThreadProcessId", "GetWindow"):
            self.assertIs(functions[name].argtypes[0], ctypes.c_void_p)
        self.assertIs(functions["GetForegroundWindow"].restype, ctypes.c_void_p)
        self.assertIs(functions["GetWindow"].restype, ctypes.c_void_p)
        self.assertIs(functions["EnumWindows"].argtypes[0]._argtypes_[0], ctypes.c_void_p)
        self.assertIs(functions["EnumWindows"].argtypes[1], ctypes.c_ssize_t)
        self.assertNotIn("SetWindowPos", functions)
        self.assertNotIn("BringWindowToTop", functions)
        style_name = "GetWindowLongPtrW" if ctypes.sizeof(ctypes.c_void_p) == 8 else "GetWindowLongW"
        self.assertIs(functions[style_name].restype, ctypes.c_ssize_t)


class WindowStateTests(unittest.TestCase):
    def assert_no_window_changes(self, api):
        self.assertFalse(any(call[0] in {"restore", "activate", "pinned", "topmost", "wait"} for call in api.calls))

    def test_foreground_state_is_read_without_requesting_activation(self):
        api = FakeAPI()
        api.foreground = TARGET
        self.assertTrue(activation.is_window_foreground(TARGET, api=api))
        self.assert_no_window_changes(api)

    def test_minimized_window_is_never_restored_by_state_check(self):
        api = FakeAPI(minimized=True)
        api.foreground = TARGET
        self.assertFalse(activation.is_window_foreground(TARGET, api=api))
        self.assertTrue(api.minimized)
        self.assert_no_window_changes(api)

    def test_background_window_is_not_brought_forward(self):
        api = FakeAPI()
        self.assertFalse(activation.is_window_foreground(TARGET, api=api))
        self.assertEqual(api.foreground, OTHER)
        self.assert_no_window_changes(api)

    def test_recycled_handle_is_not_confirmed(self):
        api = FakeAPI()
        api.foreground = TARGET
        with patch.object(api, "window_identity", side_effect=[(11, 23), (99, 101)]):
            self.assertFalse(activation.is_window_foreground(TARGET, api=api))
        self.assert_no_window_changes(api)

    def test_invalid_handles_do_not_access_window_api(self):
        api = FakeAPI()
        for hwnd in (None, False, True, "42", 0, -1, 1 << (ctypes.sizeof(ctypes.c_void_p) * 8)):
            with self.subTest(hwnd=hwnd):
                self.assertFalse(activation.is_window_foreground(hwnd, api=api))
        self.assertEqual(api.calls, [])

    def test_unavailable_window_state_is_false_without_focus_fallback(self):
        api = FakeAPI()
        with patch.object(api, "is_minimized", side_effect=OSError("Window closed")):
            self.assertFalse(activation.is_window_foreground(TARGET, api=api))
        self.assert_no_window_changes(api)


class FakeProcessWindows:
    def __init__(self, windows, foreground=0):
        self.windows = windows
        self.foreground = foreground
        self.calls = []

    def foreground_window(self):
        self.calls.append(("foreground",))
        return self.foreground

    def enumerate_windows(self):
        self.calls.append(("enumerate",))
        return tuple(self.windows)

    def is_window(self, hwnd):
        self.calls.append(("valid", hwnd))
        return hwnd in self.windows and self.windows[hwnd].get("valid", True)

    def window_identity(self, hwnd):
        self.calls.append(("identity", hwnd))
        return (11, self.windows[hwnd]["pid"])

    def is_visible(self, hwnd):
        self.calls.append(("visible", hwnd))
        return self.windows[hwnd].get("visible", True)

    def is_tool_window(self, hwnd):
        self.calls.append(("tool", hwnd))
        return self.windows[hwnd].get("tool", False)

    def owner_window(self, hwnd):
        self.calls.append(("owner", hwnd))
        return self.windows[hwnd].get("owner", 0)


class ProcessWindowTests(unittest.TestCase):
    def test_prefers_visible_foreground_window_of_exact_process(self):
        api = FakeProcessWindows({1: {"pid": 23}, TARGET: {"pid": 23, "owner": 1}}, foreground=TARGET)
        self.assertEqual(activation.find_process_window(23, api=api), TARGET)
        self.assertNotIn(("enumerate",), api.calls)

    def test_skips_other_process_hidden_tool_and_owned_when_main_exists(self):
        api = FakeProcessWindows({
            1: {"pid": 999},
            2: {"pid": 23, "visible": False},
            3: {"pid": 23, "tool": True},
            4: {"pid": 23, "owner": 5},
            TARGET: {"pid": 23},
        }, foreground=1)
        self.assertEqual(activation.find_process_window(23, api=api), TARGET)

    def test_falls_back_to_owned_window_of_same_process(self):
        api = FakeProcessWindows({1: {"pid": 999}, TARGET: {"pid": 23, "owner": 7}})
        self.assertEqual(activation.find_process_window(23, api=api), TARGET)

    def test_never_returns_window_of_other_process_or_invisible_window(self):
        api = FakeProcessWindows({1: {"pid": 999}, 2: {"pid": 23, "visible": False}, 3: {"pid": 23, "tool": True}}, foreground=1)
        self.assertIsNone(activation.find_process_window(23, api=api))

    def test_ignores_window_that_disappeared_during_enumeration(self):
        api = FakeProcessWindows({1: {"pid": 23, "valid": False}, TARGET: {"pid": 23}})
        self.assertEqual(activation.find_process_window(23, api=api), TARGET)

    def test_invalid_process_id_never_accesses_api(self):
        api = FakeProcessWindows({})
        for process_id in (None, True, False, "23", 0, -1, 1 << 32):
            with self.subTest(process_id=process_id):
                self.assertIsNone(activation.find_process_window(process_id, api=api))
        self.assertEqual(api.calls, [])

    def test_unsupported_platform_returns_none_without_loading_user32(self):
        with patch("window_activation.os.name", "posix"), patch("window_activation.ctypes.WinDLL", create=True) as loader:
            self.assertIsNone(activation.find_process_window(23))
        loader.assert_not_called()


if __name__ == "__main__":
    unittest.main()
