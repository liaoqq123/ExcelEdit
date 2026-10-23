"""只读查询目标窗口和所属进程；导入时不加载 Win32，不改变焦点或窗口状态。"""

from __future__ import annotations

import ctypes
from ctypes import wintypes
import os
from typing import Protocol


GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
GW_OWNER = 4


class WindowStateAPI(Protocol):
    """可注入的只读窗口状态后端。"""

    def is_window(self, hwnd: int) -> bool: ...
    def window_identity(self, hwnd: int) -> tuple[int, int]: ...
    def is_minimized(self, hwnd: int) -> bool: ...
    def foreground_window(self) -> int: ...


class ProcessWindowAPI(Protocol):
    """只读进程窗口查询后端，不包含激活或等待操作。"""

    def is_window(self, hwnd: int) -> bool: ...
    def window_identity(self, hwnd: int) -> tuple[int, int]: ...
    def foreground_window(self) -> int: ...
    def enumerate_windows(self) -> tuple[int, ...]: ...
    def is_visible(self, hwnd: int) -> bool: ...
    def owner_window(self, hwnd: int) -> int: ...
    def is_tool_window(self, hwnd: int) -> bool: ...


class _WindowsAPI:
    """惰性加载且所有 HWND/LONG_PTR 都按指针宽度绑定。"""

    def __init__(self) -> None:
        if os.name != "nt":
            raise OSError("Window activation is only supported on Windows")
        self._user32 = ctypes.WinDLL("user32", use_last_error=True)
        hwnd_type = ctypes.c_void_p
        self._bind("IsWindow", (hwnd_type,), wintypes.BOOL)
        self._bind("IsWindowVisible", (hwnd_type,), wintypes.BOOL)
        self._bind("IsIconic", (hwnd_type,), wintypes.BOOL)
        self._bind("GetForegroundWindow", (), hwnd_type)
        self._bind("GetWindowThreadProcessId", (hwnd_type, ctypes.POINTER(wintypes.DWORD)), wintypes.DWORD)
        self._bind("GetWindow", (hwnd_type, wintypes.UINT), hwnd_type)
        callback_factory = getattr(ctypes, "WINFUNCTYPE", ctypes.CFUNCTYPE)
        self._enum_callback_type = callback_factory(wintypes.BOOL, hwnd_type, ctypes.c_ssize_t)
        self._bind("EnumWindows", (self._enum_callback_type, ctypes.c_ssize_t), wintypes.BOOL)
        self._style_function_name = "GetWindowLongPtrW" if ctypes.sizeof(ctypes.c_void_p) == 8 else "GetWindowLongW"
        self._bind(self._style_function_name, (hwnd_type, ctypes.c_int), ctypes.c_ssize_t)

    def _bind(self, name: str, argument_types: tuple[object, ...], result_type: object) -> None:
        function = getattr(self._user32, name)
        function.argtypes = argument_types
        function.restype = result_type

    def is_window(self, hwnd: int) -> bool:
        return bool(self._user32.IsWindow(hwnd))

    def window_identity(self, hwnd: int) -> tuple[int, int]:
        process_id = wintypes.DWORD()
        thread_id = self._user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
        return int(thread_id), int(process_id.value)

    def enumerate_windows(self) -> tuple[int, ...]:
        handles: list[int] = []

        @self._enum_callback_type
        def collect(hwnd: int, _parameter: int) -> bool:
            if hwnd:
                handles.append(int(hwnd))
            return True

        if not self._user32.EnumWindows(collect, 0):
            raise OSError(ctypes.get_last_error(), "Could not enumerate desktop windows")
        return tuple(handles)

    def is_visible(self, hwnd: int) -> bool:
        return bool(self._user32.IsWindowVisible(hwnd))

    def owner_window(self, hwnd: int) -> int:
        return int(self._user32.GetWindow(hwnd, GW_OWNER) or 0)

    def is_minimized(self, hwnd: int) -> bool:
        return bool(self._user32.IsIconic(hwnd))

    def foreground_window(self) -> int:
        return int(self._user32.GetForegroundWindow() or 0)

    def _extended_style(self, hwnd: int) -> int:
        ctypes.set_last_error(0)
        style = getattr(self._user32, self._style_function_name)(hwnd, GWL_EXSTYLE)
        error = ctypes.get_last_error()
        if not style and error:
            raise OSError(error, "Could not read the target window style")
        return int(style)

    def is_tool_window(self, hwnd: int) -> bool:
        return bool(self._extended_style(hwnd) & WS_EX_TOOLWINDOW)

def _same_window(api: WindowStateAPI, hwnd: int, identity: tuple[int, int]) -> bool:
    return api.is_window(hwnd) and api.window_identity(hwnd) == identity


def is_window_foreground(hwnd: int, *, api: WindowStateAPI | None = None) -> bool:
    """只读核对目标是否在前台且未最小化，不恢复窗口、不请求激活。"""
    if isinstance(hwnd, bool) or not isinstance(hwnd, int) or not 0 < hwnd < 1 << (ctypes.sizeof(ctypes.c_void_p) * 8):
        return False
    try:
        backend = api if api is not None else _WindowsAPI()
        if not backend.is_window(hwnd):
            return False
        identity = backend.window_identity(hwnd)
        return bool(all(identity) and not backend.is_minimized(hwnd) and
                    backend.foreground_window() == hwnd and _same_window(backend, hwnd, identity))
    except Exception:
        return False


def _visible_process_window(api: ProcessWindowAPI, hwnd: int, process_id: int) -> bool:
    return bool(hwnd and api.is_window(hwnd) and api.window_identity(hwnd)[1] == process_id and api.is_visible(hwnd))


def find_process_window(process_id: int, *, api: ProcessWindowAPI | None = None) -> int | None:
    """只读查找指定进程的可见窗口；不等待、不激活、不按标题匹配。

    优先保留该进程当前的前台窗口，否则选择无 owner 的非工具顶层窗口；
    若只有 owned 窗口可用，回退到该进程的第一个可见非工具 owned 窗口。
    """

    if isinstance(process_id, bool) or not isinstance(process_id, int) or not 0 < process_id <= 0xFFFFFFFF:
        return None
    try:
        backend = api if api is not None else _WindowsAPI()
        try:
            foreground = backend.foreground_window()
            if _visible_process_window(backend, foreground, process_id):
                return foreground
        except Exception:
            pass
        owned_candidate: int | None = None
        for hwnd in backend.enumerate_windows():
            try:
                if not _visible_process_window(backend, hwnd, process_id) or backend.is_tool_window(hwnd):
                    continue
                if not backend.owner_window(hwnd):
                    return hwnd
                if owned_candidate is None:
                    owned_candidate = hwnd
            except Exception:
                # Windows may disappear while EnumWindows' snapshot is queried.
                continue
        if owned_candidate is not None and _visible_process_window(backend, owned_candidate, process_id):
            return owned_candidate
    except Exception:
        pass
    return None
