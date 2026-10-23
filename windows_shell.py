"""Windows Shell 打开接口；原生 DLL 仅在实际调用时加载。"""

import ctypes
import os
from contextlib import contextmanager
from functools import lru_cache
from pathlib import Path
from types import SimpleNamespace


DWORD = ctypes.c_uint32
HRESULT = ctypes.c_int32
HANDLE = ctypes.c_void_p
COINIT_APARTMENTTHREADED = 0x2
SHELL_OPEN_FLAGS = 0x100 | 0x40 | 0x400
SW_SHOWNORMAL = 1


class _ShellDisplayHandle(ctypes.Union):
    _fields_ = [("hIcon", HANDLE), ("hMonitor", HANDLE)]


class SHELLEXECUTEINFOW(ctypes.Structure):
    """使用固定宽度 Win32 整数与指针宽度句柄，兼容 32/64 位。"""

    _anonymous_ = ("display",)
    _fields_ = [
        ("cbSize", DWORD), ("fMask", DWORD), ("hwnd", HANDLE),
        ("lpVerb", ctypes.c_wchar_p), ("lpFile", ctypes.c_wchar_p),
        ("lpParameters", ctypes.c_wchar_p), ("lpDirectory", ctypes.c_wchar_p),
        ("nShow", ctypes.c_int32), ("hInstApp", HANDLE), ("lpIDList", HANDLE),
        ("lpClass", ctypes.c_wchar_p), ("hkeyClass", HANDLE),
        ("dwHotKey", DWORD), ("display", _ShellDisplayHandle), ("hProcess", HANDLE),
    ]


@lru_cache(maxsize=1)
def _native_api():
    if os.name != "nt":
        raise OSError("Windows Shell 接口仅支持 Windows。")
    ole32 = ctypes.WinDLL("ole32", use_last_error=True)
    shell32 = ctypes.WinDLL("shell32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    signatures = [
        (ole32, "CoInitializeEx", [HANDLE, DWORD], HRESULT),
        (ole32, "CoUninitialize", [], None),
        (ole32, "CoTaskMemFree", [HANDLE], None),
        (shell32, "SHParseDisplayName", [ctypes.c_wchar_p, HANDLE, ctypes.POINTER(HANDLE), DWORD, ctypes.POINTER(DWORD)], HRESULT),
        (shell32, "SHOpenFolderAndSelectItems", [HANDLE, DWORD, ctypes.POINTER(HANDLE), DWORD], HRESULT),
        (shell32, "ShellExecuteExW", [ctypes.POINTER(SHELLEXECUTEINFOW)], ctypes.c_int32),
        (kernel32, "GetProcessId", [HANDLE], DWORD),
        (kernel32, "CloseHandle", [HANDLE], ctypes.c_int32),
    ]
    functions = {"GetLastError": ctypes.get_last_error}
    for library, name, arguments, result in signatures:
        function = getattr(library, name)
        function.argtypes = arguments
        function.restype = result
        functions[name] = function
    return SimpleNamespace(**functions)


def _check_hresult(operation: str, result) -> None:
    value = HRESULT(getattr(result, "value", result)).value
    if value < 0:
        raise OSError(f"{operation} 失败，HRESULT 0x{value & 0xffffffff:08X}")


@contextmanager
def _com_sta(api):
    _check_hresult("CoInitializeEx", api.CoInitializeEx(None, COINIT_APARTMENTTHREADED))
    try:
        yield
    finally:
        api.CoUninitialize()


def reveal_item(path: Path, *, api=None) -> None:
    """在真实父目录中选中文件或子文件夹，避免命令行路径解析歧义。"""

    api = _native_api() if api is None else api
    pidl = HANDLE()
    with _com_sta(api):
        try:
            _check_hresult("SHParseDisplayName", api.SHParseDisplayName(str(path), None, ctypes.byref(pidl), 0, None))
            if not pidl.value:
                raise OSError("SHParseDisplayName 未返回有效的项目标识。")
            # cidl=0 时，完整 PIDL 表示要在其父目录中选中的单个项目。
            _check_hresult("SHOpenFolderAndSelectItems", api.SHOpenFolderAndSelectItems(pidl, 0, None, 0))
        finally:
            if pidl.value:
                api.CoTaskMemFree(pidl)


def launch_path(path: Path, *, api=None) -> int | None:
    """通过默认程序打开路径，返回可用的进程 ID，并始终释放进程句柄。"""

    api = _native_api() if api is None else api
    info = SHELLEXECUTEINFOW()
    info.cbSize = ctypes.sizeof(info)
    info.fMask = SHELL_OPEN_FLAGS
    info.lpVerb = "open"
    info.lpFile = str(path)
    info.nShow = SW_SHOWNORMAL
    with _com_sta(api):
        try:
            if not api.ShellExecuteExW(ctypes.byref(info)):
                code = api.GetLastError()
                raise OSError(code, f"ShellExecuteExW 打开失败，Windows 错误 {code}")
            if not info.hProcess:
                return None
            process_id = api.GetProcessId(info.hProcess)
            return int(getattr(process_id, "value", process_id)) or None
        finally:
            if info.hProcess:
                api.CloseHandle(info.hProcess)
