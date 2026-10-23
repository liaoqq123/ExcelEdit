"""打开目标并分别确认位置与前台状态；所有调用由界面后台线程执行。"""

from __future__ import annotations

import base64
import ctypes
from dataclasses import dataclass
import json
import os
from pathlib import Path
import subprocess
import sys
from threading import Lock
from time import monotonic, sleep

from excel_common import EXCEL_EXTENSIONS
from window_activation import find_process_window, is_window_foreground
from windows_shell import launch_path, reveal_item


OPEN_TIMEOUT_SECONDS = 20
EXPLORER_WAIT_SECONDS = 5
EXCEL_LOOKUP_WAIT_SECONDS = 8
_EXCEL_LOCK = Lock()


@dataclass(frozen=True)
class OpenResult:
    opened: bool
    located: bool = False
    foreground: bool = False
    hwnd: int = 0
    message: str = ""


def _absolute_path(path: Path) -> Path:
    # abspath preserves a user's visible directory alias instead of resolving a junction.
    return Path(os.path.abspath(os.fspath(path)))


def _com_hwnd(value) -> int:
    """Shell/Office 的 signed Long HWND 保留低 32 位；不截断 64 位句柄。"""
    if isinstance(value, bool):
        return 0
    try:
        handle = int(value or 0)
    except (TypeError, ValueError):
        return 0
    if -(1 << 31) <= handle < 0:
        handle &= 0xFFFFFFFF
    return handle if 0 < handle < (1 << (ctypes.sizeof(ctypes.c_void_p) * 8)) else 0


def _run_powershell(script: str, payload: dict, timeout: float = OPEN_TIMEOUT_SECONDS) -> dict:
    """固定脚本与 JSON 标准输入分离，路径不作为 PowerShell 源码拼接。"""
    executable = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    command = str(executable) if executable.is_file() else "powershell.exe"
    prefix = "$ErrorActionPreference='Stop'; $ProgressPreference='SilentlyContinue'; [Console]::OutputEncoding=[Text.UTF8Encoding]::new($false); $request=[Console]::In.ReadToEnd() | ConvertFrom-Json;\n"
    encoded = base64.b64encode((prefix + script).encode("utf-16-le")).decode("ascii")
    completed = subprocess.run(
        [command, "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
        input=json.dumps(payload, ensure_ascii=True), capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=timeout,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if completed.returncode:
        raise OSError((completed.stderr or completed.stdout).strip() or "Windows 打开助手执行失败")
    try:
        result = json.loads(completed.stdout.strip())
    except (ValueError, TypeError) as exc:
        raise OSError("Windows 打开助手未返回有效的窗口信息") from exc
    if not isinstance(result, dict):
        raise OSError("Windows 打开助手返回的窗口信息无效")
    return result


_EXPLORER_SCRIPT = r"""
$folder = [IO.Path]::GetFullPath([string]$request.folder).TrimEnd([char]'\')
$target = [string]$request.target
$lastError = ''
$matchedHandle = 0
$shell = New-Object -ComObject Shell.Application
$deadline = [DateTime]::UtcNow.AddSeconds([double]$request.wait)
do {
    foreach ($window in @($shell.Windows())) {
        try {
            $actual = [IO.Path]::GetFullPath([string]$window.Document.Folder.Self.Path).TrimEnd([char]'\')
            if (-not [string]::Equals($actual, $folder, [StringComparison]::OrdinalIgnoreCase)) { continue }
            $selected = [string]::IsNullOrEmpty($target)
            if (-not $selected) {
                $leaf = [IO.Path]::GetFileName($target.TrimEnd([char]'\'))
                $item = $window.Document.Folder.ParseName($leaf)
                if ($null -ne $item) {
                    # SELECT | DESELECTOTHERS | ENSUREVISIBLE | FOCUSED
                    [void]$window.Document.SelectItem($item, 29)
                    foreach ($chosen in @($window.Document.SelectedItems())) {
                        if ([string]::Equals([string]$chosen.Path, $target, [StringComparison]::OrdinalIgnoreCase)) {
                            $selected = $true
                            break
                        }
                    }
                }
            }
            $matchedHandle = [long]$window.HWND
            if (-not $selected) { continue }
            [pscustomobject]@{ hwnd=$matchedHandle; located=$selected; error='' } | ConvertTo-Json -Compress
            exit 0
        } catch { $lastError=$_.Exception.Message }
    }
    Start-Sleep -Milliseconds 150
} while ([DateTime]::UtcNow -lt $deadline)
[pscustomobject]@{ hwnd=$matchedHandle; located=$false; error=$lastError } | ConvertTo-Json -Compress
"""


def _explorer_window(folder: Path, target: Path | None) -> dict:
    return _run_powershell(_EXPLORER_SCRIPT, {
        "folder": str(folder), "target": str(target) if target else "", "wait": EXPLORER_WAIT_SECONDS,
    }, timeout=EXPLORER_WAIT_SECONDS + 5)


def open_folder(path: Path, *, reveal: bool) -> OpenResult:
    """reveal 定位文件或子文件夹；普通打开则进入目标文件夹。"""
    target = _absolute_path(path)
    exists = target.exists()
    if not reveal:
        if not target.is_dir():
            raise FileNotFoundError(f"文件夹不存在或无法访问：{target}")
        folder, selection = target, None
    elif target.parent == target:
        if not target.is_dir():
            raise FileNotFoundError(f"目录不存在或无法访问：{target}")
        folder, selection = target, None
    else:
        folder, selection = target.parent, target if exists else None
        if not folder.is_dir():
            raise FileNotFoundError(f"所在目录不存在或无法访问：{folder}")

    message = "目标已移动或删除，已打开仍存在的父目录" if reveal and not exists else ""
    if selection is not None:
        try:
            reveal_item(selection)
        except OSError:
            # Selection can fail independently of opening the containing folder.
            launch_path(folder)
    else:
        launch_path(folder)

    try:
        window = _explorer_window(folder, selection)
    except (OSError, subprocess.SubprocessError) as exc:
        return OpenResult(True, message=message or f"已提交打开请求，无法核对资源管理器窗口：{exc}")
    hwnd = _com_hwnd(window.get("hwnd"))
    located = bool(window.get("located")) and (exists or not reveal)
    if not hwnd:
        message = message or "已提交打开请求，但未找到对应目录窗口"
    elif selection is not None and not located:
        message = "目录已打开，但未能选中目标文件或子文件夹"
    return OpenResult(True, located, is_window_foreground(hwnd) if hwnd else False, hwnd, message)


_EXCEL_LOOKUP_SCRIPT = r"""
# Read the running object table without binding or starting an OLE document.
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ExcelEdit {
    public static class RunningWorkbooks {
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable table);
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx context);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsIconic(IntPtr hwnd);

        public static bool CanNavigate(long handle) {
            if (handle >= Int32.MinValue && handle < 0) handle &= UInt32.MaxValue;
            if (handle <= 0) return false;
            IntPtr hwnd = IntPtr.Size == 8 ? new IntPtr(handle) : new IntPtr(unchecked((int)handle));
            return IsWindow(hwnd) && !IsIconic(hwnd) && GetForegroundWindow() == hwnd;
        }

        private static void Release(object value) {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        public static object Find(string path) {
            IRunningObjectTable table = null;
            IEnumMoniker names = null;
            IBindCtx context = null;
            try {
                Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out table));
                Marshal.ThrowExceptionForHR(CreateBindCtx(0, out context));
                table.EnumRunning(out names);
                IMoniker[] current = new IMoniker[1];
                while (names.Next(1, current, IntPtr.Zero) == 0) {
                    try {
                        string displayName;
                        current[0].GetDisplayName(context, null, out displayName);
                        if (!String.Equals(Path.GetFullPath(displayName), path, StringComparison.OrdinalIgnoreCase)) continue;
                        object workbook;
                        if (table.GetObject(current[0], out workbook) == 0 && workbook != null) return workbook;
                    } catch (Exception) {
                        // Unrelated ROT entries and disappearing windows are normal.
                    } finally {
                        Release(current[0]);
                        current[0] = null;
                    }
                }
                return null;
            } finally {
                Release(context);
                Release(names);
                Release(table);
            }
        }
    }
}
'@

function Find-OpenedWorkbook([string]$filePath) {
    try {
        $candidate = [ExcelEdit.RunningWorkbooks]::Find($filePath)
        if ($null -ne $candidate -and [string]::Equals(
            [IO.Path]::GetFullPath([string]$candidate.FullName), $filePath,
            [StringComparison]::OrdinalIgnoreCase)) { return ,$candidate }
    } catch { }
    # Some Office versions register the application before the file moniker.
    try {
        $application = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        foreach ($candidate in @($application.Workbooks)) {
            if ([string]::Equals([IO.Path]::GetFullPath([string]$candidate.FullName),
                $filePath, [StringComparison]::OrdinalIgnoreCase)) { return ,$candidate }
        }
    } catch { }
    return $null
}

function Test-TargetForeground([long]$handle) {
    return [ExcelEdit.RunningWorkbooks]::CanNavigate($handle)
}
"""


_EXCEL_NAVIGATION_SCRIPT = r"""
$filePath = [IO.Path]::GetFullPath([string]$request.path)
$sheetName = [string]$request.sheet
$workbook = $null
$bookWindow = $null
# The caller has already submitted exactly one normal Shell open request.
$opened = $true
$located = $false
$handle = 0
$problem = ''
try {
    $deadline = [DateTime]::UtcNow.AddSeconds([double]$request.wait)
    do {
        $candidate = Find-OpenedWorkbook $filePath
        if ($null -ne $candidate) {
            try {
                if ([string]::Equals([IO.Path]::GetFullPath([string]$candidate.FullName),
                    $filePath, [StringComparison]::OrdinalIgnoreCase) -and $candidate.Application.Visible) {
                    foreach ($window in @($candidate.Windows)) {
                        if ($window.Visible) {
                            if ($null -eq $bookWindow) {
                                $workbook = $candidate
                                $bookWindow = $window
                            }
                            if (Test-TargetForeground ([long]$window.Hwnd)) {
                                $workbook = $candidate
                                $bookWindow = $window
                                break
                            }
                        }
                    }
                }
            } catch { }
        }
        if ($null -ne $bookWindow) { break }
        if ([DateTime]::UtcNow -ge $deadline) { break }
        Start-Sleep -Milliseconds 150
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($null -eq $bookWindow) {
        throw '文件已交给系统打开，但未连接到可见的目标工作簿；自动定位未完成'
    }
    # Navigation never changes application or document visibility. Shell owns
    # opening; an unavailable COM connection must not leave an invisible book.
    $excel = $workbook.Application
    $handle = [long]$bookWindow.Hwnd
    if (-not (Test-TargetForeground $handle)) {
        throw '目标窗口已最小化或切换，已跳过后续自动定位'
    }
    if (-not [string]::IsNullOrEmpty($sheetName)) {
        $worksheet = $workbook.Worksheets.Item($sheetName)
        if (-not (Test-TargetForeground $handle)) {
            throw '目标窗口已最小化或切换，已跳过后续自动定位'
        }
        [void]$worksheet.Activate()
        if ($null -ne $request.row -and $null -ne $request.column) {
            $cell = $worksheet.Cells.Item([int]$request.row, [int]$request.column)
            if (-not (Test-TargetForeground $handle)) {
                throw '目标窗口已最小化或切换，已跳过后续自动定位'
            }
            [void]$excel.Goto($cell, $true)
        }
    }
    $located = $true
} catch { $problem = $_.Exception.Message }
[pscustomobject]@{ opened=$opened; located=$located; hwnd=$handle; error=$problem } | ConvertTo-Json -Compress
"""


_EXCEL_SCRIPT = _EXCEL_LOOKUP_SCRIPT + _EXCEL_NAVIGATION_SCRIPT


def open_excel(path: Path, sheet: str | None = None, row: int | None = None, column: int | None = None) -> OpenResult:
    target = _absolute_path(path)
    if not target.is_file():
        raise FileNotFoundError(f"Excel 文件不存在或无法访问：{target}")
    if ((row is None) != (column is None) or
        (row is not None and (isinstance(row, bool) or isinstance(column, bool) or
                             not isinstance(row, int) or not isinstance(column, int) or
                             not (1 <= row <= 1_048_576 and 1 <= column <= 16_384)))):
        raise ValueError("Excel 单元格位置无效")
    if not sheet:
        return _open_associated(target)
    # Opening is independent of COM, like double-clicking the file in Explorer.
    # Even a busy navigation request must not prevent the file from opening.
    launch_path(target)
    if not _EXCEL_LOCK.acquire(timeout=OPEN_TIMEOUT_SECONDS):
        return OpenResult(True, message="文件已交给系统打开；其他 Excel 定位尚未完成，本次自动定位未执行")
    try:
        try:
            response = _run_powershell(_EXCEL_SCRIPT, {
                "path": str(target), "sheet": sheet, "row": row, "column": column,
                "wait": EXCEL_LOOKUP_WAIT_SECONDS,
            })
        except subprocess.TimeoutExpired:
            return OpenResult(True, message="文件已交给系统打开；Excel 自动定位超时，请检查 Excel 窗口")
        except (OSError, subprocess.SubprocessError) as exc:
            return OpenResult(True, message=f"文件已交给系统打开；Excel 自动定位不可用：{exc}")
        hwnd = _com_hwnd(response.get("hwnd"))
        located = bool(response.get("located")) and bool(hwnd)
        message = str(response.get("error") or ("" if located else "自动定位未完成，请查看已打开的文件"))
        return OpenResult(True, located, is_window_foreground(hwnd) if hwnd else False, hwnd, message)
    finally:
        _EXCEL_LOCK.release()


def _open_associated(path: Path, *, located: bool = True, message: str = "") -> OpenResult:
    process_id = launch_path(path)
    hwnd = 0
    if process_id:
        deadline = monotonic() + 3
        while monotonic() < deadline:
            hwnd = find_process_window(process_id) or 0
            if hwnd:
                break
            sleep(0.1)
    return OpenResult(True, located, is_window_foreground(hwnd) if hwnd else False, hwnd, message)


def open_target(path: Path, sheet: str | None = None, row: int | None = None, column: int | None = None, *, reveal: bool = False) -> OpenResult:
    target = _absolute_path(path)
    if not sys.platform.startswith("win"):
        destination = target.parent if reveal else target
        command = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.Popen([command, str(destination)])
        return OpenResult(True, message="已提交系统打开请求")
    if reveal or target.is_dir():
        return open_folder(target, reveal=reveal)
    if not target.is_file():
        raise FileNotFoundError(f"文件不存在或无法访问：{target}")
    if target.suffix.lower() in EXCEL_EXTENSIONS:
        return open_excel(target, sheet, row, column)
    return _open_associated(target)
