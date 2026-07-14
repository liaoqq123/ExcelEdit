"""普通文件名检索功能。

这个模块只查固定目录下的文件名和完整路径，不读取文件内容。
"""

import os
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from threading import Event


@dataclass(frozen=True)
class FileSearchResult:
    """文件检索结果，供界面列表直接展示和双击打开目录。"""

    path: Path
    file_name: str
    file_address: str


class FileSearchError(RuntimeError):
    """文件检索路径校验失败时抛出的错误。"""

    pass


def search_files(folder_path: str | Path, keyword: str = "", cancel_event: Event | None = None) -> list[FileSearchResult]:
    """递归检索文件名和文件地址，关键词为空时返回全部文件。"""

    folder = _validate_folder(folder_path)
    normalized_keyword = keyword.strip().lower()
    results: list[FileSearchResult] = []

    for file_path in _iter_file_paths(folder, cancel_event):
        if _is_cancelled(cancel_event):
            break
        file_name = file_path.name
        file_address = str(file_path)
        if normalized_keyword and not _matches_keyword(file_name, file_address, normalized_keyword):
            continue
        results.append(FileSearchResult(path=file_path, file_name=file_name, file_address=file_address))

    results.sort(key=lambda result: result.file_address.lower())
    return results


def _iter_file_paths(folder: Path, cancel_event: Event | None) -> Iterator[Path]:
    """用 os.scandir 递归遍历目录，比逐个 Path.iterdir 更轻量。"""

    stack = [folder]

    while stack:
        if _is_cancelled(cancel_event):
            break
        current_folder = stack.pop()
        try:
            with os.scandir(current_folder) as entries:
                for entry in entries:
                    if _is_cancelled(cancel_event):
                        break
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            stack.append(Path(entry.path))
                            continue
                        if entry.is_file(follow_symlinks=False):
                            yield Path(entry.path)
                    except OSError:
                        continue
        except OSError:
            continue


def _matches_keyword(file_name: str, file_address: str, keyword: str) -> bool:
    """判断关键词是否出现在文件名或完整路径中。"""

    return keyword in file_name.lower() or keyword in file_address.lower()


def _validate_folder(folder_path: str | Path) -> Path:
    """确认待检索路径存在且是文件夹。"""

    folder = Path(folder_path)
    if not folder.exists():
        raise FileSearchError(f"Folder does not exist: {folder}")
    if not folder.is_dir():
        raise FileSearchError(f"Not a folder: {folder}")
    return folder


def _is_cancelled(cancel_event: Event | None) -> bool:
    """后台检索线程轮询这个状态来尽快响应取消按钮。"""

    return cancel_event is not None and cancel_event.is_set()
