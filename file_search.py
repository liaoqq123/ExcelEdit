"""普通文件名检索功能。

这个模块只查固定目录下的文件名和完整路径，不读取文件内容。
"""

import os
import re
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from pathlib import Path
from threading import Event


DEFAULT_MAX_FILE_SEARCH_RESULTS = 20_000
FILE_SUFFIX_SEPARATOR_PATTERN = re.compile(r"[\s,，;；]+")


@dataclass(frozen=True)
class FileSearchResult:
    """文件检索结果，供界面列表直接展示和双击打开目录。"""

    path: Path
    file_name: str
    file_address: str
    is_directory: bool = False


class FileSearchError(RuntimeError):
    """文件检索路径校验失败时抛出的错误。"""

    pass


def search_files(
    folder_path: str | Path,
    keyword: str = "",
    cancel_event: Event | None = None,
    max_results: int | None = DEFAULT_MAX_FILE_SEARCH_RESULTS,
    suffixes: str | Iterable[str] | None = None,
    exact_match: bool = False,
    include_folders: bool = False,
) -> list[FileSearchResult]:
    """递归检索文件和可选的文件夹，可按一个或多个文件后缀过滤文件。"""

    folder = _validate_folder(folder_path)
    normalized_keyword = keyword.strip().lower()
    normalized_suffixes = normalize_file_suffixes(suffixes)
    results: list[FileSearchResult] = []

    for result_path, is_directory in _iter_search_paths(folder, cancel_event, include_folders):
        if _is_cancelled(cancel_event):
            break
        file_name = result_path.name
        file_address = str(result_path)
        if not is_directory and normalized_suffixes and not file_name.lower().endswith(normalized_suffixes):
            continue
        if normalized_keyword and not _matches_keyword(
            file_name,
            file_address,
            normalized_keyword,
            exact_match,
        ):
            continue
        results.append(
            FileSearchResult(
                path=result_path,
                file_name=file_name,
                file_address=file_address,
                is_directory=is_directory,
            )
        )
        if max_results is not None and len(results) >= max_results:
            break

    results.sort(key=lambda result: result.file_address.lower())
    return results


def normalize_file_suffixes(suffixes: str | Iterable[str] | None) -> tuple[str, ...]:
    """规范化用户输入的后缀，兼容 xlsx、.xlsx 和 *.xlsx 等写法。"""

    if suffixes is None:
        return ()
    raw_suffixes = FILE_SUFFIX_SEPARATOR_PATTERN.split(suffixes) if isinstance(suffixes, str) else suffixes
    normalized_suffixes: list[str] = []
    for raw_suffix in raw_suffixes:
        suffix = str(raw_suffix).strip().lower()
        if suffix.startswith("*"):
            suffix = suffix[1:]
        if not suffix:
            continue
        if not suffix.startswith("."):
            suffix = f".{suffix}"
        if suffix != "." and suffix not in normalized_suffixes:
            normalized_suffixes.append(suffix)
    return tuple(normalized_suffixes)


def _iter_search_paths(
    folder: Path,
    cancel_event: Event | None,
    include_folders: bool,
) -> Iterator[tuple[Path, bool]]:
    """用 os.scandir 递归遍历文件，并按需返回子文件夹。"""

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
                            directory_path = Path(entry.path)
                            stack.append(directory_path)
                            if include_folders:
                                yield directory_path, True
                            continue
                        if entry.is_file(follow_symlinks=False):
                            yield Path(entry.path), False
                    except OSError:
                        continue
        except OSError:
            continue


def _matches_keyword(
    file_name: str,
    file_address: str,
    keyword: str,
    exact_match: bool = False,
) -> bool:
    """判断关键词是否出现在文件名或完整路径中。"""

    normalized_values = (file_name.lower(), file_address.lower())
    if exact_match:
        return keyword in normalized_values
    return any(keyword in value for value in normalized_values)


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
