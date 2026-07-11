import os
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from threading import Event

from excel_common import ExcelReadError


@dataclass(frozen=True)
class FileSearchResult:
    path: Path
    file_name: str
    file_address: str


def search_files(folder_path: str | Path, keyword: str = "", cancel_event: Event | None = None) -> list[FileSearchResult]:
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
    return keyword in file_name.lower() or keyword in file_address.lower()


def _validate_folder(folder_path: str | Path) -> Path:
    folder = Path(folder_path)
    if not folder.exists():
        raise ExcelReadError(f"Folder does not exist: {folder}")
    if not folder.is_dir():
        raise ExcelReadError(f"Not a folder: {folder}")
    return folder


def _is_cancelled(cancel_event: Event | None) -> bool:
    return cancel_event is not None and cancel_event.is_set()
