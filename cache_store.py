"""应用缓存的统一读取、迁移和原子写入。"""

import json
import os
import sys
import tempfile
from pathlib import Path
from threading import RLock


APP_CACHE_DIRECTORY_NAME = "NailongSearchMaster"
CACHE_FILE_NAME = "cache_data.json"


def get_app_directory() -> Path:
    """获取源码或打包程序所在目录。"""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_user_cache_file() -> Path:
    """返回当前用户可写的缓存路径，避免安装目录权限导致保存失败。"""

    return get_user_cache_directory() / CACHE_FILE_NAME


def get_user_cache_directory() -> Path:
    """返回当前用户的软件数据目录，供缓存和用户选择的资源共用。"""

    local_app_data = os.environ.get("LOCALAPPDATA")
    base_directory = Path(local_app_data) if local_app_data else Path.home() / ".local" / "share"
    return base_directory / APP_CACHE_DIRECTORY_NAME


CACHE_FILE = get_user_cache_file()
LEGACY_CACHE_FILE = get_app_directory() / CACHE_FILE_NAME
_CACHE_LOCK = RLock()


def read_cache_data(
    cache_file: Path = CACHE_FILE,
    legacy_cache_file: Path | None = LEGACY_CACHE_FILE,
) -> dict[str, object]:
    """读取缓存；首次升级时兼容程序目录中的旧缓存文件。"""

    with _CACHE_LOCK:
        if cache_file.exists():
            return _read_json_object(cache_file)
        if legacy_cache_file is not None:
            return _read_json_object(legacy_cache_file)
        return {}


def update_cache_data(
    updates: dict[str, object],
    cache_file: Path = CACHE_FILE,
    legacy_cache_file: Path | None = LEGACY_CACHE_FILE,
) -> bool:
    """在同一把锁内合并并原子写入缓存，避免并发覆盖和半截 JSON。"""

    with _CACHE_LOCK:
        if cache_file.exists():
            payload = _read_json_object(cache_file)
        elif legacy_cache_file is not None:
            payload = _read_json_object(legacy_cache_file)
        else:
            payload = {}
        payload.update(updates)
        return _atomic_write_json(cache_file, payload)


def _read_json_object(file_path: Path) -> dict[str, object]:
    """读取 JSON 字典，文件不存在或损坏时返回空字典。"""

    try:
        raw_data = json.loads(file_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return raw_data if isinstance(raw_data, dict) else {}


def _atomic_write_json(file_path: Path, payload: dict[str, object]) -> bool:
    """先写同目录临时文件，再用原子替换提交完整 JSON。"""

    temporary_path: Path | None = None
    try:
        file_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=file_path.parent,
            prefix=f".{file_path.name}.",
            suffix=".tmp",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
            json.dump(payload, temporary_file, ensure_ascii=False, indent=2)
            temporary_file.write("\n")
            temporary_file.flush()
            os.fsync(temporary_file.fileno())
        os.replace(temporary_path, file_path)
        return True
    except OSError:
        return False
    finally:
        if temporary_path is not None and temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                pass
