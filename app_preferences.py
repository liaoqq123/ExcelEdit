"""应用偏好的规范化、旧配置迁移与保存；不依赖窗口组件。"""

import os

from cache_store import read_cache_data as read_shared_cache_data
from cache_store import update_cache_data as update_shared_cache_data
from cell_search import DEFAULT_DATA_FILTER_CONFIG, DataFilterConfig, validate_data_filter_config
from excel_common import EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS, ExcelReadError
from input_validation import parse_positive_int
from worksheet_search import (
    DEFAULT_REFERENCE_LOOKUP_CONFIG,
    WORKSHEET_SHEET_CACHE_KEY,
    ReferenceLookupConfig,
    validate_reference_lookup_config,
)


DEFAULT_DISABLED_SHEET_MARKER = "$"
FOLDER_HISTORY_LIMIT = 20


def load_cache_data() -> dict[str, object]:
    """读取并规范化缓存数据，避免坏缓存影响启动。"""

    raw_data = read_shared_cache_data()
    if not raw_data:
        return {}

    folder_path = raw_data.get("folder_path", "")
    keyword = raw_data.get("keyword", "")
    disabled_sheet_marker = raw_data.get("disabled_sheet_marker", DEFAULT_DISABLED_SHEET_MARKER)
    help_url = raw_data.get("help_url", "")
    background_image_path = raw_data.get("background_image_path", "")
    file_suffixes = raw_data.get("file_suffixes", "")
    return {
        "pinned_folders": normalize_folder_history(raw_data.get("pinned_folders"))[:10],
        "preview_enabled": raw_data.get("preview_enabled", True) is not False,
        "result_density": "紧凑" if raw_data.get("result_density") == "紧凑" else "舒适",
        "file_result_view": "表格" if raw_data.get("file_result_view") == "表格" else "缩略图",
        "folder_path": folder_path if isinstance(folder_path, str) else "",
        "folder_history": normalize_folder_history(
            raw_data.get("folder_history"),
            folder_path if isinstance(folder_path, str) else "",
        ),
        "keyword": keyword if isinstance(keyword, str) else "",
        "reference_config": load_reference_config(raw_data.get("reference_config")),
        "disabled_sheet_marker": disabled_sheet_marker if isinstance(disabled_sheet_marker, str) else "",
        "data_filter_config": load_data_filter_config(raw_data.get("data_filter_config")),
        "help_url": help_url if isinstance(help_url, str) else "",
        "background_image_path": background_image_path if isinstance(background_image_path, str) else "",
        "only_enabled_sheets": bool(raw_data.get("only_enabled_sheets", False)),
        "only_enabled_data": bool(raw_data.get("only_enabled_data", False)),
        "only_specified_suffixes": bool(raw_data.get("only_specified_suffixes", False)),
        "file_suffixes": file_suffixes if isinstance(file_suffixes, str) else "",
        "search_folders": bool(raw_data.get("search_folders", False)),
        "exact_match": bool(raw_data.get("exact_match", False)),
    }


def save_cache_data(cache_data: dict[str, object]) -> bool:
    """保存缓存数据，同时保留工作表名称缓存等其他键。"""

    reference_config = cache_data.get("reference_config", DEFAULT_REFERENCE_LOOKUP_CONFIG)
    if not isinstance(reference_config, ReferenceLookupConfig):
        reference_config = DEFAULT_REFERENCE_LOOKUP_CONFIG
    data_filter_config = cache_data.get("data_filter_config", DEFAULT_DATA_FILTER_CONFIG)
    if not isinstance(data_filter_config, DataFilterConfig):
        data_filter_config = DEFAULT_DATA_FILTER_CONFIG

    updates: dict[str, object] = {
        "pinned_folders": normalize_folder_history(cache_data.get("pinned_folders"))[:10],
        "preview_enabled": cache_data.get("preview_enabled", True) is not False,
        "result_density": "紧凑" if cache_data.get("result_density") == "紧凑" else "舒适",
        "file_result_view": "表格" if cache_data.get("file_result_view") == "表格" else "缩略图",
        "folder_path": str(cache_data.get("folder_path", "")),
        "folder_history": normalize_folder_history(cache_data.get("folder_history")),
        "keyword": str(cache_data.get("keyword", "")),
        "reference_config": reference_config_to_dict(reference_config),
        "disabled_sheet_marker": str(cache_data.get("disabled_sheet_marker", DEFAULT_DISABLED_SHEET_MARKER)),
        "data_filter_config": data_filter_config_to_dict(data_filter_config),
        "help_url": str(cache_data.get("help_url", "")),
        "background_image_path": str(cache_data.get("background_image_path", "")),
        "only_enabled_sheets": bool(cache_data.get("only_enabled_sheets", False)),
        "only_enabled_data": bool(cache_data.get("only_enabled_data", False)),
        "only_specified_suffixes": bool(cache_data.get("only_specified_suffixes", False)),
        "file_suffixes": str(cache_data.get("file_suffixes", "")),
        "search_folders": bool(cache_data.get("search_folders", False)),
        "exact_match": bool(cache_data.get("exact_match", False)),
    }
    worksheet_sheet_cache = cache_data.get(WORKSHEET_SHEET_CACHE_KEY)
    if isinstance(worksheet_sheet_cache, dict):
        updates[WORKSHEET_SHEET_CACHE_KEY] = worksheet_sheet_cache
    return update_shared_cache_data(updates)


def normalize_folder_history(raw_history: object, current_folder: str = "") -> list[str]:
    """清理文件夹历史记录：去空、去重、保留最近使用项。"""

    candidates: list[str] = []
    if current_folder:
        candidates.append(current_folder)
    if isinstance(raw_history, list):
        candidates.extend(item for item in raw_history if isinstance(item, str))

    history: list[str] = []
    seen_keys: set[str] = set()
    for candidate in candidates:
        folder_path = candidate.strip()
        if not folder_path:
            continue
        folder_key = os.path.normcase(os.path.normpath(folder_path))
        if folder_key in seen_keys:
            continue
        seen_keys.add(folder_key)
        history.append(folder_path)
        if len(history) >= FOLDER_HISTORY_LIMIT:
            break
    return history


def load_reference_config(raw_config: object) -> ReferenceLookupConfig:
    """从缓存恢复引用查找配置，非法时回到默认值。"""

    if not isinstance(raw_config, dict):
        return DEFAULT_REFERENCE_LOOKUP_CONFIG

    try:
        config = ReferenceLookupConfig(
            sample_text=str(raw_config.get("sample_text", DEFAULT_REFERENCE_LOOKUP_CONFIG.sample_text)),
            reference_row_index=parse_positive_int(
                raw_config.get("reference_row_index", DEFAULT_REFERENCE_LOOKUP_CONFIG.reference_row_index),
                "引用行",
                EXCEL_MAX_ROWS,
            ),
            table_name=str(raw_config.get("table_name", DEFAULT_REFERENCE_LOOKUP_CONFIG.table_name)),
            field_name=str(raw_config.get("field_name", DEFAULT_REFERENCE_LOOKUP_CONFIG.field_name)),
            field_row_index=parse_positive_int(
                raw_config.get("field_row_index", DEFAULT_REFERENCE_LOOKUP_CONFIG.field_row_index),
                "引用字段行",
                EXCEL_MAX_ROWS,
            ),
        )
        validate_reference_lookup_config(config)
    except ExcelReadError:
        return DEFAULT_REFERENCE_LOOKUP_CONFIG
    return config


def load_data_filter_config(raw_config: object) -> DataFilterConfig:
    """从缓存恢复启用数据过滤配置，并兼容旧版行过滤字段。"""

    if not isinstance(raw_config, dict):
        return DEFAULT_DATA_FILTER_CONFIG

    try:
        # 旧版本只有 row_filter_mode，这里把它迁移到禁用行/区间行两个独立开关。
        old_row_filter_mode = str(raw_config.get("row_filter_mode", ""))
        old_row_marker_column_index = raw_config.get(
            "row_marker_column_index",
            DEFAULT_DATA_FILTER_CONFIG.disabled_row_marker_column_index,
        )
        config = DataFilterConfig(
            enable_header_filter=bool(raw_config.get("enable_header_filter", False)),
            header_row_count=parse_positive_int(
                raw_config.get(
                    "header_row_count",
                    DEFAULT_DATA_FILTER_CONFIG.header_row_count,
                ),
                "表头长度",
                EXCEL_MAX_ROWS,
            ),
            enable_blank_header_column_filter=bool(
                raw_config.get("enable_blank_header_column_filter", False)
            ),
            enable_column_filter=bool(raw_config.get("enable_column_filter", False)),
            column_marker_row_index=parse_positive_int(
                raw_config.get(
                    "column_marker_row_index",
                    DEFAULT_DATA_FILTER_CONFIG.column_marker_row_index,
                ),
                "列判断行",
                EXCEL_MAX_ROWS,
            ),
            disabled_column_marker=str(
                raw_config.get(
                    "disabled_column_marker",
                    DEFAULT_DATA_FILTER_CONFIG.disabled_column_marker,
                )
            ),
            enable_disabled_row_filter=bool(
                raw_config.get(
                    "enable_disabled_row_filter",
                    old_row_filter_mode == "禁用",
                )
            ),
            disabled_row_marker_column_index=parse_positive_int(
                raw_config.get(
                    "disabled_row_marker_column_index",
                    old_row_marker_column_index,
                ),
                "行禁用判断列",
                EXCEL_MAX_COLUMNS,
            ),
            disabled_row_contains=str(raw_config.get("disabled_row_contains", "")),
            enable_range_row_filter=bool(
                raw_config.get(
                    "enable_range_row_filter",
                    old_row_filter_mode == "区间",
                )
            ),
            range_row_marker_column_index=parse_positive_int(
                raw_config.get(
                    "range_row_marker_column_index",
                    old_row_marker_column_index,
                ),
                "行区间判断列",
                EXCEL_MAX_COLUMNS,
            ),
            range_start_text=str(raw_config.get("range_start_text", "")),
            range_end_text=str(raw_config.get("range_end_text", "")),
        )
        validate_data_filter_config(config)
    except ExcelReadError:
        return DEFAULT_DATA_FILTER_CONFIG
    return config


def reference_config_to_dict(config: ReferenceLookupConfig) -> dict[str, object]:
    """把引用配置转换成可写入 JSON 的字典。"""

    return {
        "sample_text": config.sample_text,
        "reference_row_index": config.reference_row_index,
        "table_name": config.table_name,
        "field_name": config.field_name,
        "field_row_index": config.field_row_index,
    }


def data_filter_config_to_dict(config: DataFilterConfig) -> dict[str, object]:
    """把启用数据过滤配置转换成可写入 JSON 的字典。"""

    return {
        "enable_header_filter": config.enable_header_filter,
        "header_row_count": config.header_row_count,
        "enable_blank_header_column_filter": config.enable_blank_header_column_filter,
        "enable_column_filter": config.enable_column_filter,
        "column_marker_row_index": config.column_marker_row_index,
        "disabled_column_marker": config.disabled_column_marker,
        "enable_disabled_row_filter": config.enable_disabled_row_filter,
        "disabled_row_marker_column_index": config.disabled_row_marker_column_index,
        "disabled_row_contains": config.disabled_row_contains,
        "enable_range_row_filter": config.enable_range_row_filter,
        "range_row_marker_column_index": config.range_row_marker_column_index,
        "range_start_text": config.range_start_text,
        "range_end_text": config.range_end_text,
    }


def normalize_help_url(help_url: str) -> str:
    """补全帮助链接协议，方便用户只输入域名。"""

    url = help_url.strip()
    if not url:
        return ""
    known_schemes = ("http://", "https://", "file://", "mailto:")
    if url.lower().startswith(known_schemes):
        return url
    return f"https://{url}"
