"""无窗口偏好设置回归；所有缓存读写都隔离到临时文件或假读取器。"""

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app_preferences import (
    data_filter_config_to_dict,
    load_cache_data,
    load_data_filter_config,
    load_reference_config,
    normalize_folder_history,
    normalize_help_url,
    save_cache_data,
)
from cache_store import read_cache_data, update_cache_data
from cell_search import DEFAULT_DATA_FILTER_CONFIG
from worksheet_search import DEFAULT_REFERENCE_LOOKUP_CONFIG, WORKSHEET_SHEET_CACHE_KEY


class CachePreferenceTests(unittest.TestCase):
    def test_view_preference_round_trip_preserves_shared_sheet_index_cache(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            cache_file = Path(directory) / "cache.json"
            index_cache = {"existing.xlsx": {"sheets": ["技能"]}}
            update_cache_data({WORKSHEET_SHEET_CACHE_KEY: index_cache}, cache_file, None)
            preferences = {
                "pinned_folders": [r"D:\data\one", r"D:\data\one", r"D:\data\two"],
                "preview_enabled": False,
                "result_density": "紧凑",
                "file_result_view": "表格",
            }
            with patch("app_preferences.update_shared_cache_data", side_effect=lambda updates: update_cache_data(updates, cache_file, None)):
                self.assertTrue(save_cache_data(preferences))
            with patch("app_preferences.read_shared_cache_data", side_effect=lambda: read_cache_data(cache_file, None)):
                loaded = load_cache_data()

            self.assertEqual(loaded["pinned_folders"], [r"D:\data\one", r"D:\data\two"])
            self.assertFalse(loaded["preview_enabled"])
            self.assertEqual(loaded["result_density"], "紧凑")
            self.assertEqual(loaded["file_result_view"], "表格")
            self.assertEqual(read_cache_data(cache_file, None)[WORKSHEET_SHEET_CACHE_KEY], index_cache)

    def test_malformed_preference_values_are_normalized_without_reading_real_cache(self) -> None:
        raw = {
            "folder_path": 123,
            "keyword": ["wrong type"],
            "folder_history": [r" D:\data ", None, r"D:\data"],
            "file_suffixes": None,
            "reference_config": {"reference_row_index": "Infinity"},
            "data_filter_config": {"header_row_count": "NaN"},
        }
        with patch("app_preferences.read_shared_cache_data", return_value=raw):
            loaded = load_cache_data()
        self.assertEqual(loaded["folder_path"], "")
        self.assertEqual(loaded["keyword"], "")
        self.assertEqual(loaded["file_suffixes"], "")
        self.assertEqual(loaded["folder_history"], [r"D:\data"])
        self.assertIs(loaded["reference_config"], DEFAULT_REFERENCE_LOOKUP_CONFIG)
        self.assertIs(loaded["data_filter_config"], DEFAULT_DATA_FILTER_CONFIG)

    def test_old_row_filter_fields_migrate_to_independent_settings(self) -> None:
        disabled = load_data_filter_config({
            "row_filter_mode": "禁用", "row_marker_column_index": 2,
            "disabled_row_contains": "$",
        })
        self.assertTrue(disabled.enable_disabled_row_filter)
        self.assertFalse(disabled.enable_range_row_filter)
        self.assertEqual(disabled.disabled_row_marker_column_index, 2)

        ranged = load_data_filter_config({
            "row_filter_mode": "区间", "row_marker_column_index": 3,
            "range_start_text": "Start", "range_end_text": "End",
        })
        self.assertTrue(ranged.enable_range_row_filter)
        self.assertFalse(ranged.enable_disabled_row_filter)
        self.assertEqual(ranged.range_row_marker_column_index, 3)
        serialized = data_filter_config_to_dict(ranged)
        self.assertNotIn("row_filter_mode", serialized)
        self.assertTrue(serialized["enable_range_row_filter"])

    def test_explicit_new_row_filter_flags_override_legacy_mode(self) -> None:
        config = load_data_filter_config({
            "row_filter_mode": "区间", "enable_range_row_filter": False,
        })
        self.assertFalse(config.enable_range_row_filter)
        self.assertFalse(config.enable_disabled_row_filter)

    def test_reference_row_input_uses_shared_validation(self) -> None:
        config = load_reference_config({"reference_row_index": "第６行"})
        self.assertEqual(config.reference_row_index, 6)
        self.assertEqual(config.table_name, DEFAULT_REFERENCE_LOOKUP_CONFIG.table_name)

    def test_folder_history_and_help_links_keep_existing_normalization(self) -> None:
        history = normalize_folder_history([f"D:/folder/{index}" for index in range(30)], "D:/current")
        self.assertEqual(history[0], "D:/current")
        self.assertEqual(len(history), 20)
        self.assertEqual(normalize_help_url(" example.com/help "), "https://example.com/help")
        self.assertEqual(normalize_help_url("file:///D:/help.html"), "file:///D:/help.html")
        self.assertEqual(normalize_help_url(""), "")


if __name__ == "__main__":
    unittest.main()
