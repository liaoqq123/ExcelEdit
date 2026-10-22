"""验证扫描覆盖范围、限流、错误明细及取消状态。"""

import os
from pathlib import Path
import tempfile
from threading import Event
import unittest
from unittest.mock import patch

import xlsxwriter

from cell_search import search_excel_cells
from file_search import search_files
from json_validation import validate_excel_json_columns
from search_status import SearchTracker


class SearchStatusTests(unittest.TestCase):
    def test_file_limit_and_completed_are_distinct(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            for index in range(4):
                (Path(directory) / f"{index}.txt").touch()
            for limit, expected in ((3, "limit"), (10, "completed")):
                snapshots = []
                tracker = SearchTracker(snapshots.append)
                results = search_files(directory, max_results=limit, tracker=tracker)
                self.assertEqual(tracker.stop_reason, expected)
                self.assertEqual(tracker.result_count, len(results))
                self.assertEqual(snapshots[-1].stop_reason, expected)
                self.assertEqual(snapshots[-1].scanned_count, len(results))

    def test_cancelled_scan_emits_final_state(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            for index in range(20):
                (Path(directory) / f"{index}.txt").touch()
            cancel = Event()
            snapshots = []

            def progress(snapshot):
                snapshots.append(snapshot)
                if snapshot.scanned_count:
                    cancel.set()

            tracker = SearchTracker(progress)
            search_files(directory, cancel_event=cancel, tracker=tracker)
            self.assertEqual(tracker.stop_reason, "cancelled")
            self.assertLess(tracker.scanned_count, 20)
            self.assertEqual(snapshots[-1].stop_reason, "cancelled")

    def test_progress_is_throttled_and_final_snapshot_is_forced(self) -> None:
        snapshots = []
        tracker = SearchTracker(snapshots.append)
        with patch("search_status.monotonic", return_value=tracker.started):
            for index in range(1000):
                tracker.scan(Path(str(index)))
            tracker.finish()
        self.assertEqual(len(snapshots), 2)
        self.assertEqual(snapshots[-1].scanned_count, 1000)

    def test_unreadable_directory_is_reported_without_hiding_other_files(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            blocked = root / "blocked"
            blocked.mkdir()
            (root / "good.txt").touch()
            scandir = os.scandir

            def scan(path):
                if Path(path) == blocked:
                    raise PermissionError("test access denied")
                return scandir(path)

            tracker = SearchTracker()
            with patch("file_search.os.scandir", side_effect=scan):
                results = search_files(root, tracker=tracker)
            self.assertEqual([result.file_name for result in results], ["good.txt"])
            self.assertEqual(tracker.issues[0].path, blocked)
            self.assertIn("access denied", tracker.issues[0].message)

    def test_cell_limit_and_bad_workbook_retain_error_cause(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "00-bad.xlsx").write_bytes(b"invalid zip")
            workbook = xlsxwriter.Workbook(root / "good.xlsx")
            sheet = workbook.add_worksheet("Data")
            for row in range(5):
                sheet.write(row, 0, "needle")
            workbook.close()
            tracker = SearchTracker()
            matches = search_excel_cells(root, "needle", max_results=2, tracker=tracker)
            self.assertEqual(len(matches), 2)
            self.assertEqual(tracker.stop_reason, "limit")
            self.assertIn("zip", tracker.issues[0].message.lower())

    def test_json_limit_reports_incomplete_scan(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = xlsxwriter.Workbook(Path(directory) / "data.xlsx")
            sheet = workbook.add_worksheet("Data")
            sheet.write("A4", "{}")
            sheet.write("A6", "[]")
            sheet.write("A7", "{bad}")
            sheet.write("A8", "{also bad}")
            workbook.close()
            tracker = SearchTracker()
            report = validate_excel_json_columns(directory, max_errors=1, tracker=tracker)
            self.assertEqual(len(report.errors), 1)
            self.assertTrue(report.limit_reached)
            self.assertEqual(tracker.stop_reason, "limit")


if __name__ == "__main__":
    unittest.main()
