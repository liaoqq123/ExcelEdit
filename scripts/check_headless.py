"""Run the explicit test suite that does not create application windows.

Run from any directory with this script's absolute path. Tests that require Tk
are deliberately absent; they are never discovered or started. The Excel script
regressions use managed fakes inside a hidden PowerShell subprocess.
"""

from __future__ import annotations

import argparse
from contextlib import ExitStack
from pathlib import Path
import sys
import unittest
from unittest.mock import patch


SAFE_TESTS = (
    "test_core",
    "test_excel_file_scope",
    "test_search_status",
    "test_result_export",
    "test_preview_image_lifecycle",
    "test_gui_performance",
    "test_result_preview.PreviewReaderTests",
    "test_windows_shell",
    "test_window_activation",
    "test_desktop_open",
    "test_result_actions",
    "test_excel_navigation",
    "test_app_preferences",
    "test_result_filter",
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run only tests that do not create application windows.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Show each test name.")
    parser.add_argument("--list", action="store_true", help="List the selected modules/classes without importing tests.")
    args = parser.parse_args(argv)

    if args.list:
        print("\n".join(SAFE_TESTS))
        return 0

    repository = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(repository))
    sys.path.insert(0, str(repository / "tests"))

    # These guards must be active during imports as well as test execution.
    # An accidental UI/native call fails the check instead of opening a window.
    with ExitStack() as guards:
        guards.enter_context(patch("_tkinter.create", side_effect=AssertionError("Headless checks must not create Tk windows")))
        guards.enter_context(patch("ctypes.WinDLL", side_effect=AssertionError("Headless checks must not call native window APIs"), create=True))
        guards.enter_context(patch("os.startfile", side_effect=AssertionError("Headless checks must not open external applications"), create=True))
        loader = unittest.TestLoader()
        suite = unittest.TestSuite(loader.loadTestsFromName(name) for name in SAFE_TESTS)
        result = unittest.TextTestRunner(verbosity=2 if args.verbose else 1).run(suite)

    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
