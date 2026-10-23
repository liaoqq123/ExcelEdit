"""Execute Excel navigation against managed fakes; never start Excel or COM.

Only the navigation body is executed. The running-workbook lookup is replaced by
a fake function, and the fake Application rejects application-wide visibility
changes. PowerShell itself runs with CREATE_NO_WINDOW.
"""

import base64
import json
import os
from pathlib import Path
import subprocess
import unittest


_FAKES = r"""
using System;
using System.Collections;
using System.Collections.Generic;

public sealed class NavigationFakeApp {
    public bool IsVisible = true;
    public int VisibilityWrites;
    public int GotoCalls;
    public int LastRow;
    public int LastColumn;
    public bool Scroll;
    public bool Visible {
        get { return IsVisible; }
        set {
            VisibilityWrites++;
            throw new InvalidOperationException("Application-wide visibility is prohibited");
        }
    }
    public void Goto(object cell, bool scroll) {
        NavigationFakeCell selected = (NavigationFakeCell)cell;
        GotoCalls++;
        LastRow = selected.Row;
        LastColumn = selected.Column;
        Scroll = scroll;
    }
}

public sealed class NavigationFakeWindow {
    public long Hwnd;
    public bool IsVisible;
    public int State;
    public int VisibilityWrites;
    public int StateWrites;
    public int Activations;
    public NavigationFakeWindow(long hwnd, bool visible, int state) {
        Hwnd = hwnd; IsVisible = visible; State = state;
    }
    public bool Visible {
        get { return IsVisible; }
        set { VisibilityWrites++; IsVisible = value; }
    }
    public int WindowState {
        get { return State; }
        set { StateWrites++; State = value; }
    }
    public void Activate() {
        Activations++;
        throw new InvalidOperationException("Background navigation must not activate a window");
    }
}

public sealed class NavigationFakeWindows : IEnumerable {
    public List<NavigationFakeWindow> Items = new List<NavigationFakeWindow>();
    public int Count { get { return Items.Count; } }
    public NavigationFakeWindow Item(int index) { return Items[index - 1]; }
    public IEnumerator GetEnumerator() { return Items.GetEnumerator(); }
}

public sealed class NavigationFakeCell {
    public int Row;
    public int Column;
    public NavigationFakeCell(int row, int column) { Row = row; Column = column; }
}

public sealed class NavigationFakeCells {
    public NavigationFixture Fixture;
    public bool MinimizeOnItemLookup;
    public NavigationFakeCell Item(int row, int column) {
        if (MinimizeOnItemLookup) {
            Fixture.Target.State = -4140;
            Fixture.ForegroundHandle = 0;
        }
        return new NavigationFakeCell(row, column);
    }
}

public sealed class NavigationFakeSheet {
    public NavigationFixture Fixture;
    public bool MinimizeOnActivate;
    public int Activations;
    public NavigationFakeCells Cells = new NavigationFakeCells();
    public void Activate() {
        Activations++;
        if (MinimizeOnActivate) {
            Fixture.Target.State = -4140;
            Fixture.ForegroundHandle = 0;
        }
    }
}

public sealed class NavigationFakeSheets {
    public NavigationFixture Fixture;
    public bool SwitchOnItemLookup;
    public NavigationFakeSheet Data = new NavigationFakeSheet();
    public NavigationFakeSheet Item(string name) {
        if (name != "Data") throw new InvalidOperationException("Worksheet does not exist");
        if (SwitchOnItemLookup) Fixture.ForegroundHandle = 9999;
        return Data;
    }
}

public sealed class NavigationFakeBook {
    public string FullName = @"C:\fixture\Book.xlsx";
    public NavigationFakeApp Application;
    public NavigationFakeWindows Windows = new NavigationFakeWindows();
    public NavigationFakeSheets Worksheets = new NavigationFakeSheets();
    public int Activations;
    public void Activate() {
        if (!Application.IsVisible) throw new InvalidOperationException("Application remains hidden");
        if (Windows.Count == 0 || !Windows.Item(1).Visible)
            throw new InvalidOperationException("Workbook activation selects its hidden first view");
        Activations++;
    }
}

public sealed class NavigationFixture {
    public NavigationFakeApp App = new NavigationFakeApp();
    public NavigationFakeWindow Target = new NavigationFakeWindow(4321, true, -4137);
    public NavigationFakeWindow HiddenView = new NavigationFakeWindow(5432, false, -4143);
    public NavigationFakeWindow Other = new NavigationFakeWindow(9876, false, -4143);
    public NavigationFakeBook Book = new NavigationFakeBook();
    public bool Found = true;
    public int Lookups;
    public int AvailableAfterLookups = 1;
    public int MinimizeOnLookupNumber;
    public long ForegroundHandle = 4321;
    public int ForegroundChecks;
    public NavigationFixture() {
        Book.Application = App;
        Book.Windows.Items.Add(Target);
        Book.Worksheets.Fixture = this;
        Book.Worksheets.Data.Fixture = this;
        Book.Worksheets.Data.Cells.Fixture = this;
    }
}
"""


@unittest.skipUnless(os.name == "nt", "Windows PowerShell is required for fake-script execution")
class ExcelNavigationScriptTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import desktop_open

        navigation = desktop_open._EXCEL_NAVIGATION_SCRIPT
        executable = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"
        if not executable.is_file():
            raise unittest.SkipTest("Windows PowerShell is unavailable")
        scenarios = [
            {"name": "workbook", "sheet": "", "row": None, "column": None},
            {"name": "cell", "sheet": "Data", "row": 1093, "column": 3},
            {"name": "missing_book", "sheet": "Data", "row": 7, "column": 3},
            {"name": "missing_sheet", "sheet": "Missing", "row": 7, "column": 3},
            {"name": "hidden_app", "sheet": "Data", "row": 7, "column": 3},
            {"name": "visible_second_view", "sheet": "Data", "row": 7, "column": 3},
            {"name": "wrong_fullname", "sheet": "Data", "row": 7, "column": 3},
            {"name": "only_hidden_view", "sheet": "Data", "row": 7, "column": 3},
            {"name": "minimized", "sheet": "Data", "row": 7, "column": 3},
            {"name": "switched_app", "sheet": "Data", "row": 7, "column": 3},
            {"name": "minimized_waiting_for_connection", "sheet": "Data", "row": 7, "column": 3},
            {"name": "switched_during_sheet_lookup", "sheet": "Data", "row": 7, "column": 3},
            {"name": "minimized_during_navigation", "sheet": "Data", "row": 7, "column": 3},
            {"name": "minimized_during_cell_lookup", "sheet": "Data", "row": 7, "column": 3},
            {"name": "foreground_second_view", "sheet": "Data", "row": 7, "column": 3},
        ]
        source = (
            "$ErrorActionPreference='Stop'\nAdd-Type -TypeDefinition @'\n"
            + _FAKES + "\n'@\n"
            + "$navigation = {\n" + navigation + "\n}\n"
            + "$scenarios = '" + json.dumps(scenarios, ensure_ascii=True) + "' | ConvertFrom-Json\n"
            + r"""
function Find-OpenedWorkbook($filePath) {
    $script:fixture.Lookups++
    if ($script:fixture.Lookups -eq $script:fixture.MinimizeOnLookupNumber) {
        $script:fixture.Target.State = -4140
        $script:fixture.ForegroundHandle = 0
    }
    if ($script:fixture.Lookups -lt $script:fixture.AvailableAfterLookups) { return $null }
    if ($script:fixture.Found) { return $script:fixture.Book }
    return $null
}
function Test-TargetForeground([long]$hwnd) {
    $script:fixture.ForegroundChecks++
    foreach ($view in $script:fixture.Book.Windows.Items) {
        if ($view.Hwnd -eq $hwnd) {
            return ($script:fixture.ForegroundHandle -eq $hwnd -and $view.State -ne -4140)
        }
    }
    return $false
}
$results = foreach ($scenario in $scenarios) {
    $script:fixture = [NavigationFixture]::new()
    if ($scenario.name -eq 'missing_book') { $script:fixture.Found = $false }
    if ($scenario.name -eq 'hidden_app') { $script:fixture.App.IsVisible = $false }
    if ($scenario.name -eq 'visible_second_view') {
        $script:fixture.Book.Windows.Items.Insert(0, $script:fixture.HiddenView)
    }
    if ($scenario.name -eq 'wrong_fullname') {
        $script:fixture.Book.FullName = 'C:\other\Book.xlsx'
    }
    if ($scenario.name -eq 'only_hidden_view') {
        $script:fixture.Target.IsVisible = $false
        $script:fixture.Other.IsVisible = $true
    }
    if ($scenario.name -eq 'minimized') { $script:fixture.Target.State = -4140 }
    if ($scenario.name -eq 'switched_app') { $script:fixture.ForegroundHandle = 9999 }
    if ($scenario.name -eq 'minimized_waiting_for_connection') {
        $script:fixture.AvailableAfterLookups = 2
        $script:fixture.MinimizeOnLookupNumber = 2
    }
    if ($scenario.name -eq 'switched_during_sheet_lookup') {
        $script:fixture.Book.Worksheets.SwitchOnItemLookup = $true
    }
    if ($scenario.name -eq 'minimized_during_navigation') {
        $script:fixture.Book.Worksheets.Data.MinimizeOnActivate = $true
    }
    if ($scenario.name -eq 'minimized_during_cell_lookup') {
        $script:fixture.Book.Worksheets.Data.Cells.MinimizeOnItemLookup = $true
    }
    if ($scenario.name -eq 'foreground_second_view') {
        $script:fixture.Other.IsVisible = $true
        $script:fixture.Book.Windows.Items.Add($script:fixture.Other)
        $script:fixture.ForegroundHandle = $script:fixture.Other.Hwnd
    }
    $request = [pscustomobject]@{
        path='C:\fixture\Book.xlsx'; sheet=$scenario.sheet;
        row=$scenario.row; column=$scenario.column; wait=0
    }
    if ($scenario.name -eq 'minimized_waiting_for_connection') { $request.wait = 1 }
    $result = & $navigation | ConvertFrom-Json
    [pscustomobject]@{
        name=$scenario.name; result=$result;
        appVisibleWrites=$script:fixture.App.VisibilityWrites;
        appVisible=$script:fixture.App.IsVisible;
        gotoCalls=$script:fixture.App.GotoCalls;
        row=$script:fixture.App.LastRow; column=$script:fixture.App.LastColumn;
        scroll=$script:fixture.App.Scroll;
        targetVisible=$script:fixture.Target.IsVisible;
        targetVisibleWrites=$script:fixture.Target.VisibilityWrites;
        targetState=$script:fixture.Target.State;
        targetStateWrites=$script:fixture.Target.StateWrites;
        targetActivations=$script:fixture.Target.Activations;
        hiddenViewVisible=$script:fixture.HiddenView.IsVisible;
        hiddenViewVisibleWrites=$script:fixture.HiddenView.VisibilityWrites;
        otherVisible=$script:fixture.Other.IsVisible;
        otherVisibleWrites=$script:fixture.Other.VisibilityWrites;
        otherActivations=$script:fixture.Other.Activations;
        otherStateWrites=$script:fixture.Other.StateWrites;
        bookActivations=$script:fixture.Book.Activations;
        sheetActivations=$script:fixture.Book.Worksheets.Data.Activations;
        lookups=$script:fixture.Lookups; foregroundChecks=$script:fixture.ForegroundChecks
    }
}
ConvertTo-Json -InputObject @($results) -Depth 8 -Compress
"""
        )
        bootstrap = (
            "[Console]::InputEncoding=[Text.UTF8Encoding]::new($false);"
            "[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false);"
            "& ([ScriptBlock]::Create([Console]::In.ReadToEnd()))"
        )
        command = [
            str(executable), "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand",
            base64.b64encode(bootstrap.encode("utf-16-le")).decode("ascii"),
        ]
        completed = subprocess.run(
            command, input=source, capture_output=True, text=True, encoding="utf-8",
            timeout=25, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if completed.returncode:
            raise AssertionError(f"Fake navigation harness failed: {completed.stderr}")
        cls.scenarios = {item["name"]: item for item in json.loads(completed.stdout)}

    def test_target_workbook_is_visible_without_revealing_other_windows(self):
        item = self.scenarios["workbook"]
        self.assertTrue(item["result"]["located"])
        self.assertEqual(item["result"]["hwnd"], 4321)
        self.assertTrue(item["targetVisible"])
        self.assertEqual(item["appVisibleWrites"], 0)
        self.assertEqual(item["otherVisibleWrites"], 0)
        self.assertFalse(item["otherVisible"])
        self.assertEqual(item["targetState"], -4137)
        self.assertEqual(item["targetStateWrites"], 0)

    def test_cell_navigation_activates_sheet_and_scrolls_to_requested_cell(self):
        item = self.scenarios["cell"]
        self.assertTrue(item["result"]["located"])
        self.assertEqual(item["sheetActivations"], 1)
        self.assertEqual(item["gotoCalls"], 1)
        self.assertEqual((item["row"], item["column"]), (1093, 3))
        self.assertTrue(item["scroll"])
        self.assertEqual(item["appVisibleWrites"], 0)

    def test_missing_workbook_does_not_claim_navigation_or_activate_any_window(self):
        item = self.scenarios["missing_book"]
        self.assertTrue(item["result"]["opened"])
        self.assertFalse(item["result"]["located"])
        self.assertEqual(item["result"]["hwnd"], 0)
        self.assertTrue(item["result"]["error"])
        self.assertEqual(item["targetActivations"], 0)
        self.assertEqual(item["bookActivations"], 0)
        self.assertGreaterEqual(item["lookups"], 1)

    def test_missing_sheet_preserves_workbook_handle_and_reports_partial_result(self):
        item = self.scenarios["missing_sheet"]
        self.assertTrue(item["result"]["opened"])
        self.assertFalse(item["result"]["located"])
        self.assertEqual(item["result"]["hwnd"], 4321)
        self.assertTrue(item["result"]["error"])
        self.assertTrue(item["targetVisible"])
        self.assertEqual(item["gotoCalls"], 0)

    def test_hidden_application_is_not_treated_as_visible_success(self):
        item = self.scenarios["hidden_app"]
        self.assertTrue(item["result"]["opened"])
        self.assertFalse(item["result"]["located"])
        self.assertTrue(item["result"]["error"])
        self.assertFalse(item["appVisible"])
        self.assertEqual(item["appVisibleWrites"], 0)
        self.assertEqual(item["targetActivations"], 0)
        self.assertEqual(item["bookActivations"], 0)

    def test_existing_visible_view_is_used_without_revealing_hidden_view(self):
        item = self.scenarios["visible_second_view"]
        self.assertTrue(item["result"]["located"])
        self.assertEqual(item["result"]["hwnd"], 4321)
        self.assertFalse(item["hiddenViewVisible"])
        self.assertEqual(item["hiddenViewVisibleWrites"], 0)
        self.assertEqual(item["bookActivations"], 0)
        self.assertEqual(item["targetActivations"], 0)

    def test_wrong_fullname_candidate_is_not_activated(self):
        item = self.scenarios["wrong_fullname"]
        self.assertTrue(item["result"]["opened"])
        self.assertFalse(item["result"]["located"])
        self.assertEqual(item["result"]["hwnd"], 0)
        self.assertTrue(item["result"]["error"])
        self.assertEqual(item["targetActivations"], 0)
        self.assertEqual(item["bookActivations"], 0)
        self.assertEqual(item["sheetActivations"], 0)
        self.assertEqual(item["appVisibleWrites"], 0)

    def test_workbook_with_only_hidden_view_is_not_revealed_or_activated(self):
        item = self.scenarios["only_hidden_view"]
        self.assertTrue(item["result"]["opened"])
        self.assertFalse(item["result"]["located"])
        self.assertEqual(item["result"]["hwnd"], 0)
        self.assertTrue(item["result"]["error"])
        self.assertFalse(item["targetVisible"])
        self.assertEqual(item["targetVisibleWrites"], 0)
        self.assertEqual(item["targetActivations"], 0)
        self.assertEqual(item["bookActivations"], 0)
        self.assertTrue(item["otherVisible"])
        self.assertEqual(item["otherVisibleWrites"], 0)
        self.assertEqual(item["appVisibleWrites"], 0)

    def assert_navigation_stopped_without_window_changes(self, item, *, sheet_activations=0):
        self.assertTrue(item["result"]["opened"])
        self.assertFalse(item["result"]["located"])
        self.assertEqual(item["result"]["hwnd"], 4321)
        self.assertTrue(item["result"]["error"])
        self.assertEqual(item["targetStateWrites"], 0)
        self.assertEqual(item["targetActivations"], 0)
        self.assertEqual(item["bookActivations"], 0)
        self.assertEqual(item["sheetActivations"], sheet_activations)
        self.assertEqual(item["gotoCalls"], 0)
        self.assertEqual(item["targetVisibleWrites"], 0)
        self.assertEqual(item["otherVisibleWrites"], 0)
        self.assertEqual(item["appVisibleWrites"], 0)

    def test_minimized_target_is_not_restored_or_activated(self):
        item = self.scenarios["minimized"]
        self.assert_navigation_stopped_without_window_changes(item)
        self.assertEqual(item["targetState"], -4140)

    def test_switch_to_other_application_cancels_navigation_without_retry(self):
        item = self.scenarios["switched_app"]
        self.assert_navigation_stopped_without_window_changes(item)
        self.assertEqual(item["lookups"], 1)

    def test_minimized_while_waiting_for_workbook_connection_is_not_reopened(self):
        item = self.scenarios["minimized_waiting_for_connection"]
        self.assert_navigation_stopped_without_window_changes(item)
        self.assertEqual(item["targetState"], -4140)
        self.assertEqual(item["lookups"], 2)

    def test_switch_during_sheet_lookup_stops_before_sheet_activation(self):
        item = self.scenarios["switched_during_sheet_lookup"]
        self.assert_navigation_stopped_without_window_changes(item)

    def test_minimize_during_sheet_activation_stops_before_cell_navigation(self):
        item = self.scenarios["minimized_during_navigation"]
        self.assert_navigation_stopped_without_window_changes(item, sheet_activations=1)
        self.assertEqual(item["targetState"], -4140)

    def test_minimize_during_cell_lookup_stops_before_goto(self):
        item = self.scenarios["minimized_during_cell_lookup"]
        self.assert_navigation_stopped_without_window_changes(item, sheet_activations=1)
        self.assertEqual(item["targetState"], -4140)

    def test_current_foreground_view_is_preferred_to_first_visible_view(self):
        item = self.scenarios["foreground_second_view"]
        self.assertTrue(item["result"]["located"])
        self.assertEqual(item["result"]["hwnd"], 9876)
        self.assertEqual(item["targetActivations"], 0)
        self.assertEqual(item["otherActivations"], 0)
        self.assertEqual(item["bookActivations"], 0)
        self.assertEqual(item["targetStateWrites"], 0)
        self.assertEqual(item["otherStateWrites"], 0)
        self.assertEqual(item["targetVisibleWrites"], 0)
        self.assertEqual(item["otherVisibleWrites"], 0)
        self.assertEqual(item["appVisibleWrites"], 0)
        self.assertEqual(item["gotoCalls"], 1)


if __name__ == "__main__":
    unittest.main()
