import json
import os
import subprocess
import sys
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from threading import Event, Thread
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from cell_search import (
    DEFAULT_DATA_FILTER_CONFIG,
    DataFilterConfig,
    ExcelCellMatch,
    search_excel_cells,
    validate_data_filter_config,
)
from excel_common import ExcelReadError
from file_search import FileSearchResult, search_files
from popup_gui import SettingsDialog, parse_positive_int
from worksheet_search import (
    DEFAULT_REFERENCE_LOOKUP_CONFIG,
    ExcelSheetReference,
    ExcelWorkbookInfo,
    ReferenceLookupConfig,
    WORKSHEET_SHEET_CACHE_KEY,
    find_sheet_reference_matches,
    scan_excel_workbooks,
    validate_reference_lookup_config,
)


def get_app_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CACHE_FILE = get_app_directory() / "cache_data.json"
DEFAULT_DISABLED_SHEET_MARKER = "$"
WORKSHEET_RESULT_COLUMNS = (
    ("workbook", "工作簿", 360, 220, True),
    ("sheet", "工作表", 360, 220, True),
    ("status", "启用状态", 150, 110, False),
    ("action", "操作", 132, 120, False),
)
CELL_RESULT_COLUMNS = (
    ("workbook", "工作簿", 240, 160, True),
    ("sheet", "工作表", 220, 140, True),
    ("row", "行", 70, 60, False),
    ("column", "列", 70, 60, False),
    ("value", "数据内容", 500, 260, True),
)
FILE_RESULT_COLUMNS = (
    ("file_name", "文件名", 320, 180, True),
    ("file_address", "文件地址", 760, 360, True),
)


@dataclass(frozen=True)
class ResultTarget:
    file_path: Path
    sheet_name: str | None = None
    row_index: int | None = None
    column_index: int | None = None
    open_containing_folder: bool = False


class ExcelEditApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Excel 工作簿检索")
        self.geometry("1080x680")
        self.minsize(840, 520)

        self.item_targets: dict[str, ResultTarget] = {}
        self.reference_items_by_source: dict[str, list[str]] = {}
        self.reference_searching_items: set[str] = set()
        self.action_buttons: dict[str, ctk.CTkButton] = {}
        self.action_layout_job: str | None = None
        self.current_result_columns: tuple[str, ...] = ()
        self.search_mode_var = ctk.StringVar(value="工作表检索")
        self.only_enabled_sheets_var = ctk.BooleanVar(value=False)
        self.only_enabled_data_var = ctk.BooleanVar(value=False)
        self.current_keyword = ""
        self.reference_config = DEFAULT_REFERENCE_LOOKUP_CONFIG
        self.disabled_sheet_marker = DEFAULT_DISABLED_SHEET_MARKER
        self.data_filter_config = DEFAULT_DATA_FILTER_CONFIG
        self.help_url = ""
        self.settings_dialog = SettingsDialog(self, DEFAULT_DISABLED_SHEET_MARKER, self._apply_settings)
        self.settings_button: ctk.CTkButton | None = None
        self.help_button: ctk.CTkButton | None = None
        self.cancel_search_button: ctk.CTkButton | None = None
        self.search_cancel_event: Event | None = None

        self._build_layout()
        self._load_cached_inputs()
        self.protocol("WM_DELETE_WINDOW", self.close_app)

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=16, pady=(16, 4), sticky="ew")
        top_frame.grid_columnconfigure(2, weight=1)

        self.settings_button = ctk.CTkButton(
            top_frame,
            text="设置",
            width=82,
            height=26,
            font=("Microsoft YaHei UI", 12),
            command=self.open_settings,
        )
        self.settings_button.grid(row=0, column=0, padx=(0, 8), pady=0, sticky="w")

        self.help_button = ctk.CTkButton(
            top_frame,
            text="使用说明",
            width=92,
            height=26,
            font=("Microsoft YaHei UI", 12),
            command=self.open_help_url,
        )
        self.help_button.grid(row=0, column=1, padx=(0, 8), pady=0, sticky="w")

        folder_frame = ctk.CTkFrame(self)
        folder_frame.grid(row=1, column=0, padx=16, pady=(4, 8), sticky="ew")
        folder_frame.grid_columnconfigure(1, weight=1)

        folder_label = ctk.CTkLabel(folder_frame, text="文件夹地址", width=90, anchor="w")
        folder_label.grid(row=0, column=0, padx=(12, 8), pady=12)

        self.folder_entry = ctk.CTkEntry(folder_frame, placeholder_text="请选择或粘贴需要检索的文件夹")
        self.folder_entry.grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        self.folder_entry.bind("<FocusOut>", lambda _event: self._save_cache_from_inputs())

        choose_button = ctk.CTkButton(folder_frame, text="选择文件夹", width=120, command=self.choose_folder)
        choose_button.grid(row=0, column=2, padx=(8, 12), pady=12)

        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=2, column=0, padx=16, pady=8, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)

        keyword_label = ctk.CTkLabel(input_frame, text="输入内容", width=90, anchor="w")
        keyword_label.grid(row=0, column=0, padx=(12, 8), pady=12)

        self.keyword_entry = ctk.CTkEntry(input_frame, placeholder_text="可输入工作簿、工作表、单元格、文件名或文件地址关键词；留空则显示全部")
        self.keyword_entry.grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        self.keyword_entry.bind("<Return>", lambda _event: self.run_selected_search())
        self.keyword_entry.bind("<FocusOut>", lambda _event: self._save_cache_from_inputs())

        action_frame = ctk.CTkFrame(self)
        action_frame.grid(row=3, column=0, padx=16, pady=8, sticky="ew")
        action_frame.grid_columnconfigure(1, weight=1)

        self.search_mode_menu = ctk.CTkOptionMenu(
            action_frame,
            values=["工作表检索", "单元格检索", "文件检索"],
            variable=self.search_mode_var,
            width=140,
            fg_color="#ffffff",
            button_color="#e5e7eb",
            button_hover_color="#d1d5db",
            dropdown_fg_color="#ffffff",
            dropdown_hover_color="#e5e7eb",
            dropdown_text_color="#111827",
            text_color="#111827",
        )
        self.search_mode_menu.grid(row=0, column=0, padx=(12, 8), pady=(12, 6), sticky="w")

        self.cancel_search_button = ctk.CTkButton(
            action_frame,
            text="取消",
            width=82,
            fg_color="#dc2626",
            hover_color="#b91c1c",
            text_color="#ffffff",
            command=self.cancel_current_search,
        )
        self.cancel_search_button.grid(row=0, column=2, padx=(8, 0), pady=(12, 6), sticky="e")
        self.cancel_search_button.grid_remove()

        self.run_search_button = ctk.CTkButton(action_frame, text="检索", width=110, command=self.run_selected_search)
        self.run_search_button.grid(row=0, column=3, padx=(8, 12), pady=(12, 6), sticky="e")

        self.status_label = ctk.CTkLabel(action_frame, text="请选择文件夹后执行检索", anchor="w")
        self.status_label.grid(row=0, column=1, padx=8, pady=(12, 6), sticky="ew")

        self.only_enabled_sheets_checkbox = ctk.CTkCheckBox(
            action_frame,
            text="仅检索启用表格",
            variable=self.only_enabled_sheets_var,
            command=self._save_cache_from_inputs,
        )
        self.only_enabled_sheets_checkbox.grid(row=1, column=0, padx=(12, 8), pady=(0, 12), sticky="w")

        self.only_enabled_data_checkbox = ctk.CTkCheckBox(
            action_frame,
            text="仅检索启用数据",
            variable=self.only_enabled_data_var,
            command=self._save_cache_from_inputs,
        )
        self.only_enabled_data_checkbox.grid(row=1, column=1, padx=8, pady=(0, 12), sticky="w")

        list_frame = ctk.CTkFrame(self)
        list_frame.grid(row=4, column=0, padx=16, pady=(8, 16), sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)

        self._configure_tree_style()
        self.result_tree = ttk.Treeview(
            list_frame,
            columns=(),
            show="headings",
            selectmode="browse",
        )
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        self.result_tree.bind("<ButtonRelease-1>", self._handle_result_click)
        self.result_tree.bind("<Double-1>", self._handle_result_double_click)
        self.result_tree.bind("<Configure>", lambda _event: self._schedule_action_button_layout())
        self.result_tree.tag_configure("reference", foreground="#2563eb")
        self._configure_result_columns(WORKSHEET_RESULT_COLUMNS)

        vertical_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self._scroll_result_tree_y)
        vertical_scroll.grid(row=0, column=1, sticky="ns")
        horizontal_scroll = ttk.Scrollbar(list_frame, orient="horizontal", command=self._scroll_result_tree_x)
        horizontal_scroll.grid(row=1, column=0, sticky="ew")
        self.result_tree.configure(
            yscrollcommand=lambda *args: self._update_scrollbar(vertical_scroll, *args),
            xscrollcommand=lambda *args: self._update_scrollbar(horizontal_scroll, *args),
        )

    def _configure_tree_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview", rowheight=30, font=("Microsoft YaHei UI", 10))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))

    def _configure_result_columns(
        self,
        columns: tuple[tuple[str, str, int, int, bool], ...],
        tree_heading: str | None = None,
        tree_width: int = 0,
        tree_min_width: int = 0,
        tree_stretch: bool = False,
    ) -> None:
        column_keys = tuple(column[0] for column in columns)
        self.current_result_columns = column_keys
        self.result_tree.configure(
            columns=column_keys,
            displaycolumns=column_keys,
            show="tree headings" if tree_heading else "headings",
        )
        self.result_tree.heading("#0", text=tree_heading or "")
        self.result_tree.column(
            "#0",
            width=tree_width,
            minwidth=tree_min_width,
            stretch=tree_stretch,
        )

        for key, title, width, min_width, stretch in columns:
            self.result_tree.heading(key, text=title)
            self.result_tree.column(key, width=width, minwidth=min_width, stretch=stretch)

    def _get_column_key(self, column_id: str) -> str | None:
        if not column_id.startswith("#"):
            return None
        try:
            column_index = int(column_id[1:]) - 1
        except ValueError:
            return None
        if column_index < 0 or column_index >= len(self.current_result_columns):
            return None
        return self.current_result_columns[column_index]

    def _create_action_button(self, item_id: str) -> None:
        button = ctk.CTkButton(
            self.result_tree,
            text="查找引用",
            width=96,
            height=24,
            corner_radius=5,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            text_color="#ffffff",
            command=lambda item=item_id: self._click_reference_button(item),
        )
        self.action_buttons[item_id] = button

    def _click_reference_button(self, item_id: str) -> None:
        if self.result_tree.exists(item_id):
            self.result_tree.selection_set(item_id)
            self.result_tree.focus(item_id)
        self._find_references_for_item(item_id)

    def _scroll_result_tree_y(self, *args: object) -> None:
        self.result_tree.yview(*args)
        self._schedule_action_button_layout()

    def _scroll_result_tree_x(self, *args: object) -> None:
        self.result_tree.xview(*args)
        self._schedule_action_button_layout()

    def _update_scrollbar(self, scrollbar: ttk.Scrollbar, *args: object) -> None:
        scrollbar.set(*args)
        self._schedule_action_button_layout()

    def _schedule_action_button_layout(self) -> None:
        if self.action_layout_job is not None:
            return
        self.action_layout_job = self.after_idle(self._layout_action_buttons)

    def _layout_action_buttons(self) -> None:
        self.action_layout_job = None
        if "action" not in self.current_result_columns:
            for button in self.action_buttons.values():
                button.place_forget()
            return

        for item_id, button in self.action_buttons.items():
            if not self.result_tree.exists(item_id):
                button.place_forget()
                continue

            bbox = self.result_tree.bbox(item_id, "action")
            if not bbox:
                button.place_forget()
                continue

            x, y, width, height = bbox
            button_width = min(104, max(86, width - 12))
            button_height = min(24, max(20, height - 6))
            button.place(
                x=x + max(4, (width - button_width) // 2),
                y=y + max(3, (height - button_height) // 2),
                width=button_width,
                height=button_height,
            )

    def choose_folder(self) -> None:
        folder_path = filedialog.askdirectory(title="选择要检索的文件夹")
        if not folder_path:
            return

        self.folder_entry.delete(0, "end")
        self.folder_entry.insert(0, folder_path)
        self._save_cache_from_inputs()

    def open_settings(self) -> None:
        self.settings_dialog.open(
            self.reference_config,
            self.disabled_sheet_marker,
            self.data_filter_config,
            self.help_url,
        )

    def open_help_url(self) -> None:
        url = normalize_help_url(self.help_url)
        if not url:
            messagebox.showinfo("使用说明", "请先在设置的“关于”页签里填写使用说明链接。")
            return
        webbrowser.open(url)

    def _apply_settings(
        self,
        reference_config: ReferenceLookupConfig,
        disabled_sheet_marker: str,
        data_filter_config: DataFilterConfig,
        help_url: str,
    ) -> None:
        self.reference_config = reference_config
        self.disabled_sheet_marker = disabled_sheet_marker
        self.data_filter_config = data_filter_config
        self.help_url = help_url.strip()
        self._save_cache_from_inputs()
        self._refresh_sheet_statuses()
        self.status_label.configure(text="设置已保存")

    def run_selected_search(self) -> None:
        search_mode = self.search_mode_var.get()
        if search_mode == "单元格检索":
            self.start_cell_search()
            return
        if search_mode == "文件检索":
            self.start_file_search()
            return
        self.start_scan()

    def start_scan(self) -> None:
        folder_text = self.folder_entry.get().strip()
        if not folder_text:
            messagebox.showinfo("缺少文件夹", "请先选择或输入一个文件夹地址。")
            return

        folder_path = Path(folder_text)
        if not folder_path.exists() or not folder_path.is_dir():
            messagebox.showerror("文件夹无效", "请输入有效的文件夹地址。")
            return

        self.current_keyword = self.keyword_entry.get().strip()
        self._save_cache_from_inputs()
        cancel_event = self._begin_search()
        self.status_label.configure(text="正在检索 Excel 文件，请稍等")
        self._clear_results()

        worker = Thread(
            target=self._scan_in_background,
            args=(folder_path, self.current_keyword, cancel_event),
            daemon=True,
        )
        worker.start()

    def start_cell_search(self) -> None:
        folder_text = self.folder_entry.get().strip()
        if not folder_text:
            messagebox.showinfo("缺少文件夹", "请先选择或输入一个文件夹地址。")
            return

        folder_path = Path(folder_text)
        if not folder_path.exists() or not folder_path.is_dir():
            messagebox.showerror("文件夹无效", "请输入有效的文件夹地址。")
            return

        keyword = self.keyword_entry.get().strip()
        if not keyword:
            messagebox.showinfo("缺少搜索内容", "请先在输入框里填写要搜索的内容。")
            return

        self._save_cache_from_inputs()
        cancel_event = self._begin_search()
        disabled_marker = self.disabled_sheet_marker if self.only_enabled_sheets_var.get() else None
        data_filter_config = self._get_cell_search_data_filter_config()
        search_scope_parts: list[str] = []
        if disabled_marker is not None:
            search_scope_parts.append("启用工作表")
        if self.only_enabled_data_var.get():
            search_scope_parts.append("启用数据")
        if self.data_filter_config.enable_header_filter:
            search_scope_parts.append(f"跳过前{self.data_filter_config.header_row_count}行")
        search_scope = "、".join(search_scope_parts)
        scope_text = f"{search_scope}中" if search_scope else ""
        self.status_label.configure(text=f"正在搜索{scope_text}包含“{keyword}”的单元格")

        worker = Thread(
            target=self._search_cells_in_background,
            args=(folder_path, keyword, disabled_marker, data_filter_config, cancel_event),
            daemon=True,
        )
        worker.start()

    def start_file_search(self) -> None:
        folder_text = self.folder_entry.get().strip()
        if not folder_text:
            messagebox.showinfo("缺少文件夹", "请先选择或输入一个文件夹地址。")
            return

        folder_path = Path(folder_text)
        if not folder_path.exists() or not folder_path.is_dir():
            messagebox.showerror("文件夹无效", "请输入有效的文件夹地址。")
            return

        keyword = self.keyword_entry.get().strip()
        self._save_cache_from_inputs()
        cancel_event = self._begin_search()
        self.status_label.configure(text="正在检索文件名和文件地址，请稍等")
        self._clear_results()

        worker = Thread(
            target=self._search_files_in_background,
            args=(folder_path, keyword, cancel_event),
            daemon=True,
        )
        worker.start()

    def _get_cell_search_data_filter_config(self) -> DataFilterConfig | None:
        config = self.data_filter_config
        include_enabled_data_filters = self.only_enabled_data_var.get()
        if include_enabled_data_filters:
            return config
        if not config.enable_header_filter:
            return None
        return DataFilterConfig(
            enable_header_filter=True,
            header_row_count=config.header_row_count,
        )

    def _scan_in_background(self, folder_path: Path, keyword: str, cancel_event: Event) -> None:
        try:
            results = scan_excel_workbooks(folder_path, keyword, cancel_event=cancel_event)
        except ExcelReadError as exc:
            if cancel_event.is_set():
                self.after(0, self._search_cancelled)
            else:
                self.after(0, lambda: self._scan_failed(str(exc)))
            return

        if cancel_event.is_set():
            self.after(0, self._search_cancelled)
            return
        self.after(0, lambda: self._scan_finished(results))

    def _search_cells_in_background(
        self,
        folder_path: Path,
        keyword: str,
        disabled_marker: str | None,
        data_filter_config: DataFilterConfig | None,
        cancel_event: Event,
    ) -> None:
        try:
            matches = search_excel_cells(
                folder_path,
                keyword,
                disabled_sheet_marker=disabled_marker,
                data_filter_config=data_filter_config,
                cancel_event=cancel_event,
            )
        except ExcelReadError as exc:
            if cancel_event.is_set():
                self.after(0, self._search_cancelled)
            else:
                self.after(0, lambda: self._cell_search_failed(str(exc)))
            return

        if cancel_event.is_set():
            self.after(0, self._search_cancelled)
            return
        self.after(0, lambda: self._cell_search_finished(matches, keyword))

    def _search_files_in_background(self, folder_path: Path, keyword: str, cancel_event: Event) -> None:
        try:
            results = search_files(folder_path, keyword, cancel_event=cancel_event)
        except ExcelReadError as exc:
            if cancel_event.is_set():
                self.after(0, self._search_cancelled)
            else:
                self.after(0, lambda: self._file_search_failed(str(exc)))
            return

        if cancel_event.is_set():
            self.after(0, self._search_cancelled)
            return
        self.after(0, lambda: self._file_search_finished(results))

    def _begin_search(self) -> Event:
        cancel_event = Event()
        self.search_cancel_event = cancel_event
        self._set_search_controls_running(True)
        return cancel_event

    def cancel_current_search(self) -> None:
        if self.search_cancel_event is None:
            return

        self.search_cancel_event.set()
        if self.cancel_search_button is not None:
            self.cancel_search_button.configure(state="disabled", text="取消中...")
        self.status_label.configure(text="正在取消检索，请稍等")

    def _search_cancelled(self) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="已取消检索")

    def _scan_failed(self, message: str) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="检索失败")
        messagebox.showerror("检索失败", message)

    def _scan_finished(self, results: list[ExcelWorkbookInfo]) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_results(results)
        if not results:
            return

        workbook_count = len(results)
        row_count = self._count_display_rows(results)
        if row_count == 0 and self.only_enabled_sheets_var.get():
            self.status_label.configure(text="检索完成：未找到匹配的启用工作表")
            return
        self.status_label.configure(text=f"检索完成：{workbook_count} 个工作簿，{row_count} 条结果")

    def _cell_search_failed(self, message: str) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="单元格搜索失败")
        messagebox.showerror("单元格搜索失败", message)

    def _cell_search_finished(self, matches: list[ExcelCellMatch], keyword: str) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_cell_matches(matches)
        self.status_label.configure(text=f"单元格搜索完成：找到 {len(matches)} 条匹配")

    def _file_search_failed(self, message: str) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="文件检索失败")
        messagebox.showerror("文件检索失败", message)

    def _file_search_finished(self, results: list[FileSearchResult]) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_file_results(results)
        self.status_label.configure(text=f"文件检索完成：找到 {len(results)} 个文件")

    def _set_search_controls_running(self, is_running: bool) -> None:
        if is_running:
            self.run_search_button.configure(state="disabled", text="检索中...")
            self.search_mode_menu.configure(state="disabled")
            if self.cancel_search_button is not None:
                self.cancel_search_button.configure(state="normal", text="取消")
                self.cancel_search_button.grid()
            if self.settings_button is not None:
                self.settings_button.configure(state="disabled")
            self.only_enabled_sheets_checkbox.configure(state="disabled")
            self.only_enabled_data_checkbox.configure(state="disabled")
            return

        self.run_search_button.configure(state="normal", text="检索")
        self.search_mode_menu.configure(state="normal")
        if self.cancel_search_button is not None:
            self.cancel_search_button.grid_remove()
        if self.settings_button is not None:
            self.settings_button.configure(state="normal")
        self.only_enabled_sheets_checkbox.configure(state="normal")
        self.only_enabled_data_checkbox.configure(state="normal")

    def _render_cell_matches(self, matches: list[ExcelCellMatch]) -> None:
        self._clear_results()
        self._configure_result_columns(
            CELL_RESULT_COLUMNS,
            tree_heading="工作簿",
            tree_width=260,
            tree_min_width=180,
            tree_stretch=True,
        )

        matches_by_workbook: dict[Path, list[ExcelCellMatch]] = {}
        for match in matches:
            matches_by_workbook.setdefault(match.path, []).append(match)

        for workbook_path, workbook_matches in matches_by_workbook.items():
            parent_id = self.result_tree.insert(
                "",
                "end",
                text=workbook_path.name,
                values=(workbook_path.name, "", "", "", f"{len(workbook_matches)} 条匹配"),
                open=True,
            )
            self.item_targets[parent_id] = ResultTarget(workbook_path)

            for match in workbook_matches:
                item_id = self.result_tree.insert(
                    parent_id,
                    "end",
                    text="",
                    values=(
                        match.workbook_name,
                        match.sheet_name,
                        match.row_index,
                        match.column_index,
                        format_cell_value(match.value),
                    ),
                )
                self.item_targets[item_id] = ResultTarget(
                    match.path,
                    match.sheet_name,
                    match.row_index,
                    match.column_index,
                )

    def _render_file_results(self, results: list[FileSearchResult]) -> None:
        self._clear_results()
        self._configure_result_columns(FILE_RESULT_COLUMNS)

        if not results:
            self.status_label.configure(text="没有找到匹配的文件")
            return

        for result in results:
            item_id = self.result_tree.insert(
                "",
                "end",
                values=(result.file_name, result.file_address),
            )
            self.item_targets[item_id] = ResultTarget(result.path, open_containing_folder=True)

    def _toggle_tree_item(self, item_id: str) -> None:
        if self.result_tree.get_children(item_id):
            self.result_tree.item(item_id, open=not bool(self.result_tree.item(item_id, "open")))

    def _open_or_toggle_selected_item(self, item_id: str) -> None:
        target = self.item_targets.get(item_id)
        if target is None:
            return
        if target.sheet_name is None and self.result_tree.get_children(item_id):
            self._toggle_tree_item(item_id)
            return
        self._open_item(item_id)

    def _open_item(self, item_id: str) -> None:
        target = self.item_targets.get(item_id)
        if target is None:
            messagebox.showerror("无法打开", "没有找到对应的文件地址。")
            return

        try:
            if target.open_containing_folder:
                open_containing_folder(target.file_path)
                return
            open_file(target.file_path, target.sheet_name, target.row_index, target.column_index)
        except (OSError, subprocess.SubprocessError) as exc:
            messagebox.showerror("打开失败", f"无法打开文件：{target.file_path}\n{exc}")

    def _clear_results(self) -> None:
        self.item_targets.clear()
        self.reference_items_by_source.clear()
        self.reference_searching_items.clear()
        for button in self.action_buttons.values():
            button.destroy()
        self.action_buttons.clear()
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

    def _render_results(self, results: list[ExcelWorkbookInfo]) -> None:
        self._clear_results()
        self._configure_result_columns(WORKSHEET_RESULT_COLUMNS)

        if not results:
            self.status_label.configure(text="没有找到匹配的 Excel 文件")
            return

        has_display_rows = False
        for result in results:
            if result.error:
                item = self.result_tree.insert(
                    "",
                    "end",
                    values=(result.workbook_name, "无法读取", "读取失败", ""),
                )
                self.item_targets[item] = ResultTarget(result.path)
                has_display_rows = True
                continue

            sheet_names = self._get_display_sheet_names(result)
            if not sheet_names:
                if self.only_enabled_sheets_var.get():
                    continue
                item = self.result_tree.insert(
                    "",
                    "end",
                    values=(result.workbook_name, "无匹配工作表", "", ""),
                )
                self.item_targets[item] = ResultTarget(result.path)
                has_display_rows = True
                continue

            for sheet_name in sheet_names:
                item = self.result_tree.insert(
                    "",
                    "end",
                    values=(
                        result.workbook_name,
                        sheet_name,
                        get_sheet_enabled_status(sheet_name, self.disabled_sheet_marker),
                        "查找引用",
                    ),
                )
                self.item_targets[item] = ResultTarget(result.path, sheet_name)
                self._create_action_button(item)
                has_display_rows = True

        if not has_display_rows:
            self.status_label.configure(text="没有找到匹配的启用工作表")
        self._schedule_action_button_layout()

    def _handle_result_click(self, event: object) -> str | None:
        column_key = self._get_column_key(self.result_tree.identify_column(event.x))
        if column_key != "action":
            return None

        item_id = self.result_tree.identify_row(event.y)
        if not item_id:
            return "break"
        if not self._get_action_text(item_id):
            return "break"

        self.result_tree.selection_set(item_id)
        self.result_tree.focus(item_id)
        self._find_references_for_item(item_id)
        return "break"

    def _handle_result_double_click(self, event: object) -> str:
        column_key = self._get_column_key(self.result_tree.identify_column(event.x))
        if column_key == "action":
            item_id = self.result_tree.identify_row(event.y)
            if item_id:
                self.result_tree.selection_set(item_id)
                self.result_tree.focus(item_id)
                self._find_references_for_item(item_id)
            return "break"

        item_id = self.result_tree.identify_row(event.y)
        if item_id:
            self.result_tree.selection_set(item_id)
            self.result_tree.focus(item_id)
            self._open_or_toggle_selected_item(item_id)
        return "break"

    def _find_references_for_item(self, item_id: str) -> None:
        target = self.item_targets.get(item_id)
        if target is None or not target.sheet_name:
            return
        if item_id in self.reference_searching_items:
            return
        if not self._get_action_text(item_id):
            return

        button = self.action_buttons.get(item_id)

        config = self.reference_config
        self.reference_searching_items.add(item_id)
        if button is not None:
            button.configure(text="查找中...", state="disabled")
        self._set_action_text(item_id, "查找中...")
        self._remove_reference_rows(item_id)
        self.status_label.configure(text=f"正在查找 {target.sheet_name} 第{config.reference_row_index}行引用")

        worker = Thread(target=self._find_references_in_background, args=(item_id, target, config), daemon=True)
        worker.start()

    def _find_references_in_background(
        self,
        item_id: str,
        target: ResultTarget,
        config: ReferenceLookupConfig,
    ) -> None:
        try:
            references = find_sheet_reference_matches(target.file_path, target.sheet_name or "", config)
        except ExcelReadError as exc:
            self.after(0, lambda: self._references_failed(item_id, str(exc)))
            return

        self.after(0, lambda: self._references_finished(item_id, target.file_path, references, config))

    def _references_failed(self, item_id: str, message: str) -> None:
        self.reference_searching_items.discard(item_id)
        button = self.action_buttons.get(item_id)
        if button is not None:
            button.configure(text="查找引用", state="normal")
        self._set_action_text(item_id, "查找引用")
        self.status_label.configure(text=f"查找引用失败：{message}")

    def _references_finished(
        self,
        item_id: str,
        file_path: Path,
        references: list[ExcelSheetReference],
        config: ReferenceLookupConfig,
    ) -> None:
        self.reference_searching_items.discard(item_id)
        button = self.action_buttons.get(item_id)
        if button is not None:
            button.configure(text="查找引用", state="normal")
        self._set_action_text(item_id, "查找引用")
        self.status_label.configure(
            text=f"找到 {len(references)} 个引用" if references else f"第{config.reference_row_index}行未找到引用"
        )

        self._remove_reference_rows(item_id)
        if not references:
            empty_item = self._insert_reference_row(item_id, file_path, "未找到引用", can_open=False)
            self.reference_items_by_source[item_id] = [empty_item]
            return

        inserted_items: list[str] = []
        for reference in references:
            inserted_items.append(
                self._insert_reference_row(
                    item_id,
                    file_path,
                    reference,
                    insert_offset=len(inserted_items),
                )
            )
        self.reference_items_by_source[item_id] = inserted_items

    def _insert_reference_row(
        self,
        source_item_id: str,
        file_path: Path,
        reference: ExcelSheetReference | str,
        can_open: bool = True,
        insert_offset: int = 0,
    ) -> str:
        insert_index = self.result_tree.index(source_item_id) + 1 + insert_offset
        reference_name = reference.sheet_name if isinstance(reference, ExcelSheetReference) else reference
        reference_text = format_reference_label(reference) if isinstance(reference, ExcelSheetReference) else reference
        status = get_sheet_enabled_status(reference_name, self.disabled_sheet_marker) if can_open else ""
        item = self.result_tree.insert(
            "",
            insert_index,
            values=("", reference_text if can_open else reference_name, status, ""),
            tags=("reference",),
        )
        if can_open:
            self.item_targets[item] = ResultTarget(file_path, reference_name)
        self._schedule_action_button_layout()
        return item

    def _remove_reference_rows(self, source_item_id: str) -> None:
        for item_id in self.reference_items_by_source.pop(source_item_id, []):
            self.item_targets.pop(item_id, None)
            if self.result_tree.exists(item_id):
                self.result_tree.delete(item_id)
        self._schedule_action_button_layout()

    def _set_action_text(self, item_id: str, text: str) -> None:
        if "action" not in self.current_result_columns or not self.result_tree.exists(item_id):
            return

        action_index = self.current_result_columns.index("action")
        values = list(self.result_tree.item(item_id, "values"))
        if action_index >= len(values):
            return
        values[action_index] = text
        self.result_tree.item(item_id, values=values)

    def _get_action_text(self, item_id: str) -> str:
        if "action" not in self.current_result_columns or not self.result_tree.exists(item_id):
            return ""

        action_index = self.current_result_columns.index("action")
        values = list(self.result_tree.item(item_id, "values"))
        if action_index >= len(values):
            return ""
        return str(values[action_index]).strip()

    def _refresh_sheet_statuses(self) -> None:
        if "status" not in self.current_result_columns:
            return

        status_index = self.current_result_columns.index("status")
        for item_id in self.result_tree.get_children():
            target = self.item_targets.get(item_id)
            if target is None or not target.sheet_name:
                continue

            values = list(self.result_tree.item(item_id, "values"))
            if status_index >= len(values):
                continue
            values[status_index] = get_sheet_enabled_status(target.sheet_name, self.disabled_sheet_marker)
            self.result_tree.item(item_id, values=values)

    def _count_display_rows(self, results: list[ExcelWorkbookInfo]) -> int:
        row_count = 0
        for result in results:
            if result.error:
                row_count += 1
                continue
            sheet_count = len(self._get_display_sheet_names(result))
            if sheet_count:
                row_count += sheet_count
            elif not self.only_enabled_sheets_var.get():
                row_count += 1
        return row_count

    def _get_display_sheet_names(self, result: ExcelWorkbookInfo) -> list[str]:
        keyword = self.current_keyword.lower()
        if not keyword:
            return self._filter_enabled_sheet_names(result.sheet_names)

        workbook_matches = keyword in result.workbook_name.lower() or keyword in str(result.path).lower()
        if workbook_matches:
            return self._filter_enabled_sheet_names(result.sheet_names)

        return self._filter_enabled_sheet_names(
            [sheet_name for sheet_name in result.sheet_names if keyword in sheet_name.lower()]
        )

    def _filter_enabled_sheet_names(self, sheet_names: list[str]) -> list[str]:
        if not self.only_enabled_sheets_var.get():
            return sheet_names
        return [
            sheet_name
            for sheet_name in sheet_names
            if get_sheet_enabled_status(sheet_name, self.disabled_sheet_marker) == "启用"
        ]

    def open_selected_item(self) -> None:
        selected_items = self.result_tree.selection()
        if not selected_items:
            messagebox.showinfo("未选择", "请先在列表中选择一个工作簿或工作表。")
            return
        self._open_or_toggle_selected_item(selected_items[0])

    def _load_cached_inputs(self) -> None:
        cache_data = load_cache_data()
        folder_path = cache_data.get("folder_path", "")
        keyword = cache_data.get("keyword", "")
        reference_config = cache_data.get("reference_config", DEFAULT_REFERENCE_LOOKUP_CONFIG)
        if isinstance(reference_config, ReferenceLookupConfig):
            self.reference_config = reference_config
        disabled_sheet_marker = cache_data.get("disabled_sheet_marker", DEFAULT_DISABLED_SHEET_MARKER)
        if isinstance(disabled_sheet_marker, str):
            self.disabled_sheet_marker = disabled_sheet_marker
        data_filter_config = cache_data.get("data_filter_config", DEFAULT_DATA_FILTER_CONFIG)
        if isinstance(data_filter_config, DataFilterConfig):
            self.data_filter_config = data_filter_config
        help_url = cache_data.get("help_url", "")
        if isinstance(help_url, str):
            self.help_url = help_url
        self.only_enabled_sheets_var.set(bool(cache_data.get("only_enabled_sheets", False)))
        self.only_enabled_data_var.set(bool(cache_data.get("only_enabled_data", False)))

        if folder_path:
            self.folder_entry.insert(0, folder_path)
        if keyword:
            self.keyword_entry.insert(0, keyword)

        if folder_path or keyword:
            self.status_label.configure(text="已加载上次选择的地址和输入内容")

    def _save_cache_from_inputs(self) -> None:
        save_cache_data(
            {
                "folder_path": self.folder_entry.get().strip(),
                "keyword": self.keyword_entry.get().strip(),
                "reference_config": self.reference_config,
                "disabled_sheet_marker": self.disabled_sheet_marker,
                "data_filter_config": self.data_filter_config,
                "help_url": self.help_url,
                "only_enabled_sheets": self.only_enabled_sheets_var.get(),
                "only_enabled_data": self.only_enabled_data_var.get(),
            }
        )

    def close_app(self) -> None:
        if self.search_cancel_event is not None:
            self.search_cancel_event.set()
        self._save_cache_from_inputs()
        self.destroy()


def load_cache_data() -> dict[str, object]:
    raw_data = read_cache_file()
    if not raw_data:
        return {}

    folder_path = raw_data.get("folder_path", "")
    keyword = raw_data.get("keyword", "")
    disabled_sheet_marker = raw_data.get("disabled_sheet_marker", DEFAULT_DISABLED_SHEET_MARKER)
    help_url = raw_data.get("help_url", "")
    return {
        "folder_path": folder_path if isinstance(folder_path, str) else "",
        "keyword": keyword if isinstance(keyword, str) else "",
        "reference_config": load_reference_config(raw_data.get("reference_config")),
        "disabled_sheet_marker": disabled_sheet_marker if isinstance(disabled_sheet_marker, str) else "",
        "data_filter_config": load_data_filter_config(raw_data.get("data_filter_config")),
        "help_url": help_url if isinstance(help_url, str) else "",
        "only_enabled_sheets": bool(raw_data.get("only_enabled_sheets", False)),
        "only_enabled_data": bool(raw_data.get("only_enabled_data", False)),
    }


def read_cache_file() -> dict[str, object]:
    try:
        raw_data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}

    if not isinstance(raw_data, dict):
        return {}

    return raw_data


def save_cache_data(cache_data: dict[str, object]) -> None:
    reference_config = cache_data.get("reference_config", DEFAULT_REFERENCE_LOOKUP_CONFIG)
    if not isinstance(reference_config, ReferenceLookupConfig):
        reference_config = DEFAULT_REFERENCE_LOOKUP_CONFIG
    data_filter_config = cache_data.get("data_filter_config", DEFAULT_DATA_FILTER_CONFIG)
    if not isinstance(data_filter_config, DataFilterConfig):
        data_filter_config = DEFAULT_DATA_FILTER_CONFIG

    payload = read_cache_file()
    payload.update({
        "folder_path": str(cache_data.get("folder_path", "")),
        "keyword": str(cache_data.get("keyword", "")),
        "reference_config": reference_config_to_dict(reference_config),
        "disabled_sheet_marker": str(cache_data.get("disabled_sheet_marker", DEFAULT_DISABLED_SHEET_MARKER)),
        "data_filter_config": data_filter_config_to_dict(data_filter_config),
        "help_url": str(cache_data.get("help_url", "")),
        "only_enabled_sheets": bool(cache_data.get("only_enabled_sheets", False)),
        "only_enabled_data": bool(cache_data.get("only_enabled_data", False)),
    })
    worksheet_sheet_cache = cache_data.get(WORKSHEET_SHEET_CACHE_KEY, payload.get(WORKSHEET_SHEET_CACHE_KEY, {}))
    if isinstance(worksheet_sheet_cache, dict):
        payload[WORKSHEET_SHEET_CACHE_KEY] = worksheet_sheet_cache
    try:
        CACHE_FILE.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        pass


def load_reference_config(raw_config: object) -> ReferenceLookupConfig:
    if not isinstance(raw_config, dict):
        return DEFAULT_REFERENCE_LOOKUP_CONFIG

    try:
        config = ReferenceLookupConfig(
            sample_text=str(raw_config.get("sample_text", DEFAULT_REFERENCE_LOOKUP_CONFIG.sample_text)),
            reference_row_index=parse_positive_int(
                raw_config.get("reference_row_index", DEFAULT_REFERENCE_LOOKUP_CONFIG.reference_row_index),
                "引用行",
            ),
            table_name=str(raw_config.get("table_name", DEFAULT_REFERENCE_LOOKUP_CONFIG.table_name)),
            field_name=str(raw_config.get("field_name", DEFAULT_REFERENCE_LOOKUP_CONFIG.field_name)),
            field_row_index=parse_positive_int(
                raw_config.get("field_row_index", DEFAULT_REFERENCE_LOOKUP_CONFIG.field_row_index),
                "引用字段行",
            ),
        )
        validate_reference_lookup_config(config)
    except ExcelReadError:
        return DEFAULT_REFERENCE_LOOKUP_CONFIG
    return config


def load_data_filter_config(raw_config: object) -> DataFilterConfig:
    if not isinstance(raw_config, dict):
        return DEFAULT_DATA_FILTER_CONFIG

    try:
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
            ),
            enable_column_filter=bool(raw_config.get("enable_column_filter", False)),
            column_marker_row_index=parse_positive_int(
                raw_config.get(
                    "column_marker_row_index",
                    DEFAULT_DATA_FILTER_CONFIG.column_marker_row_index,
                ),
                "列判断行",
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
            ),
            range_start_text=str(raw_config.get("range_start_text", "")),
            range_end_text=str(raw_config.get("range_end_text", "")),
        )
        validate_data_filter_config(config)
    except ExcelReadError:
        return DEFAULT_DATA_FILTER_CONFIG
    return config


def reference_config_to_dict(config: ReferenceLookupConfig) -> dict[str, object]:
    return {
        "sample_text": config.sample_text,
        "reference_row_index": config.reference_row_index,
        "table_name": config.table_name,
        "field_name": config.field_name,
        "field_row_index": config.field_row_index,
    }


def data_filter_config_to_dict(config: DataFilterConfig) -> dict[str, object]:
    return {
        "enable_header_filter": config.enable_header_filter,
        "header_row_count": config.header_row_count,
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


def get_sheet_enabled_status(sheet_name: str, disabled_marker: str) -> str:
    if not sheet_name:
        return ""
    marker = disabled_marker.strip()
    return "未启用" if marker and sheet_name.startswith(marker) else "启用"


def normalize_help_url(help_url: str) -> str:
    url = help_url.strip()
    if not url:
        return ""
    known_schemes = ("http://", "https://", "file://", "mailto:")
    if url.lower().startswith(known_schemes):
        return url
    return f"https://{url}"


def format_cell_value(value: str) -> str:
    return " ".join(value.splitlines())


def format_reference_label(reference: ExcelSheetReference) -> str:
    parts = [f"引用表：{reference.sheet_name}"]
    if reference.reference_field_name:
        parts.append(f"字段：{reference.reference_field_name}")
    if reference.source_field_name:
        parts.append(f"本表字段：{reference.source_field_name}")
    return "，".join(parts)


def open_containing_folder(file_path: Path) -> bool:
    if file_path.is_dir():
        folder_path = file_path
    else:
        folder_path = file_path.parent

    if sys.platform.startswith("win"):
        if file_path.exists() and file_path.is_file():
            subprocess.Popen(["explorer", f"/select,{file_path}"])
            return True
        os.startfile(str(folder_path))
        return True
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(folder_path)])
        return False
    subprocess.Popen(["xdg-open", str(folder_path)])
    return False


def open_file(
    file_path: Path,
    sheet_name: str | None = None,
    row_index: int | None = None,
    column_index: int | None = None,
) -> bool:
    if sys.platform.startswith("win"):
        if sheet_name:
            try:
                open_excel_sheet(file_path, sheet_name, row_index, column_index)
                return True
            except (OSError, subprocess.SubprocessError):
                os.startfile(str(file_path))
                return False

        os.startfile(str(file_path))
        return True
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(file_path)])
        return False
    subprocess.Popen(["xdg-open", str(file_path)])
    return False


def open_excel_sheet(
    file_path: Path,
    sheet_name: str,
    row_index: int | None = None,
    column_index: int | None = None,
) -> None:
    file_arg = powershell_quote(str(file_path))
    sheet_arg = powershell_quote(sheet_name)
    row_arg = row_index if row_index is not None else "$null"
    column_arg = column_index if column_index is not None else "$null"
    script = f"""
$ErrorActionPreference = 'Stop'
$filePath = {file_arg}
$sheetName = {sheet_arg}
$rowIndex = {row_arg}
$columnIndex = {column_arg}
$excel = $null

try {{
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
}} catch {{
    $excel = New-Object -ComObject Excel.Application
}}

$excel.Visible = $true
$workbook = $null

foreach ($book in @($excel.Workbooks)) {{
    if ([string]::Equals($book.FullName, $filePath, [System.StringComparison]::OrdinalIgnoreCase)) {{
        $workbook = $book
        break
    }}
}}

if ($null -eq $workbook) {{
    $workbook = $excel.Workbooks.Open($filePath)
}}

$worksheet = $workbook.Worksheets.Item($sheetName)
$workbook.Activate()
$worksheet.Activate()

if ($null -ne $rowIndex -and $null -ne $columnIndex) {{
    $cell = $worksheet.Cells.Item($rowIndex, $columnIndex)
    $cell.Select()
    $excel.ActiveWindow.ScrollRow = [Math]::Max(1, $rowIndex - 5)
    $excel.ActiveWindow.ScrollColumn = [Math]::Max(1, $columnIndex - 3)
}}

$excel.WindowState = -4137
"""
    subprocess.run(
        ["powershell", "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-Command", script],
        check=True,
        capture_output=True,
        text=True,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )


def powershell_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def run_app() -> None:
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    app = ExcelEditApp()
    app.mainloop()
