"""应用主界面和用户交互逻辑。

这个文件负责窗口布局、缓存输入、启动后台检索线程、动态切换结果表格结构、
双击打开文件/工作表/单元格，以及在工作表结果中查找引用关系。
"""

import json
import os
import subprocess
import sys
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from threading import Event, Thread
from tkinter import filedialog, font as tkfont, messagebox, ttk

import customtkinter as ctk

from cell_search import (
    DEFAULT_DATA_FILTER_CONFIG,
    DataFilterConfig,
    ExcelCellMatch,
    search_excel_cells,
    validate_data_filter_config,
)
from excel_common import ExcelReadError
from file_search import FileSearchError, FileSearchResult, search_files
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
    """获取应用运行目录，兼容源码运行和 PyInstaller 打包后的 EXE。"""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_resource_path(relative_path: str) -> Path:
    """获取资源文件路径，兼容 PyInstaller 临时解包目录。"""

    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        return Path(bundle_dir) / relative_path
    return get_app_directory() / relative_path


CACHE_FILE = get_app_directory() / "cache_data.json"
APP_NAME = "奶龙检索大师"
APP_ICON_RELATIVE_PATH = "assets/nailong_search_master.ico"
DEFAULT_DISABLED_SHEET_MARKER = "$"
FOLDER_HISTORY_LIMIT = 20
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
    """结果列表项对应的打开目标。"""

    file_path: Path
    sheet_name: str | None = None
    row_index: int | None = None
    column_index: int | None = None
    open_containing_folder: bool = False


class ExcelEditApp(ctk.CTk):
    """主窗口类，管理所有控件状态和用户操作。"""

    def __init__(self) -> None:
        """初始化窗口状态、构建界面并加载上次缓存。"""

        super().__init__()

        self.title(APP_NAME)
        self._set_app_icon()
        self.geometry("1080x680")
        self.minsize(840, 520)

        self.item_targets: dict[str, ResultTarget] = {}
        self.reference_items_by_source: dict[str, list[str]] = {}
        self.reference_searching_items: set[str] = set()
        self.action_buttons: dict[str, ctk.CTkButton] = {}
        self.action_layout_job: str | None = None
        self.current_result_columns: tuple[str, ...] = ()
        self.folder_path_var = ctk.StringVar(value="")
        self.folder_history: list[str] = []
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

        # 先构建控件再加载缓存，因为缓存需要写回输入框和下拉框。
        self._build_layout()
        self._load_cached_inputs()
        self.protocol("WM_DELETE_WINDOW", self.close_app)

    def _set_app_icon(self) -> None:
        """设置窗口图标；缺少图标时不影响程序启动。"""

        icon_path = get_resource_path(APP_ICON_RELATIVE_PATH)
        if not icon_path.exists():
            return
        try:
            self.iconbitmap(default=str(icon_path))
        except Exception:
            pass

    def _build_layout(self) -> None:
        """创建主界面的所有控件和布局。"""

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # 顶部工具区：设置和帮助入口。
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

        # 文件夹地址区：下拉框记录历史路径，按钮保持在同一行方便选择。
        folder_frame = ctk.CTkFrame(self)
        folder_frame.grid(row=1, column=0, padx=16, pady=(4, 8), sticky="ew")
        folder_frame.grid_columnconfigure(1, weight=1)

        folder_label = ctk.CTkLabel(folder_frame, text="文件夹地址", width=90, anchor="w")
        folder_label.grid(row=0, column=0, padx=(12, 8), pady=12)

        self.folder_entry = ctk.CTkComboBox(
            folder_frame,
            variable=self.folder_path_var,
            values=[],
            command=self._select_folder_from_history,
            fg_color="#ffffff",
            button_color="#e5e7eb",
            button_hover_color="#d1d5db",
            dropdown_fg_color="#ffffff",
            dropdown_hover_color="#e5e7eb",
            dropdown_text_color="#111827",
            text_color="#111827",
        )
        self.folder_entry.grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        self.folder_entry.bind("<FocusOut>", lambda _event: self._save_cache_from_inputs())
        # CustomTkinter 的下拉菜单默认偏窄，这里接管打开逻辑来同步历史列表宽度。
        self.folder_entry._open_dropdown_menu = self._open_folder_dropdown_menu
        folder_frame.bind("<Configure>", lambda _event: self._sync_folder_dropdown_width(), add="+")
        self.after_idle(self._sync_folder_dropdown_width)

        choose_button = ctk.CTkButton(folder_frame, text="选择文件夹", width=120, command=self.choose_folder)
        choose_button.grid(row=0, column=2, padx=(8, 12), pady=12)

        # 关键词输入区：不同检索模式会复用同一个输入框。
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=2, column=0, padx=16, pady=8, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)

        keyword_label = ctk.CTkLabel(input_frame, text="输入内容", width=90, anchor="w")
        keyword_label.grid(row=0, column=0, padx=(12, 8), pady=12)

        self.keyword_entry = ctk.CTkEntry(input_frame, placeholder_text="可输入工作簿、工作表、单元格、文件名或文件地址关键词；留空则显示全部")
        self.keyword_entry.grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        self.keyword_entry.bind("<Return>", lambda _event: self.run_selected_search())
        self.keyword_entry.bind("<FocusOut>", lambda _event: self._save_cache_from_inputs())

        # 检索操作区：模式下拉框、取消按钮、检索按钮和状态栏。
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

        # 结果区使用 ttk.Treeview，便于按不同检索模式动态切换列结构。
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
        """统一结果列表的行高和字体样式。"""

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
        """按当前检索结果结构重新配置 Treeview 列。"""

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
        """把 Treeview 的 #1/#2 列标识转换成业务列名。"""

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
        """为工作表结果行创建覆盖在“操作”列上的查找引用按钮。"""

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
        """点击查找引用按钮时同步选中行并启动引用查找。"""

        if self.result_tree.exists(item_id):
            self.result_tree.selection_set(item_id)
            self.result_tree.focus(item_id)
        self._find_references_for_item(item_id)

    def _scroll_result_tree_y(self, *args: object) -> None:
        """纵向滚动后重新定位覆盖按钮。"""

        self.result_tree.yview(*args)
        self._schedule_action_button_layout()

    def _scroll_result_tree_x(self, *args: object) -> None:
        """横向滚动后重新定位覆盖按钮。"""

        self.result_tree.xview(*args)
        self._schedule_action_button_layout()

    def _update_scrollbar(self, scrollbar: ttk.Scrollbar, *args: object) -> None:
        """更新滚动条，同时安排按钮位置刷新。"""

        scrollbar.set(*args)
        self._schedule_action_button_layout()

    def _schedule_action_button_layout(self) -> None:
        """把按钮重排延迟到 Tk 空闲时执行，避免滚动时频繁计算。"""

        if self.action_layout_job is not None:
            return
        self.action_layout_job = self.after_idle(self._layout_action_buttons)

    def _layout_action_buttons(self) -> None:
        """把真实按钮摆放到 Treeview 的“操作”列单元格上。"""

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
        """打开系统文件夹选择器，并记录选择历史。"""

        folder_path = filedialog.askdirectory(title="选择要检索的文件夹")
        if not folder_path:
            return

        self.folder_path_var.set(folder_path)
        self._remember_folder_path(folder_path)
        self._save_cache_from_inputs()

    def _select_folder_from_history(self, folder_path: str) -> None:
        """从历史下拉框选择文件夹时，把它移到历史列表最前面。"""

        self.folder_path_var.set(folder_path)
        self._remember_folder_path(folder_path)
        self._save_cache_from_inputs()

    def _remember_folder_path(self, folder_path: str | Path) -> None:
        """把有效文件夹写入历史列表并刷新下拉框。"""

        self.folder_history = normalize_folder_history(self.folder_history, str(folder_path))
        self._refresh_folder_history_menu()

    def _refresh_folder_history_menu(self) -> None:
        """刷新文件夹历史下拉框中的选项。"""

        self.folder_entry.configure(values=self.folder_history)
        self._sync_folder_dropdown_width()

    def _sync_folder_dropdown_width(self) -> None:
        """让历史下拉列表宽度尽量贴近文件夹输入框宽度。"""

        dropdown_menu = getattr(self.folder_entry, "_dropdown_menu", None)
        if dropdown_menu is None:
            return

        width = self.folder_entry.winfo_width()
        if width <= 1:
            width = int(self.folder_entry.cget("width"))

        try:
            entry_widget = getattr(self.folder_entry, "_entry")
            char_width = tkfont.Font(font=entry_widget.cget("font")).measure("0")
        except Exception:
            char_width = 8

        min_character_width = max(18, (width // max(char_width, 1)) - 2)
        dropdown_menu.configure(min_character_width=min_character_width)
        # Tk 菜单宽度依赖菜单项文本宽度，所以需要重建带补位空格的菜单项。
        self._rebuild_folder_dropdown_menu(width)

    def _open_folder_dropdown_menu(self) -> None:
        """打开历史下拉菜单前先同步宽度。"""

        self._sync_folder_dropdown_width()
        dropdown_menu = getattr(self.folder_entry, "_dropdown_menu", None)
        if dropdown_menu is None:
            return

        current_height = getattr(self.folder_entry, "_current_height", self.folder_entry.winfo_height())
        y_offset = self.folder_entry._apply_widget_scaling(current_height + 0)
        dropdown_menu.open(self.folder_entry.winfo_rootx(), self.folder_entry.winfo_rooty() + y_offset)
        self.folder_entry._close_on_next_click = True

    def _rebuild_folder_dropdown_menu(self, target_width: int) -> None:
        """重建 CustomTkinter 下拉菜单项，让弹出的历史地址不再过窄。"""

        dropdown_menu = getattr(self.folder_entry, "_dropdown_menu", None)
        if dropdown_menu is None:
            return

        dropdown_menu.delete(0, "end")
        try:
            menu_font = tkfont.Font(font=dropdown_menu.cget("font"))
        except Exception:
            entry_widget = getattr(self.folder_entry, "_entry", None)
            menu_font = tkfont.Font(font=entry_widget.cget("font")) if entry_widget is not None else tkfont.Font()

        spacer = "\u00a0"
        spacer_width = max(menu_font.measure(spacer), 1)
        usable_width = max(target_width - 24, 0)
        for folder_path in self.folder_history:
            label = folder_path
            remaining_width = usable_width - menu_font.measure(label)
            if remaining_width > 0:
                # 使用不可断空格撑宽菜单项，不影响回调里真正返回的文件夹路径。
                label += spacer * max(1, (remaining_width + spacer_width - 1) // spacer_width)
            dropdown_menu.add_command(
                label=label,
                command=lambda selected_path=folder_path: self.folder_entry._dropdown_callback(selected_path),
                compound="left",
            )

    def open_settings(self) -> None:
        """打开设置弹窗。"""

        self.settings_dialog.open(
            self.reference_config,
            self.disabled_sheet_marker,
            self.data_filter_config,
            self.help_url,
        )

    def open_help_url(self) -> None:
        """打开用户在设置中填写的使用说明链接。"""

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
        """应用设置弹窗保存后的配置，并刷新当前列表里的启用状态。"""

        self.reference_config = reference_config
        self.disabled_sheet_marker = disabled_sheet_marker
        self.data_filter_config = data_filter_config
        self.help_url = help_url.strip()
        self._save_cache_from_inputs()
        self._refresh_sheet_statuses()
        self.status_label.configure(text="设置已保存")

    def run_selected_search(self) -> None:
        """根据下拉框当前选项分发到对应检索功能。"""

        search_mode = self.search_mode_var.get()
        if search_mode == "单元格检索":
            self.start_cell_search()
            return
        if search_mode == "文件检索":
            self.start_file_search()
            return
        self.start_scan()

    def start_scan(self) -> None:
        """启动工作表检索。"""

        folder_text = self.folder_entry.get().strip()
        if not folder_text:
            messagebox.showinfo("缺少文件夹", "请先选择或输入一个文件夹地址。")
            return

        folder_path = Path(folder_text)
        if not folder_path.exists() or not folder_path.is_dir():
            messagebox.showerror("文件夹无效", "请输入有效的文件夹地址。")
            return

        self.current_keyword = self.keyword_entry.get().strip()
        self._remember_folder_path(folder_path)
        self._save_cache_from_inputs()
        cancel_event = self._begin_search()
        self.status_label.configure(text="正在检索 Excel 文件，请稍等")
        self._clear_results()

        # Excel 文件数量可能较多，放到后台线程，避免界面卡死。
        worker = Thread(
            target=self._scan_in_background,
            args=(folder_path, self.current_keyword, cancel_event),
            daemon=True,
        )
        worker.start()

    def start_cell_search(self) -> None:
        """启动单元格内容检索。"""

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

        self._remember_folder_path(folder_path)
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
        if self.data_filter_config.enable_blank_header_column_filter:
            search_scope_parts.append("屏蔽空白表头列")
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
        """启动普通文件名和文件地址检索。"""

        folder_text = self.folder_entry.get().strip()
        if not folder_text:
            messagebox.showinfo("缺少文件夹", "请先选择或输入一个文件夹地址。")
            return

        folder_path = Path(folder_text)
        if not folder_path.exists() or not folder_path.is_dir():
            messagebox.showerror("文件夹无效", "请输入有效的文件夹地址。")
            return

        keyword = self.keyword_entry.get().strip()
        self._remember_folder_path(folder_path)
        self._save_cache_from_inputs()
        cancel_event = self._begin_search()
        self.status_label.configure(text="正在检索文件名和文件地址，请稍等")
        self._clear_results()

        # 文件名检索也可能遍历大量子目录，所以同样放到后台。
        worker = Thread(
            target=self._search_files_in_background,
            args=(folder_path, keyword, cancel_event),
            daemon=True,
        )
        worker.start()

    def _get_cell_search_data_filter_config(self) -> DataFilterConfig | None:
        """根据界面勾选项决定单元格检索时实际启用哪些数据过滤规则。"""

        config = self.data_filter_config
        include_enabled_data_filters = self.only_enabled_data_var.get()
        if include_enabled_data_filters:
            return config
        if not config.enable_header_filter and not config.enable_blank_header_column_filter:
            return None
        return DataFilterConfig(
            enable_header_filter=config.enable_header_filter,
            header_row_count=config.header_row_count,
            enable_blank_header_column_filter=config.enable_blank_header_column_filter,
        )

    def _scan_in_background(self, folder_path: Path, keyword: str, cancel_event: Event) -> None:
        """后台执行工作表检索，并把结果切回主线程渲染。"""

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
        """后台执行单元格检索，并把结果切回主线程渲染。"""

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
        """后台执行文件名检索，并把结果切回主线程渲染。"""

        try:
            results = search_files(folder_path, keyword, cancel_event=cancel_event)
        except FileSearchError as exc:
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
        """进入检索中状态，并创建本次检索的取消事件。"""

        cancel_event = Event()
        self.search_cancel_event = cancel_event
        self._set_search_controls_running(True)
        return cancel_event

    def cancel_current_search(self) -> None:
        """响应取消按钮，请求后台线程尽快停止。"""

        if self.search_cancel_event is None:
            return

        self.search_cancel_event.set()
        if self.cancel_search_button is not None:
            self.cancel_search_button.configure(state="disabled", text="取消中...")
        self.status_label.configure(text="正在取消检索，请稍等")

    def _search_cancelled(self) -> None:
        """后台线程确认取消后恢复界面状态。"""

        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="已取消检索")

    def _scan_failed(self, message: str) -> None:
        """显示工作表检索失败信息。"""

        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="检索失败")
        messagebox.showerror("检索失败", message)

    def _scan_finished(self, results: list[ExcelWorkbookInfo]) -> None:
        """渲染工作表检索结果并更新状态栏统计。"""

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
        """显示单元格检索失败信息。"""

        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="单元格搜索失败")
        messagebox.showerror("单元格搜索失败", message)

    def _cell_search_finished(self, matches: list[ExcelCellMatch], keyword: str) -> None:
        """渲染单元格检索结果并更新状态栏统计。"""

        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_cell_matches(matches)
        self.status_label.configure(text=f"单元格搜索完成：找到 {len(matches)} 条匹配")

    def _file_search_failed(self, message: str) -> None:
        """显示文件检索失败信息。"""

        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="文件检索失败")
        messagebox.showerror("文件检索失败", message)

    def _file_search_finished(self, results: list[FileSearchResult]) -> None:
        """渲染文件检索结果并更新状态栏统计。"""

        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_file_results(results)
        self.status_label.configure(text=f"文件检索完成：找到 {len(results)} 个文件")

    def _set_search_controls_running(self, is_running: bool) -> None:
        """切换检索中/空闲状态下按钮和选项的可用性。"""

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
        """按工作簿分组展示单元格检索结果。"""

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
            # 父节点展示工作簿和命中数量，子节点才绑定具体单元格位置。
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
        """展示普通文件检索结果，双击时打开文件所在目录。"""

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
        """展开或收起有子节点的结果行。"""

        if self.result_tree.get_children(item_id):
            self.result_tree.item(item_id, open=not bool(self.result_tree.item(item_id, "open")))

    def _open_or_toggle_selected_item(self, item_id: str) -> None:
        """双击父节点时展开/收起，双击具体目标时打开。"""

        target = self.item_targets.get(item_id)
        if target is None:
            return
        if target.sheet_name is None and self.result_tree.get_children(item_id):
            self._toggle_tree_item(item_id)
            return
        self._open_item(item_id)

    def _open_item(self, item_id: str) -> None:
        """打开结果项对应的文件、工作表、单元格或所在目录。"""

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
        """清空结果列表、行目标映射和悬浮操作按钮。"""

        self.item_targets.clear()
        self.reference_items_by_source.clear()
        self.reference_searching_items.clear()
        for button in self.action_buttons.values():
            button.destroy()
        self.action_buttons.clear()
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

    def _render_results(self, results: list[ExcelWorkbookInfo]) -> None:
        """展示工作表检索结果，每个工作表一行并附带查找引用按钮。"""

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
                # 工作表结果行可双击打开指定工作表，也可点击按钮查找引用。
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
        """处理单击事件，点击“操作”列时触发引用查找。"""

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
        """处理双击事件，默认打开对应目标。"""

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
        """启动指定工作表的引用关系查找。"""

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

        # 引用查找可能会打开工作簿，所以放到后台避免阻塞界面。
        worker = Thread(target=self._find_references_in_background, args=(item_id, target, config), daemon=True)
        worker.start()

    def _find_references_in_background(
        self,
        item_id: str,
        target: ResultTarget,
        config: ReferenceLookupConfig,
    ) -> None:
        """后台读取引用关系，并切回主线程更新列表。"""

        try:
            references = find_sheet_reference_matches(target.file_path, target.sheet_name or "", config)
        except ExcelReadError as exc:
            self.after(0, lambda: self._references_failed(item_id, str(exc)))
            return

        self.after(0, lambda: self._references_finished(item_id, target.file_path, references, config))

    def _references_failed(self, item_id: str, message: str) -> None:
        """引用查找失败时恢复按钮状态并显示错误。"""

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
        """把查找到的引用表插入到源工作表行下方。"""

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
        """在源工作表行下面插入一条引用结果行。"""

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
        """删除某个源工作表下已经展开过的引用结果。"""

        for item_id in self.reference_items_by_source.pop(source_item_id, []):
            self.item_targets.pop(item_id, None)
            if self.result_tree.exists(item_id):
                self.result_tree.delete(item_id)
        self._schedule_action_button_layout()

    def _set_action_text(self, item_id: str, text: str) -> None:
        """更新 Treeview 数据里的操作列文本，用于控制按钮是否可点击。"""

        if "action" not in self.current_result_columns or not self.result_tree.exists(item_id):
            return

        action_index = self.current_result_columns.index("action")
        values = list(self.result_tree.item(item_id, "values"))
        if action_index >= len(values):
            return
        values[action_index] = text
        self.result_tree.item(item_id, values=values)

    def _get_action_text(self, item_id: str) -> str:
        """读取某行操作列文本。"""

        if "action" not in self.current_result_columns or not self.result_tree.exists(item_id):
            return ""

        action_index = self.current_result_columns.index("action")
        values = list(self.result_tree.item(item_id, "values"))
        if action_index >= len(values):
            return ""
        return str(values[action_index]).strip()

    def _refresh_sheet_statuses(self) -> None:
        """设置修改后刷新当前工作表列表里的启用/未启用状态。"""

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
        """统计工作表检索完成后实际展示的行数。"""

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
        """按关键词和启用状态计算某个工作簿需要展示的工作表。"""

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
        """在勾选“仅检索启用表格”时过滤掉未启用工作表。"""

        if not self.only_enabled_sheets_var.get():
            return sheet_names
        return [
            sheet_name
            for sheet_name in sheet_names
            if get_sheet_enabled_status(sheet_name, self.disabled_sheet_marker) == "启用"
        ]

    def open_selected_item(self) -> None:
        """保留给快捷入口调用；当前界面主要通过双击打开结果。"""

        selected_items = self.result_tree.selection()
        if not selected_items:
            messagebox.showinfo("未选择", "请先在列表中选择一个工作簿或工作表。")
            return
        self._open_or_toggle_selected_item(selected_items[0])

    def _load_cached_inputs(self) -> None:
        """启动时读取上次文件夹、关键词、历史地址和设置。"""

        cache_data = load_cache_data()
        folder_path = cache_data.get("folder_path", "")
        folder_history = cache_data.get("folder_history", [])
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

        if isinstance(folder_history, list):
            self.folder_history = [folder for folder in folder_history if isinstance(folder, str)]
            self._refresh_folder_history_menu()
        if folder_path:
            self.folder_path_var.set(folder_path)
        if keyword:
            self.keyword_entry.insert(0, keyword)

        if folder_path or keyword:
            self.status_label.configure(text="已加载上次选择的地址和输入内容")

    def _save_cache_from_inputs(self) -> None:
        """把当前输入框、历史地址和设置写入缓存文件。"""

        folder_path = self.folder_path_var.get().strip()
        if folder_path and Path(folder_path).is_dir():
            self._remember_folder_path(folder_path)

        save_cache_data(
            {
                "folder_path": folder_path,
                "folder_history": self.folder_history,
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
        """关闭应用前请求取消后台检索并保存当前输入。"""

        if self.search_cancel_event is not None:
            self.search_cancel_event.set()
        self._save_cache_from_inputs()
        self.destroy()


def load_cache_data() -> dict[str, object]:
    """读取并规范化缓存数据，避免坏缓存影响启动。"""

    raw_data = read_cache_file()
    if not raw_data:
        return {}

    folder_path = raw_data.get("folder_path", "")
    keyword = raw_data.get("keyword", "")
    disabled_sheet_marker = raw_data.get("disabled_sheet_marker", DEFAULT_DISABLED_SHEET_MARKER)
    help_url = raw_data.get("help_url", "")
    return {
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
        "only_enabled_sheets": bool(raw_data.get("only_enabled_sheets", False)),
        "only_enabled_data": bool(raw_data.get("only_enabled_data", False)),
    }


def read_cache_file() -> dict[str, object]:
    """读取原始 JSON 缓存文件，读取失败时返回空字典。"""

    try:
        raw_data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}

    if not isinstance(raw_data, dict):
        return {}

    return raw_data


def save_cache_data(cache_data: dict[str, object]) -> None:
    """保存缓存数据，同时保留工作表名称缓存等其他键。"""

    reference_config = cache_data.get("reference_config", DEFAULT_REFERENCE_LOOKUP_CONFIG)
    if not isinstance(reference_config, ReferenceLookupConfig):
        reference_config = DEFAULT_REFERENCE_LOOKUP_CONFIG
    data_filter_config = cache_data.get("data_filter_config", DEFAULT_DATA_FILTER_CONFIG)
    if not isinstance(data_filter_config, DataFilterConfig):
        data_filter_config = DEFAULT_DATA_FILTER_CONFIG

    payload = read_cache_file()
    payload.update({
        "folder_path": str(cache_data.get("folder_path", "")),
        "folder_history": normalize_folder_history(cache_data.get("folder_history")),
        "keyword": str(cache_data.get("keyword", "")),
        "reference_config": reference_config_to_dict(reference_config),
        "disabled_sheet_marker": str(cache_data.get("disabled_sheet_marker", DEFAULT_DISABLED_SHEET_MARKER)),
        "data_filter_config": data_filter_config_to_dict(data_filter_config),
        "help_url": str(cache_data.get("help_url", "")),
        "only_enabled_sheets": bool(cache_data.get("only_enabled_sheets", False)),
        "only_enabled_data": bool(cache_data.get("only_enabled_data", False)),
    })
    # 工作表名称缓存由 worksheet_search 写入同一个文件，这里保存 UI 缓存时不能覆盖它。
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


def get_sheet_enabled_status(sheet_name: str, disabled_marker: str) -> str:
    """根据工作表名称前缀判断启用状态。"""

    if not sheet_name:
        return ""
    marker = disabled_marker.strip()
    return "未启用" if marker and sheet_name.startswith(marker) else "启用"


def normalize_help_url(help_url: str) -> str:
    """补全帮助链接协议，方便用户只输入域名。"""

    url = help_url.strip()
    if not url:
        return ""
    known_schemes = ("http://", "https://", "file://", "mailto:")
    if url.lower().startswith(known_schemes):
        return url
    return f"https://{url}"


def format_cell_value(value: str) -> str:
    """把单元格内容压成单行，避免结果列表行高被换行撑开。"""

    return " ".join(value.splitlines())


def format_reference_label(reference: ExcelSheetReference) -> str:
    """把引用关系格式化成列表里显示的文本。"""

    parts = [f"引用表：{reference.sheet_name}"]
    if reference.reference_field_name:
        parts.append(f"字段：{reference.reference_field_name}")
    if reference.source_field_name:
        parts.append(f"本表字段：{reference.source_field_name}")
    return "，".join(parts)


def open_containing_folder(file_path: Path) -> bool:
    """打开文件所在目录；Windows 下尽量选中文件本身。"""

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
    """打开文件；Windows 下可进一步定位到指定工作表和单元格。"""

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
    """通过 Excel COM 打开指定工作表，并在需要时选中指定单元格。"""

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
    # 优先复用已经打开的 Excel 进程，避免重复启动多个 Excel 窗口。
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
}} catch {{
    $excel = New-Object -ComObject Excel.Application
}}

$excel.Visible = $true
$workbook = $null

foreach ($book in @($excel.Workbooks)) {{
    # 如果目标工作簿已经打开，直接复用它，避免打开第二份同名文件。
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
    # 单元格检索结果双击打开时，把命中单元格选中并滚动到附近。
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
    """把 Python 字符串安全转成 PowerShell 单引号字符串。"""

    return "'" + value.replace("'", "''") + "'"


def set_windows_app_user_model_id() -> None:
    """设置 Windows 任务栏应用 ID，让窗口图标和分组更稳定。"""

    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Nailong.SearchMaster")
    except Exception:
        pass


def run_app() -> None:
    """配置 CustomTkinter 并进入主事件循环。"""

    set_windows_app_user_model_id()
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    app = ExcelEditApp()
    app.mainloop()
