"""设置弹窗界面。

这里集中管理引用规则、未启用标识、启用数据过滤和帮助链接等配置项。
"""

from collections.abc import Callable
from tkinter import filedialog, messagebox
from decimal import Decimal, InvalidOperation
import re
import unicodedata

import customtkinter as ctk

from cell_search import (
    DEFAULT_DATA_FILTER_CONFIG,
    DataFilterConfig,
    validate_data_filter_config,
)
from excel_common import EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS, ExcelReadError
from worksheet_search import ReferenceLookupConfig, validate_reference_lookup_config


class SettingsDialog:
    """应用设置弹窗，负责展示、读取、校验并回传配置。"""

    def __init__(
        self,
        master: ctk.CTk,
        default_disabled_marker: str,
        on_save: Callable[[ReferenceLookupConfig, str, DataFilterConfig, str, str], bool | None],
    ) -> None:
        self.master = master
        self.default_disabled_marker = default_disabled_marker
        self.on_save = on_save
        self.window: ctk.CTkToplevel | None = None
        self.reference_setting_entries: dict[str, ctk.CTkEntry] = {}
        self.data_filter_setting_entries: dict[str, ctk.CTkEntry] = {}
        self.enable_header_filter_var: ctk.BooleanVar | None = None
        self.enable_blank_header_column_filter_var: ctk.BooleanVar | None = None
        self.enable_column_filter_var: ctk.BooleanVar | None = None
        self.enable_disabled_row_filter_var: ctk.BooleanVar | None = None
        self.enable_range_row_filter_var: ctk.BooleanVar | None = None
        self.disabled_marker_entry: ctk.CTkEntry | None = None
        self.help_url_entry: ctk.CTkEntry | None = None
        self.background_image_entry: ctk.CTkEntry | None = None

    def open(
        self,
        reference_config: ReferenceLookupConfig,
        disabled_sheet_marker: str,
        data_filter_config: DataFilterConfig,
        help_url: str,
        background_image_path: str,
    ) -> None:
        """打开设置窗口，并把当前配置填回输入框。"""

        if self.window is None or not self.window.winfo_exists():
            self._build_window()
        self._fill_entries(
            reference_config,
            disabled_sheet_marker,
            data_filter_config,
            help_url,
            background_image_path,
        )

        if self.window is None:
            return
        self.window.lift()
        self.window.focus_force()

    def _build_window(self) -> None:
        """创建设置窗口控件，只在第一次打开或窗口被销毁后重建。"""

        window = ctk.CTkToplevel(self.master)
        window.title("设置")
        window.geometry("660x660")
        window.minsize(580, 540)
        window.grid_columnconfigure(0, weight=1)
        window.grid_rowconfigure(0, weight=1)
        window.protocol("WM_DELETE_WINDOW", self.close)
        window.transient(self.master)

        tab_view = ctk.CTkTabview(window)
        tab_view.grid(row=0, column=0, padx=16, pady=(16, 8), sticky="nsew")
        tab_view.add("引用数据修改")
        tab_view.add("基础配置")
        tab_view.add("未启用标识")
        tab_view.add("关于")

        reference_tab = tab_view.tab("引用数据修改")
        reference_tab.grid_columnconfigure(1, weight=1)

        # 引用配置决定“查找引用”按钮如何从指定行提取目标表名和字段名。
        fields = (
            ("sample_text", "引用样例", "AAA:BBB.CCC"),
            ("reference_row_index", "引用行", "6"),
            ("table_name", "引用表格名称", "BBB"),
            ("field_name", "引用字段", "CCC"),
            ("field_row_index", "引用字段行", "1"),
        )
        self.reference_setting_entries.clear()
        for row_index, (key, label_text, placeholder) in enumerate(fields):
            label = ctk.CTkLabel(reference_tab, text=label_text, width=110, anchor="w")
            label.grid(row=row_index, column=0, padx=(12, 8), pady=8, sticky="w")
            entry = ctk.CTkEntry(reference_tab, placeholder_text=placeholder)
            entry.grid(row=row_index, column=1, padx=(8, 12), pady=8, sticky="ew")
            self.reference_setting_entries[key] = entry

        self.data_filter_setting_entries.clear()

        base_tab = tab_view.tab("基础配置")
        base_tab.grid_columnconfigure(1, weight=1)

        self.enable_header_filter_var = ctk.BooleanVar(value=DEFAULT_DATA_FILTER_CONFIG.enable_header_filter)
        header_filter_checkbox = ctk.CTkCheckBox(
            base_tab,
            text="启用表头长度",
            variable=self.enable_header_filter_var,
            command=self._update_data_filter_entry_states,
        )
        header_filter_checkbox.grid(row=0, column=0, columnspan=2, padx=(12, 8), pady=(14, 6), sticky="w")
        self._add_data_filter_entry(
            base_tab,
            1,
            "header_row_count",
            "表头长度",
            str(DEFAULT_DATA_FILTER_CONFIG.header_row_count),
        )
        self.enable_blank_header_column_filter_var = ctk.BooleanVar(
            value=DEFAULT_DATA_FILTER_CONFIG.enable_blank_header_column_filter
        )
        blank_header_filter_checkbox = ctk.CTkCheckBox(
            base_tab,
            text="屏蔽空白表头",
            variable=self.enable_blank_header_column_filter_var,
            command=self._update_data_filter_entry_states,
        )
        blank_header_filter_checkbox.grid(row=2, column=0, columnspan=2, padx=(12, 8), pady=(6, 14), sticky="w")

        disabled_tab = tab_view.tab("未启用标识")
        disabled_tab.grid_columnconfigure(1, weight=1)

        # 未启用标识既可以作用在工作表名称，也可以作用在数据列/数据行。
        disabled_label = ctk.CTkLabel(disabled_tab, text="表格未启用标识", width=130, anchor="w")
        disabled_label.grid(row=0, column=0, padx=(12, 8), pady=8, sticky="w")
        self.disabled_marker_entry = ctk.CTkEntry(disabled_tab, placeholder_text=self.default_disabled_marker)
        self.disabled_marker_entry.grid(row=0, column=1, padx=(8, 12), pady=8, sticky="ew")

        self.enable_column_filter_var = ctk.BooleanVar(value=DEFAULT_DATA_FILTER_CONFIG.enable_column_filter)
        column_filter_checkbox = ctk.CTkCheckBox(
            disabled_tab,
            text="启用列禁用",
            variable=self.enable_column_filter_var,
            command=self._update_data_filter_entry_states,
        )
        column_filter_checkbox.grid(row=1, column=0, columnspan=2, padx=(12, 8), pady=(12, 4), sticky="w")

        column_fields = (
            ("column_marker_row_index", "列判断行", str(DEFAULT_DATA_FILTER_CONFIG.column_marker_row_index)),
            ("disabled_column_marker", "列未启用标识", DEFAULT_DATA_FILTER_CONFIG.disabled_column_marker),
        )
        for row_index, (key, label_text, placeholder) in enumerate(column_fields, start=2):
            self._add_data_filter_entry(disabled_tab, row_index, key, label_text, placeholder)

        self.enable_disabled_row_filter_var = ctk.BooleanVar(
            value=DEFAULT_DATA_FILTER_CONFIG.enable_disabled_row_filter
        )
        disabled_row_filter_checkbox = ctk.CTkCheckBox(
            disabled_tab,
            text="启用行禁用",
            variable=self.enable_disabled_row_filter_var,
            command=self._update_data_filter_entry_states,
        )
        disabled_row_filter_checkbox.grid(row=4, column=0, columnspan=2, padx=(12, 8), pady=(12, 4), sticky="w")

        disabled_row_fields = (
            (
                "disabled_row_marker_column_index",
                "行禁用判断列",
                str(DEFAULT_DATA_FILTER_CONFIG.disabled_row_marker_column_index),
            ),
            ("disabled_row_contains", "行未启用标识", ""),
        )
        for row_index, (key, label_text, placeholder) in enumerate(disabled_row_fields, start=5):
            self._add_data_filter_entry(disabled_tab, row_index, key, label_text, placeholder)

        self.enable_range_row_filter_var = ctk.BooleanVar(value=DEFAULT_DATA_FILTER_CONFIG.enable_range_row_filter)
        range_row_filter_checkbox = ctk.CTkCheckBox(
            disabled_tab,
            text="启用行区间",
            variable=self.enable_range_row_filter_var,
            command=self._update_data_filter_entry_states,
        )
        range_row_filter_checkbox.grid(row=7, column=0, columnspan=2, padx=(12, 8), pady=(12, 4), sticky="w")

        range_row_fields = (
            (
                "range_row_marker_column_index",
                "行区间判断列",
                str(DEFAULT_DATA_FILTER_CONFIG.range_row_marker_column_index),
            ),
            ("range_start_text", "区间起始数据", ""),
            ("range_end_text", "区间终止数据", ""),
        )
        for row_index, (key, label_text, placeholder) in enumerate(range_row_fields, start=8):
            self._add_data_filter_entry(disabled_tab, row_index, key, label_text, placeholder)

        self._update_data_filter_entry_states()

        about_tab = tab_view.tab("关于")
        about_tab.grid_columnconfigure(1, weight=1)

        help_label = ctk.CTkLabel(about_tab, text="使用说明", width=110, anchor="w")
        help_label.grid(row=0, column=0, padx=(12, 8), pady=14, sticky="w")
        self.help_url_entry = ctk.CTkEntry(about_tab, placeholder_text="https://example.com")
        self.help_url_entry.grid(row=0, column=1, padx=(8, 12), pady=14, sticky="ew")

        background_image_label = ctk.CTkLabel(about_tab, text="软件背景图", width=110, anchor="w")
        background_image_label.grid(row=1, column=0, padx=(12, 8), pady=8, sticky="w")
        self.background_image_entry = ctk.CTkEntry(about_tab, placeholder_text="未选择背景图片")
        self.background_image_entry.grid(row=1, column=1, padx=(8, 8), pady=8, sticky="ew")
        choose_background_button = ctk.CTkButton(
            about_tab,
            text="选择图片",
            width=90,
            command=self._choose_background_image,
        )
        choose_background_button.grid(row=1, column=2, padx=(0, 8), pady=8)
        clear_background_button = ctk.CTkButton(
            about_tab,
            text="清除",
            width=70,
            fg_color="#64748b",
            hover_color="#475569",
            command=self._clear_background_image,
        )
        clear_background_button.grid(row=1, column=3, padx=(0, 12), pady=8)

        button_frame = ctk.CTkFrame(window, fg_color="transparent")
        button_frame.grid(row=1, column=0, padx=16, pady=(4, 16), sticky="e")

        return_button = ctk.CTkButton(
            button_frame,
            text="返回",
            width=90,
            fg_color="#64748b",
            hover_color="#475569",
            command=self.close,
        )
        return_button.grid(row=0, column=0, padx=8)

        save_button = ctk.CTkButton(
            button_frame,
            text="保存修改",
            width=90,
            command=self._save_settings,
        )
        save_button.grid(row=0, column=1, padx=(8, 0))

        self.window = window

    def _add_data_filter_entry(
        self,
        parent: ctk.CTkFrame,
        row_index: int,
        key: str,
        label_text: str,
        placeholder: str,
    ) -> None:
        """添加一行启用数据过滤配置输入框，并按 key 保存引用。"""

        label = ctk.CTkLabel(parent, text=label_text, width=130, anchor="w")
        label.grid(row=row_index, column=0, padx=(12, 8), pady=6, sticky="w")
        entry = ctk.CTkEntry(parent, placeholder_text=placeholder)
        entry.grid(row=row_index, column=1, padx=(8, 12), pady=6, sticky="ew")
        self.data_filter_setting_entries[key] = entry

    def _update_data_filter_entry_states(self) -> None:
        """根据复选框开关启用或禁用对应输入框。"""

        header_state = (
            "normal"
            if self._get_bool_var_value(self.enable_header_filter_var)
            or self._get_bool_var_value(self.enable_blank_header_column_filter_var)
            else "disabled"
        )
        self._set_data_filter_entry_state("header_row_count", header_state)

        column_state = "normal" if self._get_bool_var_value(self.enable_column_filter_var) else "disabled"
        for key in ("column_marker_row_index", "disabled_column_marker"):
            self._set_data_filter_entry_state(key, column_state)

        disabled_row_state = "normal" if self._get_bool_var_value(self.enable_disabled_row_filter_var) else "disabled"
        for key in ("disabled_row_marker_column_index", "disabled_row_contains"):
            self._set_data_filter_entry_state(key, disabled_row_state)

        range_row_state = "normal" if self._get_bool_var_value(self.enable_range_row_filter_var) else "disabled"
        for key in ("range_row_marker_column_index", "range_start_text", "range_end_text"):
            self._set_data_filter_entry_state(key, range_row_state)

    def _set_data_filter_entry_state(self, key: str, state: str) -> None:
        """设置单个过滤输入框状态；不存在时静默跳过。"""

        entry = self.data_filter_setting_entries.get(key)
        if entry is not None:
            entry.configure(state=state)

    def _get_bool_var_value(self, variable: ctk.BooleanVar | None) -> bool:
        """安全读取 CustomTkinter 布尔变量。"""

        return bool(variable.get()) if variable is not None else False

    def _fill_entries(
        self,
        reference_config: ReferenceLookupConfig,
        disabled_sheet_marker: str,
        data_filter_config: DataFilterConfig,
        help_url: str,
        background_image_path: str,
    ) -> None:
        """把当前配置值写入窗口输入框。"""

        values = {
            "sample_text": reference_config.sample_text,
            "reference_row_index": str(reference_config.reference_row_index),
            "table_name": reference_config.table_name,
            "field_name": reference_config.field_name,
            "field_row_index": str(reference_config.field_row_index),
        }
        for key, value in values.items():
            entry = self.reference_setting_entries.get(key)
            if entry is None:
                continue
            entry.delete(0, "end")
            entry.insert(0, value)

        if self.disabled_marker_entry is not None:
            self.disabled_marker_entry.delete(0, "end")
            self.disabled_marker_entry.insert(0, disabled_sheet_marker)

        data_filter_values = {
            "header_row_count": str(data_filter_config.header_row_count),
            "column_marker_row_index": str(data_filter_config.column_marker_row_index),
            "disabled_column_marker": data_filter_config.disabled_column_marker,
            "disabled_row_marker_column_index": str(data_filter_config.disabled_row_marker_column_index),
            "disabled_row_contains": data_filter_config.disabled_row_contains,
            "range_row_marker_column_index": str(data_filter_config.range_row_marker_column_index),
            "range_start_text": data_filter_config.range_start_text,
            "range_end_text": data_filter_config.range_end_text,
        }
        for key, value in data_filter_values.items():
            entry = self.data_filter_setting_entries.get(key)
            if entry is None:
                continue
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.insert(0, value)

        if self.enable_header_filter_var is not None:
            self.enable_header_filter_var.set(data_filter_config.enable_header_filter)
        if self.enable_blank_header_column_filter_var is not None:
            self.enable_blank_header_column_filter_var.set(data_filter_config.enable_blank_header_column_filter)
        if self.enable_column_filter_var is not None:
            self.enable_column_filter_var.set(data_filter_config.enable_column_filter)
        if self.enable_disabled_row_filter_var is not None:
            self.enable_disabled_row_filter_var.set(data_filter_config.enable_disabled_row_filter)
        if self.enable_range_row_filter_var is not None:
            self.enable_range_row_filter_var.set(data_filter_config.enable_range_row_filter)
        self._update_data_filter_entry_states()

        if self.help_url_entry is not None:
            self.help_url_entry.delete(0, "end")
            self.help_url_entry.insert(0, help_url)
        self._set_background_image_entry(background_image_path)

    def _save_settings(self) -> None:
        """读取并校验所有设置，成功后通过回调交给主窗口保存。"""

        try:
            data_filter_config = self._read_data_filter_settings_entries()
            validate_data_filter_config(data_filter_config)
            config = self._read_reference_settings_entries()
            validate_reference_lookup_config(config)
        except ExcelReadError as exc:
            messagebox.showerror("设置无效", str(exc), parent=self.window)
            return

        save_succeeded = self.on_save(
            config,
            self._read_disabled_marker_entry(),
            data_filter_config,
            self._read_help_url_entry(),
            self._read_background_image_entry(),
        )
        if save_succeeded is False:
            return
        self.close()

    def _read_reference_settings_entries(self) -> ReferenceLookupConfig:
        """从引用配置页签读取 ReferenceLookupConfig。"""

        def get_entry_value(key: str) -> str:
            entry = self.reference_setting_entries.get(key)
            return entry.get().strip() if entry is not None else ""

        return ReferenceLookupConfig(
            sample_text=get_entry_value("sample_text"),
            reference_row_index=parse_positive_int(
                get_entry_value("reference_row_index"),
                "引用行",
                EXCEL_MAX_ROWS,
            ),
            table_name=get_entry_value("table_name"),
            field_name=get_entry_value("field_name"),
            field_row_index=parse_positive_int(
                get_entry_value("field_row_index"),
                "引用字段行",
                EXCEL_MAX_ROWS,
            ),
        )

    def _read_data_filter_settings_entries(self) -> DataFilterConfig:
        """从基础配置和未启用标识页签读取 DataFilterConfig。"""

        def get_entry_value(key: str) -> str:
            entry = self.data_filter_setting_entries.get(key)
            return entry.get().strip() if entry is not None else ""

        enable_column_filter = self._get_bool_var_value(self.enable_column_filter_var)
        enable_disabled_row_filter = self._get_bool_var_value(self.enable_disabled_row_filter_var)
        enable_header_filter = self._get_bool_var_value(self.enable_header_filter_var)
        enable_blank_header_column_filter = self._get_bool_var_value(self.enable_blank_header_column_filter_var)
        enable_range_row_filter = self._get_bool_var_value(self.enable_range_row_filter_var)

        return DataFilterConfig(
            enable_header_filter=enable_header_filter,
            header_row_count=(
                parse_positive_int(get_entry_value("header_row_count"), "表头长度", EXCEL_MAX_ROWS)
                if enable_header_filter or enable_blank_header_column_filter
                else DEFAULT_DATA_FILTER_CONFIG.header_row_count
            ),
            enable_blank_header_column_filter=enable_blank_header_column_filter,
            enable_column_filter=enable_column_filter,
            column_marker_row_index=(
                parse_positive_int(get_entry_value("column_marker_row_index"), "列判断行", EXCEL_MAX_ROWS)
                if enable_column_filter
                else DEFAULT_DATA_FILTER_CONFIG.column_marker_row_index
            ),
            disabled_column_marker=get_entry_value("disabled_column_marker"),
            enable_disabled_row_filter=enable_disabled_row_filter,
            disabled_row_marker_column_index=(
                parse_positive_int(
                    get_entry_value("disabled_row_marker_column_index"),
                    "行禁用判断列",
                    EXCEL_MAX_COLUMNS,
                )
                if enable_disabled_row_filter
                else DEFAULT_DATA_FILTER_CONFIG.disabled_row_marker_column_index
            ),
            disabled_row_contains=get_entry_value("disabled_row_contains"),
            enable_range_row_filter=enable_range_row_filter,
            range_row_marker_column_index=(
                parse_positive_int(
                    get_entry_value("range_row_marker_column_index"),
                    "行区间判断列",
                    EXCEL_MAX_COLUMNS,
                )
                if enable_range_row_filter
                else DEFAULT_DATA_FILTER_CONFIG.range_row_marker_column_index
            ),
            range_start_text=get_entry_value("range_start_text"),
            range_end_text=get_entry_value("range_end_text"),
        )

    def _read_disabled_marker_entry(self) -> str:
        """读取工作表未启用前缀标识。"""

        if self.disabled_marker_entry is None:
            return ""
        return self.disabled_marker_entry.get().strip()

    def _read_help_url_entry(self) -> str:
        """读取帮助链接。"""

        if self.help_url_entry is None:
            return ""
        return self.help_url_entry.get().strip()

    def _choose_background_image(self) -> None:
        """使用 Windows 文件选择窗口挑选常见图片格式。"""

        image_path = filedialog.askopenfilename(
            parent=self.window,
            title="选择软件背景图",
            filetypes=[
                ("所有图片文件", "*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp;*.tif;*.tiff;*.ico;*.avif"),
                ("所有文件", "*.*"),
            ],
        )
        if image_path:
            self._set_background_image_entry(image_path)

    def _clear_background_image(self) -> None:
        """清空待保存的背景图路径，恢复默认背景。"""

        self._set_background_image_entry("")

    def _set_background_image_entry(self, image_path: str) -> None:
        """把背景图路径写入输入框。"""

        if self.background_image_entry is None:
            return
        self.background_image_entry.delete(0, "end")
        self.background_image_entry.insert(0, image_path)

    def _read_background_image_entry(self) -> str:
        """读取背景图路径。"""

        if self.background_image_entry is None:
            return ""
        return self.background_image_entry.get().strip()

    def close(self) -> None:
        """关闭弹窗并清空控件引用，避免下次复用失效对象。"""

        if self.window is not None and self.window.winfo_exists():
            self.window.destroy()
        self.window = None
        self.reference_setting_entries.clear()
        self.data_filter_setting_entries.clear()
        self.enable_header_filter_var = None
        self.enable_blank_header_column_filter_var = None
        self.enable_column_filter_var = None
        self.enable_disabled_row_filter_var = None
        self.enable_range_row_filter_var = None
        self.disabled_marker_entry = None
        self.help_url_entry = None
        self.background_image_entry = None


class SearchSettingsDialog:
    """检索筛选弹窗，按分类集中管理全部检索条件。"""

    def __init__(
        self,
        master: ctk.CTk,
        on_save: Callable[[bool, bool, bool, str, bool, bool], None],
    ) -> None:
        self.master = master
        self.on_save = on_save
        self.window: ctk.CTkToplevel | None = None
        self.search_mode = "工作表检索"
        self.settings_frame: ctk.CTkScrollableFrame | None = None
        self.file_suffixes_entry: ctk.CTkEntry | None = None
        self.only_enabled_sheets_var = ctk.BooleanVar(value=False)
        self.only_enabled_data_var = ctk.BooleanVar(value=False)
        self.only_specified_suffixes_var = ctk.BooleanVar(value=False)
        self.file_suffixes_var = ctk.StringVar(value="")
        self.search_folders_var = ctk.BooleanVar(value=False)
        self.exact_match_var = ctk.BooleanVar(value=False)

    def open(
        self,
        search_mode: str,
        only_enabled_sheets: bool,
        only_enabled_data: bool,
        only_specified_suffixes: bool,
        file_suffixes: str,
        search_folders: bool,
        exact_match: bool,
    ) -> None:
        """打开检索设置，并把主界面当前选项复制到临时表单。"""

        self.search_mode = search_mode
        self.only_enabled_sheets_var.set(only_enabled_sheets)
        self.only_enabled_data_var.set(only_enabled_data)
        self.only_specified_suffixes_var.set(only_specified_suffixes)
        self.file_suffixes_var.set(file_suffixes)
        self.search_folders_var.set(search_folders)
        self.exact_match_var.set(exact_match)

        if self.window is None or not self.window.winfo_exists():
            self._build_window()
        self._update_file_suffix_entry_state()

        if self.window is not None:
            self.window.lift()
            self.window.focus_force()

    def refresh(self, search_mode: str) -> None:
        """记录当前检索类型，供下次打开检索设置时使用。"""

        self.search_mode = search_mode

    def _build_window(self) -> None:
        """创建检索设置窗口。"""

        window = ctk.CTkToplevel(self.master)
        window.title("检索设置")
        window.geometry("540x480")
        window.minsize(480, 400)
        window.grid_columnconfigure(0, weight=1)
        window.grid_rowconfigure(0, weight=1)
        window.protocol("WM_DELETE_WINDOW", self.close)
        window.transient(self.master)

        settings_frame = ctk.CTkScrollableFrame(window)
        settings_frame.grid(row=0, column=0, padx=16, pady=(16, 8), sticky="nsew")
        settings_frame.grid_columnconfigure(1, weight=1)

        self._add_category_title(settings_frame, 0, "通用设置")
        ctk.CTkCheckBox(
            settings_frame,
            text="精确查找",
            variable=self.exact_match_var,
        ).grid(row=1, column=0, columnspan=2, padx=16, pady=(0, 14), sticky="w")

        self._add_category_title(settings_frame, 2, "工作表检索设置")
        ctk.CTkCheckBox(
            settings_frame,
            text="仅检索启用表格（也应用于单元格检索）",
            variable=self.only_enabled_sheets_var,
        ).grid(row=3, column=0, columnspan=2, padx=16, pady=(0, 14), sticky="w")

        self._add_category_title(settings_frame, 4, "单元格检索设置")
        ctk.CTkCheckBox(
            settings_frame,
            text="仅检索启用数据",
            variable=self.only_enabled_data_var,
        ).grid(row=5, column=0, columnspan=2, padx=16, pady=(0, 14), sticky="w")

        self._add_category_title(settings_frame, 6, "文件检索设置")
        ctk.CTkCheckBox(
            settings_frame,
            text="检索文件夹",
            variable=self.search_folders_var,
        ).grid(row=7, column=0, columnspan=2, padx=16, pady=(0, 6), sticky="w")
        ctk.CTkCheckBox(
            settings_frame,
            text="仅检索指定后缀文件",
            variable=self.only_specified_suffixes_var,
            command=self._update_file_suffix_entry_state,
        ).grid(row=8, column=0, columnspan=2, padx=16, pady=6, sticky="w")
        suffix_label = ctk.CTkLabel(settings_frame, text="文件后缀", width=90, anchor="w")
        suffix_label.grid(row=9, column=0, padx=(16, 8), pady=(6, 16), sticky="w")
        self.file_suffixes_entry = ctk.CTkEntry(
            settings_frame,
            textvariable=self.file_suffixes_var,
            placeholder_text="例如：xlsx, pdf, docx",
        )
        self.file_suffixes_entry.grid(row=9, column=1, padx=(8, 16), pady=(6, 16), sticky="ew")

        button_frame = ctk.CTkFrame(window, fg_color="transparent")
        button_frame.grid(row=1, column=0, padx=16, pady=(4, 16), sticky="e")

        cancel_button = ctk.CTkButton(
            button_frame,
            text="取消",
            width=90,
            fg_color="#64748b",
            hover_color="#475569",
            command=self.close,
        )
        cancel_button.grid(row=0, column=0, padx=8)

        save_button = ctk.CTkButton(button_frame, text="保存", width=90, command=self._save)
        save_button.grid(row=0, column=1, padx=(8, 0))

        self.window = window
        self.settings_frame = settings_frame

    def _add_category_title(self, parent: ctk.CTkScrollableFrame, row_index: int, text: str) -> None:
        """在检索设置内添加分类标题。"""

        ctk.CTkLabel(
            parent,
            text=text,
            font=("Microsoft YaHei UI", 14, "bold"),
            anchor="w",
        ).grid(row=row_index, column=0, columnspan=2, padx=16, pady=(14, 8), sticky="w")

    def _update_file_suffix_entry_state(self) -> None:
        """只有选中指定后缀条件时，才能编辑后缀文本。"""

        if self.file_suffixes_entry is None:
            return
        state = "normal" if self.only_specified_suffixes_var.get() else "disabled"
        self.file_suffixes_entry.configure(state=state)

    def _save(self) -> None:
        """保存临时筛选项并关闭窗口。"""

        self.on_save(
            self.only_enabled_sheets_var.get(),
            self.only_enabled_data_var.get(),
            self.only_specified_suffixes_var.get(),
            self.file_suffixes_var.get().strip(),
            self.search_folders_var.get(),
            self.exact_match_var.get(),
        )
        self.close()

    def close(self) -> None:
        """关闭检索设置窗口。"""

        if self.window is not None and self.window.winfo_exists():
            self.window.destroy()
        self.window = None
        self.settings_frame = None
        self.file_suffixes_entry = None


def parse_positive_int(value: object, field_label: str, maximum: int | None = None) -> int:
    """把用户输入解析为正整数，并给出带字段名的错误提示。"""

    normalized_value = normalize_number_text(value)
    try:
        decimal_value = Decimal(normalized_value)
    except (InvalidOperation, ValueError):
        decorated_match = re.fullmatch(r"(?:第)?([+-]?\d+)(?:行|列)?", normalized_value)
        if decorated_match is None:
            current_value = "空" if not normalized_value else normalized_value
            raise ExcelReadError(f"{field_label}必须是正整数，当前读取到：{current_value}")
        decimal_value = Decimal(decorated_match.group(1))

    if not decimal_value.is_finite() or decimal_value != decimal_value.to_integral_value():
        raise ExcelReadError(f"{field_label}必须是正整数，当前读取到：{normalized_value}")
    if decimal_value < 1:
        raise ExcelReadError(f"{field_label}必须大于等于 1")
    if maximum is not None and decimal_value > maximum:
        raise ExcelReadError(f"{field_label}不能大于 {maximum}")
    return int(decimal_value)


def normalize_number_text(value: object) -> str:
    """统一全角数字、逗号和不可见字符，降低用户输入格式要求。"""

    text = unicodedata.normalize("NFKC", str(value)).strip()
    text = text.replace(",", "").replace("，", "")
    return "".join(char for char in text if not char.isspace() and char not in {"\u200b", "\ufeff"})
