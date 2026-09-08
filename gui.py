"""应用主界面和用户交互逻辑。

这个文件负责窗口布局、缓存输入、启动后台检索线程、切换表格或缩略图结果视图、
打开文件/工作表/单元格，以及在工作表结果中查找引用关系。
"""

import os
import subprocess
import sys
import tempfile
import webbrowser
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from queue import Empty, Queue
from threading import Event, Thread
from tkinter import Frame as TkFrame, filedialog, font as tkfont, messagebox, ttk

import customtkinter as ctk
from PIL import Image, ImageDraw, ImageOps, UnidentifiedImageError

from cache_store import get_user_cache_directory
from cache_store import read_cache_data as read_shared_cache_data
from cache_store import update_cache_data as update_shared_cache_data
from cell_search import (
    DEFAULT_MAX_SEARCH_RESULTS,
    DEFAULT_DATA_FILTER_CONFIG,
    DataFilterConfig,
    ExcelCellMatch,
    ExcelCellSearchIssue,
    search_excel_cells,
    validate_data_filter_config,
)
from excel_common import EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS, ExcelReadError
from file_search import (
    DEFAULT_MAX_FILE_SEARCH_RESULTS,
    FileSearchError,
    FileSearchResult,
    normalize_file_suffixes,
    search_files,
)
from popup_gui import SearchSettingsDialog, SettingsDialog, parse_positive_int
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


APP_NAME = "奶龙检索大师"
APP_ICON_RELATIVE_PATH = "assets/nailong_search_master.ico"
DEFAULT_DISABLED_SHEET_MARKER = "$"
FOLDER_HISTORY_LIMIT = 20
EXCEL_OPEN_TIMEOUT_SECONDS = 20
WORKSHEET_RESULT_COLUMNS = (
    ("workbook", "工作簿", 360, 220, True),
    ("sheet", "工作表", 360, 220, True),
    ("status", "启用状态", 150, 110, False),
    ("action", "操作", 132, 120, False),
)
CELL_RESULT_COLUMNS = (
    ("value", "数据内容", 620, 260, True),
    ("action", "操作", 132, 120, False),
)
FILE_TABLE_RESULT_COLUMNS = (
    ("file_name", "文件/文件夹名", 320, 180, True),
    ("item_type", "类型", 150, 120, False),
    ("file_address", "所在位置", 620, 320, True),
)
FILE_THUMBNAIL_COLUMNS = 4
FILE_THUMBNAIL_CARD_WIDTH = 210
FILE_THUMBNAIL_CARD_HEIGHT = 176
IMAGE_FILE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
FILE_THUMBNAIL_IMAGE_SIZE = (64, 64)
THUMBNAIL_BACKGROUND_COLOR = "#ffffff"
THUMBNAIL_TEXT_COLOR = "#1f2937"
THUMBNAIL_SECONDARY_TEXT_COLOR = "#64748b"
BACKGROUND_IMAGE_FILE_NAME = "software_background.png"
BACKGROUND_RENDER_MAX_SIZE = (2560, 1440)
TRANSLUCENT_STRIP_OPACITY = 0.85
TRANSLUCENT_STRIP_COLOR = (238, 241, 245)


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
        self.directory_buttons: dict[str, ctk.CTkButton] = {}
        self.action_layout_job: str | None = None
        self.current_result_columns: tuple[str, ...] = ()
        self.file_thumbnail_results: list[FileSearchResult] = []
        self.file_thumbnail_cards: list[ctk.CTkFrame] = []
        self.file_thumbnail_images: list[ctk.CTkImage] = []
        self.background_image_path = ""
        self.background_source_image: Image.Image | None = None
        self.background_fitted_image: Image.Image | None = None
        self.background_display_image: ctk.CTkImage | None = None
        self.background_image_label: ctk.CTkLabel | None = None
        self.background_render_size: tuple[int, int] | None = None
        self.background_layout_job: str | None = None
        self.translucent_strip_frames: dict[str, ctk.CTkFrame] = {}
        self.translucent_strip_overlays: dict[str, ctk.CTkLabel] = {}
        self.translucent_strip_images: dict[str, ctk.CTkImage] = {}
        self.translucent_strip_layout_keys: dict[str, tuple[int, int, int, int, int]] = {}
        self.translucent_strip_layout_job: str | None = None
        self.file_result_view_var = ctk.StringVar(value="缩略图")
        self.result_grid_lines: list[TkFrame] = []
        self.result_grid_layout_job: str | None = None
        self.folder_path_var = ctk.StringVar(value="")
        self.folder_history: list[str] = []
        self.search_mode_var = ctk.StringVar(value="工作表检索")
        self.only_enabled_sheets_var = ctk.BooleanVar(value=False)
        self.only_enabled_data_var = ctk.BooleanVar(value=False)
        self.only_specified_suffixes_var = ctk.BooleanVar(value=False)
        self.file_suffixes_var = ctk.StringVar(value="")
        self.search_folders_var = ctk.BooleanVar(value=False)
        self.exact_match_var = ctk.BooleanVar(value=False)
        self.current_keyword = ""
        self.current_exact_match = False
        self.reference_config = DEFAULT_REFERENCE_LOOKUP_CONFIG
        self.disabled_sheet_marker = DEFAULT_DISABLED_SHEET_MARKER
        self.data_filter_config = DEFAULT_DATA_FILTER_CONFIG
        self.help_url = ""
        self.settings_dialog = SettingsDialog(self, DEFAULT_DISABLED_SHEET_MARKER, self._apply_settings)
        self.search_settings_dialog = SearchSettingsDialog(self, self._apply_search_settings)
        self.settings_button: ctk.CTkButton | None = None
        self.help_button: ctk.CTkButton | None = None
        self.search_settings_button: ctk.CTkButton | None = None
        self.cancel_search_button: ctk.CTkButton | None = None
        self.search_cancel_event: Event | None = None
        self.reference_cancel_events: dict[str, Event] = {}
        self.result_generation = 0
        self.opening_targets: set[ResultTarget] = set()
        self.is_closing = False
        self.ui_callbacks: Queue[Callable[[], None]] = Queue()
        self.ui_poll_job: str | None = None

        # 先构建控件再加载缓存，因为缓存需要写回输入框和下拉框。
        self._build_layout()
        self._load_cached_inputs()
        self.protocol("WM_DELETE_WINDOW", self.close_app)
        self.ui_poll_job = self.after(50, self._poll_ui_callbacks)

    def _post_to_ui(self, callback: Callable[[], None]) -> None:
        """后台线程只写入队列，由 Tk 主线程统一执行界面更新。"""

        if not self.is_closing:
            self.ui_callbacks.put(callback)

    def _poll_ui_callbacks(self) -> None:
        """在 Tk 主线程轮询后台任务结果，避免跨线程直接调用 Tk。"""

        self.ui_poll_job = None
        if self.is_closing:
            return

        while True:
            try:
                callback = self.ui_callbacks.get_nowait()
            except Empty:
                break
            try:
                callback()
            except Exception as exc:
                self.report_callback_exception(type(exc), exc, exc.__traceback__)
            if self.is_closing:
                return

        self.ui_poll_job = self.after(50, self._poll_ui_callbacks)

    def _post_search_callback(self, cancel_event: Event, callback: Callable[[], None]) -> None:
        """只接收当前检索任务的结果，丢弃已经取消或过期的回调。"""

        def run_if_current() -> None:
            if self.is_closing or self.search_cancel_event is not cancel_event:
                return
            callback()

        self._post_to_ui(run_if_current)

    def _set_app_icon(self) -> None:
        """设置窗口图标；缺少图标时不影响程序启动。"""

        icon_path = get_resource_path(APP_ICON_RELATIVE_PATH)
        if not icon_path.exists():
            return
        try:
            self.iconbitmap(default=str(icon_path))
        except Exception:
            pass

    def _set_background_image_path(self, image_path: str, show_error: bool = True) -> bool:
        """加载背景图；图片不可读取时保持当前背景不变。"""

        normalized_path = image_path.strip()
        source_image: Image.Image | None = None
        if normalized_path:
            path = Path(normalized_path)
            if not path.is_file():
                if show_error:
                    messagebox.showerror("背景图无效", "选择的背景图片不存在或无法读取。", parent=self)
                return False
            try:
                with Image.open(path) as image:
                    image.load()
                    source_image = image.convert("RGBA")
                    source_image.thumbnail(BACKGROUND_RENDER_MAX_SIZE, Image.Resampling.LANCZOS)
            except (OSError, UnidentifiedImageError) as exc:
                if show_error:
                    messagebox.showerror("背景图无效", f"无法加载所选图片：{exc}", parent=self)
                return False

        self.background_image_path = normalized_path
        self.background_source_image = source_image
        self.background_fitted_image = None
        self.background_display_image = None
        self.background_render_size = None
        self.translucent_strip_layout_keys.clear()
        self._schedule_background_image_layout()
        return True

    def _save_background_image(self, selected_image_path: str) -> bool:
        """把用户选择的图片转存到软件数据目录，避免原文件移动后背景失效。"""

        normalized_path = selected_image_path.strip()
        if not normalized_path:
            return self._set_background_image_path("")

        source_path = Path(normalized_path)
        if not source_path.is_file():
            messagebox.showerror("背景图无效", "选择的背景图片不存在或无法读取。", parent=self)
            return False

        temporary_path: Path | None = None
        try:
            with Image.open(source_path) as image:
                image.load()
                saved_image = image.convert("RGBA")

            destination_path = get_user_cache_directory() / BACKGROUND_IMAGE_FILE_NAME
            destination_path.parent.mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile(
                dir=destination_path.parent,
                prefix=".software_background.",
                suffix=".tmp",
                delete=False,
            ) as temporary_file:
                temporary_path = Path(temporary_file.name)
            saved_image.save(temporary_path, format="PNG")
            os.replace(temporary_path, destination_path)
        except (OSError, UnidentifiedImageError) as exc:
            messagebox.showerror("背景图保存失败", f"无法保存所选图片：{exc}", parent=self)
            return False
        finally:
            if temporary_path is not None and temporary_path.exists():
                try:
                    temporary_path.unlink()
                except OSError:
                    pass

        return self._set_background_image_path(str(destination_path), show_error=False)

    def _schedule_background_image_layout(self) -> None:
        """合并窗口尺寸变化事件，避免连续拖动窗口时重复缩放图片。"""

        if self.background_layout_job is not None:
            return
        self.background_layout_job = self.after(120, self._layout_background_image)

    def _on_app_configure(self, _event: object) -> None:
        """窗口尺寸变化时，同步背景图和半透明条的尺寸。"""

        self._schedule_background_image_layout()
        self._schedule_translucent_strip_layout()

    def _layout_background_image(self) -> None:
        """将背景图按窗口尺寸等比裁切填充。"""

        self.background_layout_job = None
        if self.background_image_label is None:
            return
        if self.background_source_image is None:
            self.background_fitted_image = None
            self.background_render_size = None
            self.background_image_label.configure(image=None)
            self._schedule_translucent_strip_layout()
            return

        width = self.winfo_width()
        height = self.winfo_height()
        if width < 10 or height < 10:
            self.background_layout_job = self.after(50, self._layout_background_image)
            return
        if self.background_render_size == (width, height):
            return

        try:
            fitted_image = ImageOps.fit(
                self.background_source_image,
                (width, height),
                method=Image.Resampling.LANCZOS,
            )
        except OSError:
            return
        self.background_fitted_image = fitted_image
        self.background_render_size = (width, height)
        self.background_display_image = ctk.CTkImage(
            light_image=fitted_image,
            dark_image=fitted_image,
            size=(width, height),
        )
        self.background_image_label.configure(image=self.background_display_image)
        self.background_image_label.lower()
        self._schedule_translucent_strip_layout()

    def _register_translucent_strip(self, name: str, frame: ctk.CTkFrame) -> None:
        """注册需要叠加 85% 浅色遮罩的横向区域。"""

        # 遮罩必须是条框内部的底层控件，才能盖住 CTkFrame 自身的背景，
        # 同时让地址框、按钮等后续子控件仍显示在它的上方。
        overlay = ctk.CTkLabel(frame, text="", fg_color="transparent", corner_radius=10)
        self.translucent_strip_frames[name] = frame
        self.translucent_strip_overlays[name] = overlay
        overlay.lower()
        self._schedule_translucent_strip_layout()

    def _schedule_translucent_strip_layout(self) -> None:
        """合并半透明条的重绘请求。"""

        if self.translucent_strip_layout_job is not None:
            return
        self.translucent_strip_layout_job = self.after(120, self._layout_translucent_strips)

    def _layout_translucent_strips(self) -> None:
        """以背景图为底绘制 85% 不透明的浅色条。"""

        self.translucent_strip_layout_job = None
        opacity = round(255 * TRANSLUCENT_STRIP_OPACITY)
        for name, frame in self.translucent_strip_frames.items():
            overlay = self.translucent_strip_overlays.get(name)
            if overlay is None:
                continue
            x, y = frame.winfo_x(), frame.winfo_y()
            width, height = frame.winfo_width(), frame.winfo_height()
            if width < 2 or height < 2:
                overlay.place_forget()
                continue

            overlay.configure(width=width, height=height)
            overlay.place(x=0, y=0)
            layout_key = (x, y, width, height, id(self.background_fitted_image))
            if self.translucent_strip_layout_keys.get(name) == layout_key:
                continue
            if self.background_fitted_image is None:
                overlay.configure(image=None, fg_color="#eef1f5")
                self.translucent_strip_layout_keys[name] = layout_key
                continue

            background_crop = self.background_fitted_image.crop((x, y, x + width, y + height)).convert("RGBA")
            tint = Image.new("RGBA", background_crop.size, (*TRANSLUCENT_STRIP_COLOR, opacity))
            overlay_image = Image.alpha_composite(background_crop, tint)
            corner_mask = Image.new("L", overlay_image.size, 0)
            ImageDraw.Draw(corner_mask).rounded_rectangle(
                (0, 0, width - 1, height - 1),
                radius=min(10, width // 2, height // 2),
                fill=255,
            )
            overlay_image.putalpha(corner_mask)
            display_image = ctk.CTkImage(
                light_image=overlay_image,
                dark_image=overlay_image,
                size=(width, height),
            )
            self.translucent_strip_images[name] = display_image
            overlay.configure(image=display_image, fg_color="transparent")
            overlay.lower()
            self.translucent_strip_layout_keys[name] = layout_key

    def _build_layout(self) -> None:
        """创建主界面的所有控件和布局。"""

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        self.background_image_label = ctk.CTkLabel(self, text="", fg_color="transparent")
        self.background_image_label.place(x=0, y=0, relwidth=1, relheight=1)
        self.background_image_label.lower()
        self.bind("<Configure>", self._on_app_configure, add="+")

        # 顶部工具区：设置和帮助入口。
        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=16, pady=(16, 4), sticky="ew")
        top_frame.grid_columnconfigure(3, weight=1)
        self._register_translucent_strip("top", top_frame)

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

        self.search_settings_button = ctk.CTkButton(
            top_frame,
            text="检索设置",
            width=92,
            height=26,
            font=("Microsoft YaHei UI", 12),
            command=self.open_search_settings,
        )
        self.search_settings_button.grid(row=0, column=2, padx=0, pady=0, sticky="w")

        # 文件夹地址和关键词并排展示，减少顶部区域占用的高度。
        search_input_frame = ctk.CTkFrame(self, fg_color="transparent")
        search_input_frame.grid(row=1, column=0, padx=16, pady=(4, 8), sticky="ew")
        search_input_frame.grid_columnconfigure(1, weight=3)
        search_input_frame.grid_columnconfigure(4, weight=2)
        self._register_translucent_strip("search_input", search_input_frame)

        folder_label = ctk.CTkLabel(search_input_frame, text="文件夹地址", width=90, anchor="w")
        folder_label.grid(row=0, column=0, padx=(12, 8), pady=12)

        self.folder_entry = ctk.CTkComboBox(
            search_input_frame,
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
        self.folder_entry.grid(row=0, column=1, padx=(0, 8), pady=12, sticky="ew")
        self.folder_entry.bind("<FocusOut>", lambda _event: self._save_cache_from_inputs())
        # CustomTkinter 的下拉菜单默认偏窄，这里接管打开逻辑来同步历史列表宽度。
        self.folder_entry._open_dropdown_menu = self._open_folder_dropdown_menu
        search_input_frame.bind("<Configure>", lambda _event: self._sync_folder_dropdown_width(), add="+")
        self.after_idle(self._sync_folder_dropdown_width)

        choose_button = ctk.CTkButton(search_input_frame, text="选择文件夹", width=120, command=self.choose_folder)
        choose_button.grid(row=0, column=2, padx=(0, 16), pady=12)

        # 关键词输入框会被不同检索模式复用。
        keyword_label = ctk.CTkLabel(search_input_frame, text="输入内容", width=90, anchor="w")
        keyword_label.grid(row=0, column=3, padx=(0, 8), pady=12)

        self.keyword_entry = ctk.CTkEntry(search_input_frame, placeholder_text="可输入工作簿、工作表、单元格、文件名或文件地址关键词；留空则显示全部")
        self.keyword_entry.grid(row=0, column=4, padx=(0, 12), pady=12, sticky="ew")
        self.keyword_entry.bind("<Return>", lambda _event: self.run_selected_search())
        self.keyword_entry.bind("<FocusOut>", lambda _event: self._save_cache_from_inputs())

        # 检索操作区：模式下拉框、取消按钮、检索按钮和状态栏。
        action_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_frame.grid(row=2, column=0, padx=16, pady=8, sticky="ew")
        action_frame.grid_columnconfigure(1, weight=1)
        self._register_translucent_strip("search_action", action_frame)

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
            command=self._on_search_mode_changed,
        )
        self.search_mode_menu.grid(row=0, column=0, padx=(12, 8), pady=12, sticky="w")

        self.cancel_search_button = ctk.CTkButton(
            action_frame,
            text="取消",
            width=82,
            fg_color="#dc2626",
            hover_color="#b91c1c",
            text_color="#ffffff",
            command=self.cancel_current_search,
        )
        self.cancel_search_button.grid(row=0, column=2, padx=(8, 0), pady=12, sticky="e")
        self.cancel_search_button.grid_remove()

        self.run_search_button = ctk.CTkButton(action_frame, text="检索", width=110, command=self.run_selected_search)
        self.run_search_button.grid(row=0, column=3, padx=(8, 12), pady=12, sticky="e")

        self.status_label = ctk.CTkLabel(action_frame, text="请选择文件夹后执行检索", anchor="w")
        self.status_label.grid(row=0, column=1, padx=8, pady=12, sticky="ew")

        # 工作表与单元格检索固定使用表格；文件检索使用缩略图卡片。
        list_frame = ctk.CTkFrame(self)
        self.result_list_frame = list_frame
        list_frame.grid(row=3, column=0, padx=16, pady=(8, 16), sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        self.result_group_controls = ctk.CTkFrame(list_frame, fg_color="transparent")
        self.result_group_controls.grid(
            row=0,
            column=0,
            columnspan=2,
            padx=8,
            pady=(6, 2),
            sticky="ew",
        )
        self.result_group_controls.grid_columnconfigure(0, weight=1)

        self.collapse_all_results_button = ctk.CTkButton(
            self.result_group_controls,
            text="全部收起",
            width=88,
            height=28,
            command=self.collapse_all_results,
        )
        self.collapse_all_results_button.grid(row=0, column=1, padx=(0, 6), sticky="e")

        self.expand_all_results_button = ctk.CTkButton(
            self.result_group_controls,
            text="全部展开",
            width=88,
            height=28,
            command=self.expand_all_results,
        )
        self.expand_all_results_button.grid(row=0, column=2, sticky="e")
        self.result_group_controls.grid_remove()

        self.file_result_view_controls = ctk.CTkFrame(list_frame, fg_color="transparent")
        self.file_result_view_controls.grid(
            row=0,
            column=0,
            columnspan=2,
            padx=8,
            pady=(6, 2),
            sticky="ew",
        )
        self.file_result_view_controls.grid_columnconfigure(0, weight=1)
        self.file_result_count_label = ctk.CTkLabel(
            self.file_result_view_controls,
            text="",
            anchor="w",
        )
        self.file_result_count_label.grid(row=0, column=0, padx=(0, 12), sticky="w")
        self.file_result_view_switch = ctk.CTkSegmentedButton(
            self.file_result_view_controls,
            values=["缩略图", "表格"],
            variable=self.file_result_view_var,
            command=self._on_file_result_view_changed,
            width=176,
        )
        self.file_result_view_switch.grid(row=0, column=1, sticky="e")
        self.file_result_view_controls.grid_remove()

        self._configure_tree_style()
        self.result_tree = ttk.Treeview(
            list_frame,
            columns=(),
            show="headings",
            selectmode="browse",
        )
        self.result_tree.grid(row=1, column=0, sticky="nsew")
        self.result_tree.bind("<ButtonRelease-1>", self._handle_result_click)
        self.result_tree.bind("<Double-1>", self._handle_result_double_click)
        self.result_tree.bind("<Configure>", lambda _event: self._schedule_action_button_layout())
        self.result_tree.bind("<Configure>", lambda _event: self._schedule_result_grid_lines(), add="+")
        self.result_tree.tag_configure("reference", foreground="#2563eb")
        self.result_tree.tag_configure("cell-workbook", background="#4b5563", foreground="#ffffff")
        self.result_tree.tag_configure("cell-sheet", background="#e5e7eb", foreground="#111827")
        self._configure_result_columns(WORKSHEET_RESULT_COLUMNS)

        self.result_tree_vertical_scroll = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=self._scroll_result_tree_y,
        )
        self.result_tree_vertical_scroll.grid(row=1, column=1, sticky="ns")
        self.result_tree_horizontal_scroll = ttk.Scrollbar(
            list_frame,
            orient="horizontal",
            command=self._scroll_result_tree_x,
        )
        self.result_tree_horizontal_scroll.grid(row=2, column=0, sticky="ew")
        self.result_tree.configure(
            yscrollcommand=lambda *args: self._update_scrollbar(self.result_tree_vertical_scroll, *args),
            xscrollcommand=lambda *args: self._update_scrollbar(self.result_tree_horizontal_scroll, *args),
        )

        self.file_thumbnail_frame = ctk.CTkScrollableFrame(
            list_frame,
            fg_color=THUMBNAIL_BACKGROUND_COLOR,
            corner_radius=0,
        )
        self.file_thumbnail_frame.grid(
            row=1,
            column=0,
            columnspan=2,
            padx=4,
            pady=(0, 4),
            sticky="nsew",
        )
        for column_index in range(FILE_THUMBNAIL_COLUMNS):
            self.file_thumbnail_frame.grid_columnconfigure(column_index, weight=1, uniform="file-thumbnail")
        self.file_thumbnail_frame.grid_remove()

    def _configure_tree_style(self) -> None:
        """统一结果列表的行高和字体样式。"""

        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview", rowheight=30, font=("Microsoft YaHei UI", 10))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))

    def _show_table_results(self) -> None:
        """显示工作表或单元格检索使用的表格视图。"""

        self.file_result_view_controls.grid_remove()
        self.file_thumbnail_frame.grid_remove()
        self.result_tree.grid()
        self.result_tree_vertical_scroll.grid()
        self.result_tree_horizontal_scroll.grid()
        self._schedule_result_grid_lines()

    def _show_file_thumbnail_results(self) -> None:
        """显示文件检索使用的缩略图视图。"""

        self.result_group_controls.grid_remove()
        self.result_tree.grid_remove()
        self.result_tree_vertical_scroll.grid_remove()
        self.result_tree_horizontal_scroll.grid_remove()
        self._hide_result_grid_lines()
        self.file_result_view_controls.grid()
        self.file_thumbnail_frame.grid()

    def _show_file_table_results(self) -> None:
        """显示按文件夹分组的文件检索表格视图。"""

        self.file_thumbnail_frame.grid_remove()
        self.file_result_view_controls.grid()
        self.result_tree.grid()
        self.result_tree_vertical_scroll.grid()
        self.result_tree_horizontal_scroll.grid()
        self._schedule_result_grid_lines()

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
        self.result_tree.heading("#0", text=tree_heading or "", anchor="w")
        self.result_tree.column(
            "#0",
            width=tree_width,
            minwidth=tree_min_width,
            stretch=tree_stretch,
            anchor="w",
        )

        for key, title, width, min_width, stretch in columns:
            self.result_tree.heading(key, text=title, anchor="w")
            self.result_tree.column(
                key,
                width=width,
                minwidth=min_width,
                stretch=stretch,
                anchor="w",
            )
        self._schedule_result_grid_lines()

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

    def _create_directory_button(self, item_id: str) -> None:
        """为单元格检索中的工作簿目录行创建“打开目录”按钮。"""

        button = ctk.CTkButton(
            self.result_tree,
            text="打开目录",
            width=96,
            height=24,
            corner_radius=5,
            fg_color="#475569",
            hover_color="#334155",
            text_color="#ffffff",
            command=lambda item=item_id: self._open_directory_for_item(item),
        )
        self.directory_buttons[item_id] = button

    def _open_directory_for_item(self, item_id: str) -> None:
        """在资源管理器中打开工作簿所在目录并选中文件。"""

        target = self.item_targets.get(item_id)
        if target is None:
            return
        self._open_target(ResultTarget(target.file_path, open_containing_folder=True))

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
        self._schedule_result_grid_lines()

    def _schedule_result_grid_lines(self) -> None:
        """在表格空闲后重绘可见区域的横、竖分隔线。"""

        if self.is_closing or self.result_grid_layout_job is not None:
            return
        self.result_grid_layout_job = self.after_idle(self._layout_result_grid_lines)

    def _layout_result_grid_lines(self) -> None:
        """根据当前可见行和列在 Treeview 上画出表格网格线。"""

        self.result_grid_layout_job = None
        self._hide_result_grid_lines()
        if not self.result_tree.winfo_ismapped():
            return

        visible_item_ids = self._get_visible_result_item_ids()
        if not visible_item_ids:
            return

        tree_x = self.result_tree.winfo_x()
        tree_y = self.result_tree.winfo_y()
        tree_width = self.result_tree.winfo_width()
        tree_height = self.result_tree.winfo_height()
        first_row_bbox = self.result_tree.bbox(visible_item_ids[0])
        if not first_row_bbox:
            return

        # 表头下沿和每一可见行下沿均绘制横线。
        self._create_result_grid_line(tree_x, tree_y + first_row_bbox[1] - 1, tree_width, 1)
        for item_id in visible_item_ids:
            row_bbox = self.result_tree.bbox(item_id)
            if row_bbox:
                self._create_result_grid_line(
                    tree_x,
                    tree_y + row_bbox[1] + row_bbox[3] - 1,
                    tree_width,
                    1,
                )

        # 以首个可见行的单元格边界为准，绘制贯穿表头和内容区的竖线。
        visible_columns: list[str] = []
        if "tree" in str(self.result_tree.cget("show")):
            visible_columns.append("#0")
        visible_columns.extend(self.current_result_columns)
        for column_id in visible_columns:
            cell_bbox = self.result_tree.bbox(visible_item_ids[0], column_id)
            if cell_bbox:
                self._create_result_grid_line(
                    tree_x + cell_bbox[0] + cell_bbox[2] - 1,
                    tree_y,
                    1,
                    tree_height,
                )

    def _get_visible_result_item_ids(self) -> list[str]:
        """返回当前滚动位置里实际可见的树节点。"""

        visible_items: list[str] = []
        seen_items: set[str] = set()
        # 只探测当前视口，不遍历全部结果，避免数万条命中时滚动卡顿。
        for y_position in range(0, self.result_tree.winfo_height(), 4):
            item_id = self.result_tree.identify_row(y_position)
            if item_id and item_id not in seen_items:
                visible_items.append(item_id)
                seen_items.add(item_id)
        return visible_items

    def _create_result_grid_line(self, x: int, y: int, width: int, height: int) -> None:
        """创建一条不影响列宽的覆盖式网格线。"""

        line = TkFrame(self.result_list_frame, bg="#cbd5e1", highlightthickness=0, bd=0)
        line.place(x=x, y=y, width=width, height=height)
        line.lift()
        self.result_grid_lines.append(line)

    def _hide_result_grid_lines(self) -> None:
        """删除上一帧网格线，防止滚动或切换视图后残留。"""

        for line in self.result_grid_lines:
            line.destroy()
        self.result_grid_lines.clear()

    def _schedule_action_button_layout(self) -> None:
        """把按钮重排延迟到 Tk 空闲时执行，避免滚动时频繁计算。"""

        if self.is_closing or self.action_layout_job is not None:
            return
        self.action_layout_job = self.after_idle(self._layout_action_buttons)

    def _layout_action_buttons(self) -> None:
        """把真实按钮摆放到 Treeview 的“操作”列单元格上。"""

        self.action_layout_job = None
        if "action" not in self.current_result_columns:
            for button in (*self.action_buttons.values(), *self.directory_buttons.values()):
                button.place_forget()
            return

        button_items = (*self.action_buttons.items(), *self.directory_buttons.items())
        for item_id, button in button_items:
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
            button.configure(width=button_width, height=button_height)
            button.place(
                x=x + max(4, (width - button_width) // 2),
                y=y + max(3, (height - button_height) // 2),
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
        """用带鼠标抓取的弹出方式打开菜单，避免点击穿透到底层控件。"""

        self._sync_folder_dropdown_width()
        dropdown_menu = getattr(self.folder_entry, "_dropdown_menu", None)
        if dropdown_menu is None:
            return

        current_height = getattr(self.folder_entry, "_current_height", self.folder_entry.winfo_height())
        y_offset = self.folder_entry._apply_widget_scaling(current_height + 0)
        platform_offset = dropdown_menu._apply_widget_scaling(3)
        menu_x = int(self.folder_entry.winfo_rootx())
        menu_y = int(self.folder_entry.winfo_rooty() + y_offset + platform_offset)
        self.folder_entry._close_on_next_click = False
        # CustomTkinter 在 Windows 默认调用 Menu.post()，该方式没有完整的菜单抓取，
        # 点击菜单项时可能继续触发下方 Treeview。tk_popup() 会在菜单存续期间接管鼠标事件。
        dropdown_menu.tk_popup(menu_x, menu_y)

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
            self.background_image_path,
        )

    def open_search_settings(self) -> None:
        """打开当前检索类型对应的筛选设置窗口。"""

        self.search_settings_dialog.open(
            self.search_mode_var.get(),
            self.only_enabled_sheets_var.get(),
            self.only_enabled_data_var.get(),
            self.only_specified_suffixes_var.get(),
            self.file_suffixes_var.get(),
            self.search_folders_var.get(),
            self.exact_match_var.get(),
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
        background_image_path: str,
    ) -> bool:
        """应用设置弹窗保存后的配置，并刷新当前列表里的启用状态。"""

        if not self._save_background_image(background_image_path):
            return False
        self.reference_config = reference_config
        self.disabled_sheet_marker = disabled_sheet_marker
        self.data_filter_config = data_filter_config
        self.help_url = help_url.strip()
        self._save_cache_from_inputs()
        self._refresh_sheet_statuses()
        self.status_label.configure(text="设置已保存")
        return True

    def _apply_search_settings(
        self,
        only_enabled_sheets: bool,
        only_enabled_data: bool,
        only_specified_suffixes: bool,
        file_suffixes: str,
        search_folders: bool,
        exact_match: bool,
    ) -> None:
        """应用检索设置窗口保存的筛选项。"""

        self.only_enabled_sheets_var.set(only_enabled_sheets)
        self.only_enabled_data_var.set(only_enabled_data)
        self.only_specified_suffixes_var.set(only_specified_suffixes)
        self.file_suffixes_var.set(file_suffixes)
        self.search_folders_var.set(search_folders)
        self.exact_match_var.set(exact_match)
        self._save_cache_from_inputs()
        self.status_label.configure(text="检索设置已保存")

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

    def _on_search_mode_changed(self, selected_mode: str) -> None:
        """切换检索类型，并同步已打开的检索设置窗口。"""

        self.search_mode_var.set(selected_mode)
        self.search_settings_dialog.refresh(selected_mode)
        self._save_cache_from_inputs()

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
        self.current_exact_match = self.exact_match_var.get()
        self._remember_folder_path(folder_path)
        self._save_cache_from_inputs()
        cancel_event = self._begin_search()
        match_mode_text = "精确检索" if self.current_exact_match else "检索"
        self.status_label.configure(text=f"正在{match_mode_text} Excel 文件，请稍等")
        self._clear_results()

        # Excel 文件数量可能较多，放到后台线程，避免界面卡死。
        worker = Thread(
            target=self._scan_in_background,
            args=(folder_path, self.current_keyword, self.current_exact_match, cancel_event),
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
        exact_match = self.exact_match_var.get()
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
        match_description = "完全等于" if exact_match else "包含"
        self.status_label.configure(text=f"正在搜索{scope_text}{match_description}“{keyword}”的单元格")
        self._clear_results()

        worker = Thread(
            target=self._search_cells_in_background,
            args=(folder_path, keyword, disabled_marker, data_filter_config, exact_match, cancel_event),
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
        exact_match = self.exact_match_var.get()
        include_folders = self.search_folders_var.get()
        suffixes: tuple[str, ...] = ()
        if self.only_specified_suffixes_var.get():
            suffixes = normalize_file_suffixes(self.file_suffixes_var.get())
            if not suffixes:
                messagebox.showinfo("缺少文件后缀", "请先输入要检索的文件后缀。")
                return
        self._remember_folder_path(folder_path)
        self._save_cache_from_inputs()
        cancel_event = self._begin_search()
        suffix_scope = f"指定后缀（{'、'.join(suffixes)}）的" if suffixes else ""
        match_mode_text = "精确检索" if exact_match else "检索"
        target_text = "文件和文件夹" if include_folders else "文件"
        self.status_label.configure(text=f"正在{match_mode_text}{suffix_scope}{target_text}名称和地址，请稍等")
        self._clear_results()

        # 文件名检索也可能遍历大量子目录，所以同样放到后台。
        worker = Thread(
            target=self._search_files_in_background,
            args=(folder_path, keyword, suffixes, exact_match, include_folders, cancel_event),
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

    def _scan_in_background(
        self,
        folder_path: Path,
        keyword: str,
        exact_match: bool,
        cancel_event: Event,
    ) -> None:
        """后台执行工作表检索，并把结果切回主线程渲染。"""

        try:
            results = scan_excel_workbooks(
                folder_path,
                keyword,
                cancel_event=cancel_event,
                exact_match=exact_match,
            )
        except ExcelReadError as exc:
            message = str(exc)
            if cancel_event.is_set():
                self._post_search_callback(cancel_event, self._search_cancelled)
            else:
                self._post_search_callback(cancel_event, lambda message=message: self._scan_failed(message))
            return
        except Exception as exc:
            message = f"发生未预期错误：{exc}"
            self._post_search_callback(cancel_event, lambda message=message: self._scan_failed(message))
            return

        if cancel_event.is_set():
            self._post_search_callback(cancel_event, self._search_cancelled)
            return
        self._post_search_callback(cancel_event, lambda results=results: self._scan_finished(results))

    def _search_cells_in_background(
        self,
        folder_path: Path,
        keyword: str,
        disabled_marker: str | None,
        data_filter_config: DataFilterConfig | None,
        exact_match: bool,
        cancel_event: Event,
    ) -> None:
        """后台执行单元格检索，并把结果切回主线程渲染。"""

        issues: list[ExcelCellSearchIssue] = []
        try:
            matches = search_excel_cells(
                folder_path,
                keyword,
                disabled_sheet_marker=disabled_marker,
                data_filter_config=data_filter_config,
                cancel_event=cancel_event,
                issues=issues,
                exact_match=exact_match,
            )
        except ExcelReadError as exc:
            message = str(exc)
            if cancel_event.is_set():
                self._post_search_callback(cancel_event, self._search_cancelled)
            else:
                self._post_search_callback(cancel_event, lambda message=message: self._cell_search_failed(message))
            return
        except Exception as exc:
            message = f"发生未预期错误：{exc}"
            self._post_search_callback(cancel_event, lambda message=message: self._cell_search_failed(message))
            return

        if cancel_event.is_set():
            self._post_search_callback(cancel_event, self._search_cancelled)
            return
        self._post_search_callback(
            cancel_event,
            lambda matches=matches, keyword=keyword, issues=issues: self._cell_search_finished(
                matches,
                keyword,
                issues,
            ),
        )

    def _search_files_in_background(
        self,
        folder_path: Path,
        keyword: str,
        suffixes: tuple[str, ...],
        exact_match: bool,
        include_folders: bool,
        cancel_event: Event,
    ) -> None:
        """后台执行文件名检索，并把结果切回主线程渲染。"""

        try:
            results = search_files(
                folder_path,
                keyword,
                cancel_event=cancel_event,
                suffixes=suffixes,
                exact_match=exact_match,
                include_folders=include_folders,
            )
        except FileSearchError as exc:
            message = str(exc)
            if cancel_event.is_set():
                self._post_search_callback(cancel_event, self._search_cancelled)
            else:
                self._post_search_callback(cancel_event, lambda message=message: self._file_search_failed(message))
            return
        except Exception as exc:
            message = f"发生未预期错误：{exc}"
            self._post_search_callback(cancel_event, lambda message=message: self._file_search_failed(message))
            return

        if cancel_event.is_set():
            self._post_search_callback(cancel_event, self._search_cancelled)
            return
        self._post_search_callback(cancel_event, lambda results=results: self._file_search_finished(results))

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

    def _cell_search_finished(
        self,
        matches: list[ExcelCellMatch],
        keyword: str,
        issues: list[ExcelCellSearchIssue],
    ) -> None:
        """渲染单元格检索结果并更新状态栏统计。"""

        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_cell_matches(matches)
        status_parts = [f"找到 {len(matches)} 条匹配"]
        if len(matches) >= DEFAULT_MAX_SEARCH_RESULTS:
            status_parts.append(f"已达到显示上限 {DEFAULT_MAX_SEARCH_RESULTS} 条")
        if issues:
            status_parts.append(f"跳过 {len(issues)} 个无法读取的文件")
        self.status_label.configure(text="单元格搜索完成：" + "，".join(status_parts))

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
        file_count = sum(not result.is_directory for result in results)
        folder_count = sum(result.is_directory for result in results)
        limit_text = (
            f"，已达到显示上限 {DEFAULT_MAX_FILE_SEARCH_RESULTS} 条"
            if len(results) >= DEFAULT_MAX_FILE_SEARCH_RESULTS
            else ""
        )
        count_text = f"{file_count} 个文件"
        if folder_count:
            count_text += f"，{folder_count} 个文件夹"
        self.status_label.configure(text=f"文件检索完成：找到 {count_text}{limit_text}")

    def _set_search_controls_running(self, is_running: bool) -> None:
        """切换检索中/空闲状态下按钮和选项的可用性。"""

        if is_running:
            self.search_settings_dialog.close()
            self.run_search_button.configure(state="disabled", text="检索中...")
            self.search_mode_menu.configure(state="disabled")
            if self.cancel_search_button is not None:
                self.cancel_search_button.configure(state="normal", text="取消")
                self.cancel_search_button.grid()
            if self.settings_button is not None:
                self.settings_button.configure(state="disabled")
            if self.search_settings_button is not None:
                self.search_settings_button.configure(state="disabled")
            return

        self.run_search_button.configure(state="normal", text="检索")
        self.search_mode_menu.configure(state="normal")
        if self.cancel_search_button is not None:
            self.cancel_search_button.grid_remove()
        if self.settings_button is not None:
            self.settings_button.configure(state="normal")
        if self.search_settings_button is not None:
            self.search_settings_button.configure(state="normal")

    def _render_cell_matches(self, matches: list[ExcelCellMatch]) -> None:
        """按工作簿、工作表两级目录展示单元格命中。"""

        self._clear_results()
        self._show_table_results()
        self._configure_result_columns(
            CELL_RESULT_COLUMNS,
            tree_heading="检索目录",
            tree_width=360,
            tree_min_width=180,
            tree_stretch=True,
        )

        matches_by_workbook: dict[Path, list[ExcelCellMatch]] = {}
        for match in matches:
            matches_by_workbook.setdefault(match.path, []).append(match)

        for workbook_path, workbook_matches in matches_by_workbook.items():
            # 工作簿和工作表都是分类目录；只有具体单元格命中行会打开 Excel。
            workbook_item_id = self.result_tree.insert(
                "",
                "end",
                text=workbook_path.name,
                values=(f"{len(workbook_matches)} 条匹配", ""),
                open=True,
                tags=("cell-workbook",),
            )
            self.item_targets[workbook_item_id] = ResultTarget(workbook_path)
            self._create_directory_button(workbook_item_id)

            matches_by_sheet: dict[str, list[ExcelCellMatch]] = {}
            for match in workbook_matches:
                matches_by_sheet.setdefault(match.sheet_name, []).append(match)

            for sheet_name, sheet_matches in matches_by_sheet.items():
                sheet_item_id = self.result_tree.insert(
                    workbook_item_id,
                    "end",
                    text=sheet_name,
                    values=(f"{len(sheet_matches)} 条匹配", ""),
                    open=True,
                    tags=("cell-sheet",),
                )
                self.item_targets[sheet_item_id] = ResultTarget(workbook_path, sheet_name)

                for match in sheet_matches:
                    item_id = self.result_tree.insert(
                        sheet_item_id,
                        "end",
                        text=format_excel_cell_address(match.row_index, match.column_index),
                        values=(
                            format_cell_value(match.value),
                            "",
                        ),
                    )
                    self.item_targets[item_id] = ResultTarget(
                        match.path,
                        match.sheet_name,
                        match.row_index,
                        match.column_index,
                    )
        self._update_result_group_controls_visibility()
        self._schedule_action_button_layout()

    def _render_file_results(self, results: list[FileSearchResult]) -> None:
        """保存文件检索结果，并按当前视图渲染。"""

        self._clear_results()
        self.file_thumbnail_results = list(results)
        self._render_file_result_view()

        if not results:
            self.status_label.configure(text="没有找到匹配的文件")

    def _on_file_result_view_changed(self, _selected_view: str) -> None:
        """在缩略图和表格之间切换当前文件检索结果。"""

        if self.search_mode_var.get() != "文件检索":
            return
        self._render_file_result_view()

    def _render_file_result_view(self) -> None:
        """清理旧视图后，以用户选定方式展示完整文件结果。"""

        self._clear_file_result_view()
        total_count = len(self.file_thumbnail_results)
        self.file_result_count_label.configure(text=f"文件检索结果：{total_count} 项")
        if self.file_result_view_var.get() == "表格":
            self._show_file_table_results()
            self._render_file_results_as_table()
            return

        self._show_file_thumbnail_results()
        self._render_file_results_as_thumbnails()

    def _clear_file_result_view(self) -> None:
        """清空文件视图的卡片、表格行和临时打开目标。"""

        self.result_group_controls.grid_remove()
        self._clear_file_thumbnail_cards()
        self.item_targets.clear()
        for item_id in self.result_tree.get_children():
            self.result_tree.delete(item_id)

    def _render_file_results_as_thumbnails(self) -> None:
        """按所在文件夹分组，并将每组完整展示在标题下方。"""

        grouped_results = self._group_file_results_by_folder()
        if not grouped_results:
            empty_label = ctk.CTkLabel(
                self.file_thumbnail_frame,
                text="没有找到匹配的文件或文件夹",
                text_color=THUMBNAIL_TEXT_COLOR,
                font=("Microsoft YaHei UI", 14),
            )
            empty_label.grid(
                row=0,
                column=0,
                columnspan=FILE_THUMBNAIL_COLUMNS,
                padx=16,
                pady=28,
            )
            return

        for group_index, (folder_path, folder_results) in enumerate(grouped_results.items()):
            folder_label = ctk.CTkLabel(
                self.file_thumbnail_frame,
                text=f"📁  {folder_path}  ·  {len(folder_results)} 项",
                anchor="w",
                text_color=THUMBNAIL_TEXT_COLOR,
                font=("Microsoft YaHei UI", 12, "bold"),
            )
            folder_label.grid(
                row=group_index * 2,
                column=0,
                columnspan=FILE_THUMBNAIL_COLUMNS,
                padx=10,
                pady=(18 if group_index else 8, 2),
                sticky="ew",
            )
            section_frame = ctk.CTkFrame(
                self.file_thumbnail_frame,
                fg_color=THUMBNAIL_BACKGROUND_COLOR,
            )
            section_frame.grid(
                row=group_index * 2 + 1,
                column=0,
                columnspan=FILE_THUMBNAIL_COLUMNS,
                padx=2,
                pady=(0, 4),
                sticky="ew",
            )
            for column_index in range(FILE_THUMBNAIL_COLUMNS):
                section_frame.grid_columnconfigure(column_index, weight=1, uniform="file-thumbnail")
            for card_index, result in enumerate(folder_results):
                self._create_file_thumbnail_card(section_frame, result, card_index)

    def _render_file_results_as_table(self) -> None:
        """按所在文件夹分组，以左对齐表格展示文件结果。"""

        self._configure_result_columns(
            FILE_TABLE_RESULT_COLUMNS,
            tree_heading="文件夹",
            tree_width=360,
            tree_min_width=220,
            tree_stretch=True,
        )
        grouped_results = self._group_file_results_by_folder()
        if not grouped_results:
            return

        for folder_path, folder_results in grouped_results.items():
            folder_item_id = self.result_tree.insert(
                "",
                "end",
                text=str(folder_path),
                values=("", "", f"{len(folder_results)} 项"),
                open=True,
            )
            self.item_targets[folder_item_id] = ResultTarget(folder_path)
            for result in folder_results:
                item_type = self._get_file_thumbnail_style(result)[2]
                item_id = self.result_tree.insert(
                    folder_item_id,
                    "end",
                    text="",
                    values=(result.file_name, item_type, str(result.path)),
                )
                self.item_targets[item_id] = ResultTarget(
                    result.path,
                    open_containing_folder=not result.is_directory,
                )
        self._update_result_group_controls_visibility()

    def _group_file_results_by_folder(self) -> dict[Path, list[FileSearchResult]]:
        """保持检索排序，将同一所在文件夹的结果连续分组。"""

        grouped_results: dict[Path, list[FileSearchResult]] = {}
        for result in self.file_thumbnail_results:
            grouped_results.setdefault(result.path.parent, []).append(result)
        return grouped_results

    def _clear_file_thumbnail_cards(self) -> None:
        """删除缩略图卡片，释放图片和控件引用。"""

        for widget in self.file_thumbnail_frame.winfo_children():
            widget.destroy()
        self.file_thumbnail_cards.clear()
        self.file_thumbnail_images.clear()

    def _create_file_thumbnail_card(
        self,
        parent: ctk.CTkFrame,
        result: FileSearchResult,
        card_index: int,
    ) -> None:
        """创建单个文件或文件夹的缩略图卡片。"""

        symbol, color, item_type = self._get_file_thumbnail_style(result)
        target = ResultTarget(
            result.path,
            open_containing_folder=not result.is_directory,
        )
        card = ctk.CTkFrame(
            parent,
            width=FILE_THUMBNAIL_CARD_WIDTH,
            height=FILE_THUMBNAIL_CARD_HEIGHT,
            fg_color=THUMBNAIL_BACKGROUND_COLOR,
            border_width=1,
            border_color="#e5e7eb",
        )
        card.grid(
            row=card_index // FILE_THUMBNAIL_COLUMNS,
            column=card_index % FILE_THUMBNAIL_COLUMNS,
            padx=8,
            pady=8,
            sticky="nsew",
        )
        card.grid_propagate(False)
        card.grid_columnconfigure(1, weight=1)

        image_thumbnail = self._create_image_thumbnail(result.path)
        if image_thumbnail is not None:
            icon_label = ctk.CTkLabel(
                card,
                text="",
                image=image_thumbnail,
                width=64,
                height=64,
                corner_radius=16,
                fg_color=THUMBNAIL_BACKGROUND_COLOR,
            )
            self.file_thumbnail_images.append(image_thumbnail)
        else:
            icon_label = ctk.CTkLabel(
                card,
                text=symbol,
                width=64,
                height=64,
                corner_radius=16,
                fg_color=THUMBNAIL_BACKGROUND_COLOR,
                text_color=color,
                font=("Segoe UI Emoji", 32),
            )
        icon_label.grid(row=0, column=0, rowspan=2, padx=(12, 10), pady=(12, 8), sticky="nw")

        name_label = ctk.CTkLabel(
            card,
            text=self._shorten_thumbnail_text(result.file_name, 24),
            anchor="w",
            text_color=THUMBNAIL_TEXT_COLOR,
            font=("Microsoft YaHei UI", 12, "bold"),
        )
        name_label.grid(row=0, column=1, padx=(0, 12), pady=(14, 0), sticky="ew")
        type_label = ctk.CTkLabel(
            card,
            text=item_type,
            anchor="w",
            text_color=THUMBNAIL_SECONDARY_TEXT_COLOR,
            font=("Microsoft YaHei UI", 10),
        )
        type_label.grid(row=1, column=1, padx=(0, 12), pady=(0, 8), sticky="ew")

        location = result.path if result.is_directory else result.path.parent
        location_label = ctk.CTkLabel(
            card,
            text=self._shorten_thumbnail_text(str(location), 40),
            anchor="w",
            text_color=THUMBNAIL_SECONDARY_TEXT_COLOR,
            font=("Microsoft YaHei UI", 10),
        )
        location_label.grid(row=2, column=0, columnspan=2, padx=12, pady=(0, 8), sticky="ew")

        open_button = ctk.CTkButton(
            card,
            text="打开文件夹" if result.is_directory else "打开所在目录",
            height=28,
            command=lambda target=target: self._open_target(target),
        )
        open_button.grid(row=3, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="ew")
        self.file_thumbnail_cards.append(card)

    def _create_image_thumbnail(self, file_path: Path) -> ctk.CTkImage | None:
        """读取常见图片文件，生成等比例真实预览缩略图。"""

        if file_path.suffix.lower() not in IMAGE_FILE_SUFFIXES:
            return None
        try:
            with Image.open(file_path) as source_image:
                source_image.thumbnail(FILE_THUMBNAIL_IMAGE_SIZE, Image.Resampling.LANCZOS)
                preview_image = source_image.convert("RGBA")
        except (OSError, UnidentifiedImageError, ValueError):
            return None
        return ctk.CTkImage(
            light_image=preview_image,
            dark_image=preview_image,
            size=preview_image.size,
        )

    @staticmethod
    def _shorten_thumbnail_text(value: str, max_length: int) -> str:
        """让缩略图卡片中的文件名和路径保持单行可读。"""

        if len(value) <= max_length:
            return value
        return value[: max_length - 1] + "…"

    @staticmethod
    def _get_file_thumbnail_style(result: FileSearchResult) -> tuple[str, str, str]:
        """按文件类型返回缩略图图标、主色和说明文字。"""

        if result.is_directory:
            return "📁", "#b7791f", "文件夹"

        suffix = result.path.suffix.lower()
        if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".csv"}:
            return "📊", "#217346", "Excel / 表格文件"
        if suffix in {".doc", ".docx", ".rtf", ".txt", ".md", ".pdf"}:
            return "📄", "#2563eb", "文档"
        if suffix in IMAGE_FILE_SUFFIXES or suffix == ".svg":
            return "🖼", "#7c3aed", "图片"
        if suffix in {".mp4", ".avi", ".mkv", ".mov", ".wmv"}:
            return "🎬", "#be185d", "视频"
        if suffix in {".mp3", ".wav", ".flac", ".aac", ".m4a"}:
            return "🎵", "#0f766e", "音频"
        if suffix in {".zip", ".rar", ".7z", ".tar", ".gz"}:
            return "🗜", "#92400e", "压缩包"
        return "📄", "#475569", "文件"

    def collapse_all_results(self) -> None:
        """一键收起结果列表中的所有分组。"""

        self._set_all_result_groups_open(False)

    def expand_all_results(self) -> None:
        """一键展开结果列表中的所有分组。"""

        self._set_all_result_groups_open(True)

    def _set_all_result_groups_open(self, is_open: bool) -> None:
        """递归设置结果树中所有含子节点项目的展开状态。"""

        pending_items = list(self.result_tree.get_children())
        while pending_items:
            item_id = pending_items.pop()
            child_items = self.result_tree.get_children(item_id)
            if child_items:
                self.result_tree.item(item_id, open=is_open)
                pending_items.extend(child_items)

    def _update_result_group_controls_visibility(self) -> None:
        """仅在结果中存在可展开分组时显示批量展开和收起按钮。"""

        pending_items = list(self.result_tree.get_children())
        while pending_items:
            item_id = pending_items.pop()
            child_items = self.result_tree.get_children(item_id)
            if child_items:
                self.result_group_controls.grid()
                return
            pending_items.extend(child_items)
        self.result_group_controls.grid_remove()

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

        self._open_target(target)

    def _open_target(self, target: ResultTarget) -> None:
        """异步打开一个已确定的目标，供表格行和缩略图卡片共用。"""

        if target in self.opening_targets:
            return
        self.opening_targets.add(target)
        self.status_label.configure(text=f"正在打开：{target.file_path.name}")
        worker = Thread(target=self._open_target_in_background, args=(target,), daemon=True)
        worker.start()

    def _open_target_in_background(self, target: ResultTarget) -> None:
        """在后台打开目标，避免 Excel COM 或系统关联程序阻塞 Tk 主线程。"""

        try:
            if target.open_containing_folder:
                open_containing_folder(target.file_path)
                located = True
            else:
                located = open_file(target.file_path, target.sheet_name, target.row_index, target.column_index)
        except (OSError, subprocess.SubprocessError) as exc:
            message = str(exc)
            self._post_to_ui(
                lambda target=target, message=message: self._open_target_failed(target, message)
            )
            return
        except Exception as exc:
            message = str(exc)
            self._post_to_ui(
                lambda target=target, message=message: self._open_target_failed(target, message)
            )
            return

        self._post_to_ui(lambda target=target, located=located: self._open_target_finished(target, located))

    def _open_target_finished(self, target: ResultTarget, located: bool) -> None:
        """目标打开完成后更新状态。"""

        self.opening_targets.discard(target)
        if target.sheet_name and not located:
            self.status_label.configure(text="文件已打开，但未能自动定位到指定工作表或单元格")
        else:
            self.status_label.configure(text=f"已打开：{target.file_path.name}")

    def _open_target_failed(self, target: ResultTarget, message: str) -> None:
        """显示后台打开文件失败信息。"""

        self.opening_targets.discard(target)
        self.status_label.configure(text="打开失败")
        messagebox.showerror("打开失败", f"无法打开文件：{target.file_path}\n{message}")

    def _clear_results(self) -> None:
        """清空表格或缩略图结果、行目标映射和悬浮操作按钮。"""

        self.result_group_controls.grid_remove()
        self._hide_result_grid_lines()
        self.file_thumbnail_results.clear()
        self._clear_file_thumbnail_cards()
        self.result_generation += 1
        for cancel_event in self.reference_cancel_events.values():
            cancel_event.set()
        self.reference_cancel_events.clear()
        self.item_targets.clear()
        self.reference_items_by_source.clear()
        self.reference_searching_items.clear()
        for button in self.action_buttons.values():
            button.destroy()
        self.action_buttons.clear()
        for button in self.directory_buttons.values():
            button.destroy()
        self.directory_buttons.clear()
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

    def _render_results(self, results: list[ExcelWorkbookInfo]) -> None:
        """展示工作表检索结果，每个工作表一行并附带查找引用按钮。"""

        self._clear_results()
        self._show_table_results()
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
                        "",
                    ),
                )
                self.item_targets[item] = ResultTarget(result.path, sheet_name)
                self._create_action_button(item)
                has_display_rows = True

        if not has_display_rows:
            self.status_label.configure(text="没有找到匹配的启用工作表")
        self._schedule_action_button_layout()

    def _handle_result_click(self, event: object) -> str | None:
        """处理单击事件，按行类型执行引用查找或打开目录。"""

        column_key = self._get_column_key(self.result_tree.identify_column(event.x))
        if column_key != "action":
            return None

        item_id = self.result_tree.identify_row(event.y)
        if not item_id:
            return "break"
        if item_id in self.directory_buttons:
            self.result_tree.selection_set(item_id)
            self.result_tree.focus(item_id)
            self._open_directory_for_item(item_id)
            return "break"
        if not self._action_button_is_available(item_id):
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
                if item_id in self.directory_buttons:
                    self._open_directory_for_item(item_id)
                else:
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
        if not self._action_button_is_available(item_id):
            return

        button = self.action_buttons.get(item_id)

        config = self.reference_config
        cancel_event = Event()
        generation = self.result_generation
        self.reference_cancel_events[item_id] = cancel_event
        self.reference_searching_items.add(item_id)
        if button is not None:
            button.configure(text="查找中...", state="disabled")
        self._remove_reference_rows(item_id)
        self.status_label.configure(text=f"正在查找 {target.sheet_name} 第{config.reference_row_index}行引用")

        # 引用查找可能会打开工作簿，所以放到后台避免阻塞界面。
        worker = Thread(
            target=self._find_references_in_background,
            args=(item_id, target, config, generation, cancel_event),
            daemon=True,
        )
        worker.start()

    def _find_references_in_background(
        self,
        item_id: str,
        target: ResultTarget,
        config: ReferenceLookupConfig,
        generation: int,
        cancel_event: Event,
    ) -> None:
        """后台读取引用关系，并切回主线程更新列表。"""

        if cancel_event.is_set() or self.is_closing:
            return
        try:
            references = find_sheet_reference_matches(
                target.file_path,
                target.sheet_name or "",
                config,
                cancel_event=cancel_event,
            )
        except ExcelReadError as exc:
            message = str(exc)
            self._post_to_ui(
                lambda item_id=item_id, message=message, generation=generation, cancel_event=cancel_event: (
                    self._references_failed(item_id, message, generation, cancel_event)
                )
            )
            return
        except Exception as exc:
            message = f"发生未预期错误：{exc}"
            self._post_to_ui(
                lambda item_id=item_id, message=message, generation=generation, cancel_event=cancel_event: (
                    self._references_failed(item_id, message, generation, cancel_event)
                )
            )
            return

        if cancel_event.is_set() or self.is_closing:
            return
        self._post_to_ui(
            lambda item_id=item_id, file_path=target.file_path, references=references, config=config,
            generation=generation, cancel_event=cancel_event: self._references_finished(
                item_id,
                file_path,
                references,
                config,
                generation,
                cancel_event,
            )
        )

    def _reference_job_is_current(self, item_id: str, generation: int, cancel_event: Event) -> bool:
        """判断引用查找结果是否仍对应当前列表中的原始工作表行。"""

        return (
            not self.is_closing
            and generation == self.result_generation
            and self.reference_cancel_events.get(item_id) is cancel_event
            and not cancel_event.is_set()
            and self.result_tree.exists(item_id)
        )

    def _references_failed(
        self,
        item_id: str,
        message: str,
        generation: int,
        cancel_event: Event,
    ) -> None:
        """引用查找失败时恢复按钮状态并显示错误。"""

        if not self._reference_job_is_current(item_id, generation, cancel_event):
            return
        self.reference_cancel_events.pop(item_id, None)
        self.reference_searching_items.discard(item_id)
        button = self.action_buttons.get(item_id)
        if button is not None:
            button.configure(text="查找引用", state="normal")
        self.status_label.configure(text=f"查找引用失败：{message}")

    def _references_finished(
        self,
        item_id: str,
        file_path: Path,
        references: list[ExcelSheetReference],
        config: ReferenceLookupConfig,
        generation: int,
        cancel_event: Event,
    ) -> None:
        """把查找到的引用表插入到源工作表行下方。"""

        if not self._reference_job_is_current(item_id, generation, cancel_event):
            return
        self.reference_cancel_events.pop(item_id, None)
        self.reference_searching_items.discard(item_id)
        button = self.action_buttons.get(item_id)
        if button is not None:
            button.configure(text="查找引用", state="normal")
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

    def _action_button_is_available(self, item_id: str) -> bool:
        """判断工作表行对应的查找引用按钮当前是否可用。"""

        button = self.action_buttons.get(item_id)
        return (
            button is not None
            and self.result_tree.exists(item_id)
            and item_id not in self.reference_searching_items
            and str(button.cget("state")) != "disabled"
        )

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

        def matches(value: str) -> bool:
            normalized_value = value.lower()
            return normalized_value == keyword if self.current_exact_match else keyword in normalized_value

        workbook_matches = any(
            matches(value)
            for value in (result.workbook_name, result.path.stem, str(result.path))
        )
        if workbook_matches:
            return self._filter_enabled_sheet_names(result.sheet_names)

        return self._filter_enabled_sheet_names(
            [sheet_name for sheet_name in result.sheet_names if matches(sheet_name)]
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
        background_image_path = cache_data.get("background_image_path", "")
        if isinstance(background_image_path, str):
            self._set_background_image_path(background_image_path, show_error=False)
        self.only_enabled_sheets_var.set(bool(cache_data.get("only_enabled_sheets", False)))
        self.only_enabled_data_var.set(bool(cache_data.get("only_enabled_data", False)))
        self.only_specified_suffixes_var.set(bool(cache_data.get("only_specified_suffixes", False)))
        self.search_folders_var.set(bool(cache_data.get("search_folders", False)))
        self.exact_match_var.set(bool(cache_data.get("exact_match", False)))
        file_suffixes = cache_data.get("file_suffixes", "")
        if isinstance(file_suffixes, str):
            self.file_suffixes_var.set(file_suffixes)

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

        cache_saved = save_cache_data(
            {
                "folder_path": folder_path,
                "folder_history": self.folder_history,
                "keyword": self.keyword_entry.get().strip(),
                "reference_config": self.reference_config,
                "disabled_sheet_marker": self.disabled_sheet_marker,
                "data_filter_config": self.data_filter_config,
                "help_url": self.help_url,
                "background_image_path": self.background_image_path,
                "only_enabled_sheets": self.only_enabled_sheets_var.get(),
                "only_enabled_data": self.only_enabled_data_var.get(),
                "only_specified_suffixes": self.only_specified_suffixes_var.get(),
                "file_suffixes": self.file_suffixes_var.get().strip(),
                "search_folders": self.search_folders_var.get(),
                "exact_match": self.exact_match_var.get(),
            }
        )
        if not cache_saved and not self.is_closing:
            self.status_label.configure(text="缓存保存失败，请检查当前用户目录权限")

    def close_app(self) -> None:
        """关闭应用前请求取消后台检索并保存当前输入。"""

        if self.is_closing:
            return
        self.is_closing = True
        if self.search_cancel_event is not None:
            self.search_cancel_event.set()
        for cancel_event in self.reference_cancel_events.values():
            cancel_event.set()
        self.reference_cancel_events.clear()
        if self.ui_poll_job is not None:
            try:
                self.after_cancel(self.ui_poll_job)
            except Exception:
                pass
            self.ui_poll_job = None
        self.search_settings_dialog.close()
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
    background_image_path = raw_data.get("background_image_path", "")
    file_suffixes = raw_data.get("file_suffixes", "")
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
        "background_image_path": background_image_path if isinstance(background_image_path, str) else "",
        "only_enabled_sheets": bool(raw_data.get("only_enabled_sheets", False)),
        "only_enabled_data": bool(raw_data.get("only_enabled_data", False)),
        "only_specified_suffixes": bool(raw_data.get("only_specified_suffixes", False)),
        "file_suffixes": file_suffixes if isinstance(file_suffixes, str) else "",
        "search_folders": bool(raw_data.get("search_folders", False)),
        "exact_match": bool(raw_data.get("exact_match", False)),
    }


def read_cache_file() -> dict[str, object]:
    """读取原始 JSON 缓存文件，读取失败时返回空字典。"""

    return read_shared_cache_data()


def save_cache_data(cache_data: dict[str, object]) -> bool:
    """保存缓存数据，同时保留工作表名称缓存等其他键。"""

    reference_config = cache_data.get("reference_config", DEFAULT_REFERENCE_LOOKUP_CONFIG)
    if not isinstance(reference_config, ReferenceLookupConfig):
        reference_config = DEFAULT_REFERENCE_LOOKUP_CONFIG
    data_filter_config = cache_data.get("data_filter_config", DEFAULT_DATA_FILTER_CONFIG)
    if not isinstance(data_filter_config, DataFilterConfig):
        data_filter_config = DEFAULT_DATA_FILTER_CONFIG

    updates: dict[str, object] = {
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


def format_excel_cell_address(row_index: int, column_index: int) -> str:
    """把行列序号转换为 Excel 地址，例如第 21 行第 1 列显示为 A21。"""

    if row_index < 1 or column_index < 1:
        return ""

    letters: list[str] = []
    remaining_column = column_index
    while remaining_column:
        remaining_column, remainder = divmod(remaining_column - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters)) + str(row_index)


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
        timeout=EXCEL_OPEN_TIMEOUT_SECONDS,
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
