"""应用主界面和用户交互逻辑。

这个文件负责窗口布局、缓存输入、启动后台检索线程、切换表格或缩略图结果视图、
打开文件/工作表/单元格，以及在工作表结果中查找引用关系。
"""

import os
import subprocess
import sys
import tempfile
import webbrowser
from collections.abc import Callable, Iterator
from time import perf_counter
from dataclasses import dataclass, replace
from pathlib import Path
from queue import Empty, Queue
from threading import Event, Thread
from tkinter import Frame as TkFrame, Menu, filedialog, font as tkfont, messagebox, ttk

import customtkinter as ctk
import ui_theme as ui
from PIL import Image, ImageDraw, ImageOps, UnidentifiedImageError

from app_preferences import (
    DEFAULT_DISABLED_SHEET_MARKER,
    load_cache_data,
    normalize_folder_history,
    normalize_help_url,
    save_cache_data,
)
from cache_store import get_app_directory, get_user_cache_directory
from cell_search import (
    DEFAULT_DATA_FILTER_CONFIG,
    DataFilterConfig,
    ExcelCellMatch,
    ExcelCellSearchIssue,
    search_excel_cells,
)
from excel_common import EXCEL_EXTENSIONS, ExcelReadError
from desktop_open import OpenResult, open_target as open_desktop_target
from search_status import SearchSnapshot, SearchTracker
from file_search import (
    FileSearchError,
    FileSearchResult,
    normalize_file_suffixes,
    search_files,
)
from json_validation import (
    DEFAULT_MAX_JSON_VALIDATION_ERRORS,
    JsonValidationReport,
    validate_excel_json_columns,
)
from popup_gui import SearchSettingsDialog, SettingsDialog
from worksheet_search import (
    DEFAULT_REFERENCE_LOOKUP_CONFIG,
    ExcelSheetReference,
    ExcelWorkbookInfo,
    ReferenceLookupConfig,
    find_sheet_reference_matches,
    scan_excel_workbooks,
)


def get_resource_path(relative_path: str) -> Path:
    """获取资源文件路径，兼容 PyInstaller 临时解包目录。"""

    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        return Path(bundle_dir) / relative_path
    return get_app_directory() / relative_path


APP_NAME = "奶龙检索大师"
APP_ICON_RELATIVE_PATH = "assets/nailong_search_master.ico"
WORKSHEET_RESULT_COLUMNS = (
    ("workbook", "工作簿", 360, 220, True),
    ("sheet", "工作表", 360, 220, True),
    ("status", "启用状态", 150, 110, False),
    ("action", "操作", 132, 120, False),
)
CELL_RESULT_COLUMNS = (
    ("value", "命中内容", 500, 180, True),
)
FILE_TABLE_RESULT_COLUMNS = (
    ("file_name", "文件/文件夹名", 320, 180, True),
    ("item_type", "类型", 150, 120, False),
    ("file_address", "所在位置", 620, 320, True),
)
JSON_RESULT_COLUMNS = (
    ("sheet", "工作表", 240, 160, True),
    ("cell", "单元格", 110, 90, False),
    ("value", "JSON 内容", 360, 200, True),
    ("message", "错误原因", 320, 180, True),
)
IMAGE_FILE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
BACKGROUND_IMAGE_FILE_NAME = "software_background.png"
BACKGROUND_RENDER_MAX_SIZE = (2560, 1440)
TRANSLUCENT_STRIP_OPACITY = 0.85
TRANSLUCENT_STRIP_COLOR = tuple(int(ui.TOOLBAR_BG[index:index + 2], 16) for index in (1, 3, 5))


@dataclass(frozen=True)
class ResultTarget:
    """结果列表项对应的打开目标。"""

    file_path: Path
    sheet_name: str | None = None
    row_index: int | None = None
    column_index: int | None = None
    open_containing_folder: bool = False
    is_directory: bool = False


def get_open_action_label(target: ResultTarget | None) -> str:
    """按钮与菜单共用同一命名规则，选中变化时不读取磁盘。"""
    if target is None:
        return "打开"
    if target.is_directory:
        return "进入文件夹"
    if target.sheet_name:
        return "在 Excel 中定位"
    if target.file_path.suffix.lower() in EXCEL_EXTENSIONS:
        return "打开 Excel"
    return "打开文件"


class ExcelEditApp(ctk.CTk):
    """主窗口类，管理所有控件状态和用户操作。"""

    def __init__(self) -> None:
        """初始化窗口状态、构建界面并加载上次缓存。"""

        super().__init__()

        self.title(APP_NAME)
        self._set_app_icon()
        self.geometry("1440x900")
        self.minsize(1080, 680)

        self.item_targets: dict[str, ResultTarget] = {}
        self.reference_items_by_source: dict[str, list[str]] = {}
        self.reference_searching_items: set[str] = set()
        self.action_buttons: dict[str, ctk.CTkButton] = {}
        self.action_layout_job: str | None = None
        self.action_item_ids: set[str] = set()
        self.tree_render_job: str | None = None
        self.tree_render_serial = 0
        self.search_tracker: SearchTracker | None = None
        self.current_result_columns: tuple[str, ...] = ()
        self.file_thumbnail_results: list[FileSearchResult] = []
        self.file_thumbnail_cards: list[FileSearchResult] = []
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
        self.folder_dropdown_active = False
        self.folder_history_refresh_pending = False
        self.search_mode_var = ctk.StringVar(value="工作表检索")
        self.only_enabled_sheets_var = ctk.BooleanVar(value=False)
        self.only_enabled_data_var = ctk.BooleanVar(value=False)
        self.only_specified_suffixes_var = ctk.BooleanVar(value=False)
        self.file_suffixes_var = ctk.StringVar(value="")
        self.search_folders_var = ctk.BooleanVar(value=False)
        self.exact_match_var = ctk.BooleanVar(value=False)
        self.current_keyword = ""
        self.current_exact_match = False
        self.pinned_folders: list[str] = []
        self.preview_enabled = True
        self.selected_target: ResultTarget | None = None
        self.item_raw_values: dict[str, str] = {}
        self.result_source: tuple[str, object] | None = None
        self.local_filter_var = ctk.StringVar(value="")
        self.local_filter_job: str | None = None
        self.local_sheet_scope: tuple[Path, str] | None = None
        self._filter_rendering = False
        self.result_sheet_enabled_only = False
        self.result_disabled_marker = DEFAULT_DISABLED_SHEET_MARKER
        self.density_var = ctk.StringVar(value="舒适")
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

        if self.is_closing or self.background_layout_job is not None:
            return
        self.background_layout_job = self.after(120, self._layout_background_image)

    def _on_app_configure(self, _event: object) -> None:
        """窗口尺寸变化时，同步背景图和半透明条的尺寸。"""

        if self.is_closing or getattr(_event, "widget", None) is not self:
            return
        self._schedule_background_image_layout()
        self._schedule_translucent_strip_layout()

    def _layout_background_image(self) -> None:
        """将背景图按窗口尺寸等比裁切填充。"""

        self.background_layout_job = None
        if self.is_closing:
            return
        if self.background_image_label is None:
            return
        if self.background_source_image is None:
            self.background_fitted_image = None
            self.background_render_size = None
            if self.background_display_image is not None:
                self.background_image_label.configure(image="")
                self.background_display_image = None
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
        frame.bind("<Configure>", lambda _event: self._schedule_translucent_strip_layout(), add="+")
        self._schedule_translucent_strip_layout()

    def _schedule_translucent_strip_layout(self) -> None:
        """合并半透明条的重绘请求。"""

        if self.is_closing or self.translucent_strip_layout_job is not None:
            return
        self.translucent_strip_layout_job = self.after(120, self._layout_translucent_strips)

    def _layout_translucent_strips(self) -> None:
        """以背景图为底绘制 85% 不透明的浅色条。"""

        self.translucent_strip_layout_job = None
        if self.is_closing:
            return
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

            if self.background_fitted_image is None:
                if name in self.translucent_strip_images:
                    overlay.configure(image="")
                    self.translucent_strip_images.pop(name)
                if overlay.winfo_manager() == "place":
                    overlay.place_forget()
                self.translucent_strip_layout_keys.pop(name, None)
                continue

            layout_key = (x, y, width, height, id(self.background_fitted_image))
            if self.translucent_strip_layout_keys.get(name) == layout_key:
                continue
            overlay.configure(width=width, height=height)
            overlay.place(x=0, y=0)
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
        from main_layout import build_main_layout
        build_main_layout(self)
        self._configure_result_columns(WORKSHEET_RESULT_COLUMNS)
        self._refresh_filter_chips()

    def _configure_tree_style(self) -> None:
        """统一结果列表的行高和字体样式。"""

        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview", rowheight=36, font=("Microsoft YaHei UI", 10),
                        background=ui.TABLE_BG, fieldbackground=ui.TABLE_BG, foreground=ui.TEXT,
                        borderwidth=0, relief="flat")
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9), background=ui.HEADER_BG,
                        foreground=ui.HEADING_TEXT, relief="flat", padding=(8, 9))
        style.map("Treeview", background=[("selected", ui.SELECTION_BG)], foreground=[("selected", ui.SELECTION_TEXT)])
        style.map("Treeview.Heading", background=[("active", ui.HOVER_BG)])
        for name in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            style.configure(name, background=ui.HEADER_BG, troughcolor=ui.TABLE_BG,
                            bordercolor=ui.TABLE_BG, arrowcolor=ui.HEADING_TEXT,
                            lightcolor=ui.BORDER, darkcolor=ui.BORDER)
            style.map(name, background=[("active", ui.HOVER_BG)])

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
            fg_color=ui.ACCENT,
            hover_color=ui.ACCENT_HOVER,
            text_color=ui.ON_ACCENT,
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

        line = TkFrame(self.result_list_frame, bg=ui.DIVIDER, highlightthickness=0, bd=0)
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
        """工作表的查找引用按钮只在可见行分配；打开操作集中在底部。"""
        self.action_layout_job = None
        visible = set(self._get_visible_result_item_ids()) if "action" in self.current_result_columns else set()
        spare = []
        for item_id in list(self.action_buttons):
            if item_id not in visible:
                button = self.action_buttons.pop(item_id)
                button.place_forget()
                spare.append(button)
        for item_id in visible & self.action_item_ids:
            bbox = self.result_tree.bbox(item_id, "action")
            if not bbox:
                continue
            if item_id not in self.action_buttons:
                if spare:
                    self.action_buttons[item_id] = spare.pop()
                else:
                    self._create_action_button(item_id)
            button = self.action_buttons[item_id]
            searching = item_id in self.reference_searching_items
            button.configure(text="查找中..." if searching else "查找引用",
                             state="disabled" if searching else "normal",
                             command=lambda item=item_id: self._click_reference_button(item))
            x, y, width, height = bbox
            button_width = min(104, max(86, width - 12))
            button_height = min(24, max(20, height - 6))
            button.configure(width=button_width, height=button_height)
            button.place(x=x + max(4, (width - button_width) // 2), y=y + max(3, (height - button_height) // 2))
        for button in spare:
            button.destroy()

    def _refresh_filter_chips(self) -> None:
        from main_layout import MODES
        for child in self.filter_chips_frame.winfo_children():
            child.destroy()
        ctk.CTkLabel(self.filter_chips_frame, text="当前条件", text_color=ui.MUTED, width=64).pack(side="left", padx=(0, 8))

        def chip(text, command):
            button = ctk.CTkButton(self.filter_chips_frame, text=text, command=command, width=20, height=27,
                corner_radius=5, fg_color=ui.INSET_BG, hover_color=ui.HOVER_BG, text_color=ui.SELECTION_TEXT,
                border_width=1, border_color=ui.BORDER, font=("Microsoft YaHei UI", 11))
            button.pack(side="left", padx=(0, 8))
            if self.search_cancel_event is not None:
                button.configure(state="disabled")

        mode = self.search_mode_var.get()
        self.mode_display_var.set(next((label for label, value in MODES.items() if value == mode), "工作表"))
        if mode in {"工作表检索", "单元格检索"}:
            ctk.CTkLabel(self.filter_chips_frame, text="仅 Excel 文件", height=27, corner_radius=5,
                fg_color=ui.INSET_BG, text_color=ui.MUTED, font=("Microsoft YaHei UI", 11)).pack(side="left", padx=(0, 8))
        if mode != "JSON格式检查":
            chip("精确匹配 ▾" if self.exact_match_var.get() else "包含匹配 ▾", lambda: self._change_filter(self.exact_match_var))
        if mode in {"工作表检索", "单元格检索"} and self.only_enabled_sheets_var.get():
            chip("仅启用表格 ×", lambda: self._change_filter(self.only_enabled_sheets_var, False))
        if mode == "单元格检索":
            if self.only_enabled_data_var.get():
                chip("仅启用数据 ×", lambda: self._change_filter(self.only_enabled_data_var, False))
            config = self.data_filter_config
            if config.enable_header_filter:
                chip(f"跳过前 {config.header_row_count} 行", self.open_settings)
            if config.enable_blank_header_column_filter:
                chip("屏蔽空白表头列", self.open_settings)
        if mode == "文件检索":
            if self.only_specified_suffixes_var.get():
                suffix = self.file_suffixes_var.get().strip()
                chip(f"后缀：{suffix[:24] or '未设置'} ×", lambda: self._change_filter(self.only_specified_suffixes_var, False))
            if self.search_folders_var.get():
                chip("包含文件夹 ×", lambda: self._change_filter(self.search_folders_var, False))
        if mode == "JSON格式检查":
            chip("非 $ 工作表 · 第 4、6 行识别 JSON 列", self.open_search_settings)
        self.search_settings_button = ctk.CTkButton(self.filter_chips_frame, text="＋ 更多条件", width=90, height=27,
            fg_color="transparent", hover_color=ui.HOVER_BG, text_color=ui.ACCENT_HOVER, command=self.open_search_settings)
        self.search_settings_button.pack(side="left", padx=(0, 8))
        self.reset_filters_button = ctk.CTkButton(self.filter_chips_frame, text="重置", width=44, height=27,
            fg_color="transparent", hover_color=ui.HOVER_BG, text_color=ui.MUTED, command=self._reset_search_filters)
        self.reset_filters_button.pack(side="left")
        if self.search_cancel_event is not None:
            self.search_settings_button.configure(state="disabled")
            self.reset_filters_button.configure(state="disabled")

    def _change_filter(self, variable, value=None) -> None:
        if self.search_cancel_event is not None:
            return
        variable.set(not variable.get() if value is None else value)
        self._refresh_filter_chips()
        self._save_cache_from_inputs()
        self.status_label.configure(text="条件已更新，下次检索生效")

    def _reset_search_filters(self) -> None:
        if self.search_cancel_event is not None:
            return
        for variable in (self.exact_match_var, self.only_enabled_sheets_var, self.only_enabled_data_var,
                         self.only_specified_suffixes_var, self.search_folders_var):
            variable.set(False)
        self._refresh_filter_chips()
        self._save_cache_from_inputs()
        self.status_label.configure(text="检索条件已重置，表头规则可在设置中调整")

    def _toggle_pinned_folder(self) -> None:
        folder = self.folder_path_var.get().strip()
        if not folder or not Path(folder).is_dir():
            self.status_label.configure(text="先选择有效目录，再固定到历史列表顶部")
            return
        if folder in self.pinned_folders:
            self.pinned_folders.remove(folder)
        else:
            self.pinned_folders = [folder, *self.pinned_folders][:10]
        self._refresh_folder_history_menu()
        self._refresh_pin_button()
        self._save_cache_from_inputs()

    def _refresh_pin_button(self) -> None:
        pinned = self.folder_path_var.get().strip() in self.pinned_folders
        self.pin_folder_button.configure(text="★ 已固定" if pinned else "☆ 固定目录",
                                         text_color="#96791e" if pinned else ui.TEXT)

    def _set_density(self, _value=None) -> None:
        ttk.Style(self).configure("Treeview", rowheight=28 if self.density_var.get() == "紧凑" else 36)
        self._schedule_action_button_layout()
        self._schedule_result_grid_lines()
        self._save_cache_from_inputs()

    def _toggle_preview(self) -> None:
        self.preview_enabled = not self.preview_enabled
        self._sync_preview_visibility()
        self._save_cache_from_inputs()

    def _sync_preview_visibility(self) -> None:
        if self.preview_enabled:
            self.preview_pane.grid()
            self.preview_toggle_button.configure(text="预览 ▸")
            self._update_selected_preview()
        else:
            self.preview_pane.grid_remove()
            self.preview_pane.clear()
            self.preview_toggle_button.configure(text="◂ 预览")

    def _on_result_selected(self, _event=None) -> None:
        selection = self.result_tree.selection()
        if not selection:
            self.selected_target = None
            self.selected_value = ""
            self.selection_label.configure(text="选择结果可查看详情")
            self._sync_result_actions()
            self.preview_pane.clear()
            return
        item = selection[0]
        target = self.item_targets.get(item)
        if target is None:
            return
        self.selected_target = target
        self.selected_value = self.item_raw_values.get(item, "\t".join(str(v) for v in self.result_tree.item(item, "values") if v))
        self.selection_label.configure(text=(f"已选择 {format_excel_cell_address(target.row_index, target.column_index)} · {target.sheet_name}"
                                             if target.row_index and target.column_index else f"已选择：{target.sheet_name or target.file_path.name}"))
        self._sync_result_actions()
        self._update_selected_preview()

    def _on_file_selected(self, result: FileSearchResult) -> None:
        self.selected_target = ResultTarget(result.path, is_directory=result.is_directory)
        self.selected_value = result.file_name
        self.selection_label.configure(text=f"已选择：{result.file_name}")
        self._sync_result_actions()
        self._update_selected_preview()

    def _update_selected_preview(self) -> None:
        target = self.selected_target
        if not self.preview_enabled or target is None:
            return
        if target.sheet_name:
            self.preview_pane.show_cell(target.file_path, target.sheet_name, target.row_index or 1,
                target.column_index or 1, getattr(self, "selected_value", ""), self.current_keyword)
        else:
            self.preview_pane.show_file(target.file_path)

    def _open_file_result(self, result: FileSearchResult) -> None:
        target = ResultTarget(result.path, is_directory=result.is_directory)
        if self.selected_target != target:
            self._on_file_selected(result)
        self._open_target(target)

    def _sync_result_actions(self) -> None:
        """固定操作栏只反映当前选择，进行中的相同操作不能重复点击。"""
        target = self.selected_target
        if target is None:
            self.selection_open_button.configure(text="打开", state="disabled")
            self.selection_directory_button.configure(text="打开所在目录", state="disabled")
            return
        primary = replace(target, file_path=Path(os.path.abspath(os.fspath(target.file_path))))
        directory = ResultTarget(primary.file_path, open_containing_folder=True)
        opening = primary in self.opening_targets
        revealing = directory in self.opening_targets
        self.selection_open_button.configure(text="打开中…" if opening else get_open_action_label(target),
                                             state="disabled" if opening else "normal")
        self.selection_directory_button.configure(text="定位中…" if revealing else "打开所在目录",
                                                  state="disabled" if revealing else "normal")

    def _open_selected_target(self) -> None:
        if self.selected_target:
            self._open_target(self.selected_target)

    def _open_selected_directory(self) -> None:
        if self.selected_target:
            self._open_target(ResultTarget(self.selected_target.file_path, open_containing_folder=True))

    def _copy_selected(self, kind="value") -> str:
        target = self.selected_target
        if target is None:
            return "break"
        if kind == "path":
            value = str(target.file_path)
        elif kind == "address":
            if not target.row_index or not target.column_index:
                return "break"
            value = f"'{(target.sheet_name or '').replace(chr(39), chr(39)*2)}'!{format_excel_cell_address(target.row_index, target.column_index)}"
        else:
            value = getattr(self, "selected_value", "")
        self.clipboard_clear()
        self.clipboard_append(value)
        self.status_label.configure(text="已复制" + {"path": "文件路径", "address": "单元格地址"}.get(kind, "完整内容"))
        return "break"

    def _show_result_menu(self, event) -> str:
        item = self.result_tree.identify_row(event.y)
        if not item:
            return "break"
        self.result_tree.selection_set(item)
        self.result_tree.focus(item)
        self._on_result_selected()
        self._popup_result_menu(event)
        return "break"

    def _show_file_menu(self, result, event) -> None:
        self._on_file_selected(result)
        self._popup_result_menu(event)

    def _popup_result_menu(self, event) -> None:
        target = self.selected_target
        if target is None:
            return
        menu = Menu(self, tearoff=False, background=ui.SURFACE, foreground=ui.TEXT, activebackground=ui.HOVER_BG, activeforeground=ui.ACCENT, font=("Microsoft YaHei UI", 10))
        menu.add_command(label="复制内容", accelerator="Ctrl+C", command=self._copy_selected)
        if target.row_index:
            menu.add_command(label="复制单元格地址", command=lambda: self._copy_selected("address"))
        menu.add_command(label="复制文件路径", command=lambda: self._copy_selected("path"))
        menu.add_separator()
        menu.add_command(label=get_open_action_label(target), accelerator="Enter", command=self._open_selected_target)
        menu.add_command(label="打开所在目录", command=self._open_selected_directory)
        if target.sheet_name:
            menu.add_separator()
            menu.add_command(label="只看此工作表", command=lambda: self._scope_to_sheet(target))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
            menu.destroy()

    def _scope_to_sheet(self, target) -> None:
        self.local_sheet_scope = (target.file_path, target.sheet_name)
        self._apply_local_filter()

    def _focus_local_filter(self, _event=None) -> str:
        self.local_filter_entry.focus_set()
        self.local_filter_entry.select_range(0, "end")
        return "break"

    def _reset_local_filter(self) -> str:
        self.local_sheet_scope = None
        self.local_filter_var.set("")
        self._schedule_local_filter()
        return "break"

    def _schedule_local_filter(self) -> None:
        if self.local_filter_job is not None:
            self.after_cancel(self.local_filter_job)
        self.local_filter_job = self.after(180, self._apply_local_filter)

    def _filtered_result_source(self):
        if self.result_source is None:
            return None
        kind, data = self.result_source
        query = self.local_filter_var.get().strip().casefold()
        scope = self.local_sheet_scope
        if kind == "cells":
            data = [m for m in data if (not query or query in f"{m.path} {m.sheet_name} {format_excel_cell_address(m.row_index, m.column_index)} {m.value}".casefold())
                    and (scope is None or (m.path, m.sheet_name) == scope)]
        elif kind == "files":
            data = [r for r in data if not query or query in str(r.path).casefold()]
        elif kind == "json":
            data = replace(data, errors=[e for e in data.errors if (not query or query in f"{e.path} {e.sheet_name} {format_excel_cell_address(e.row_index, e.column_index)} {e.value} {e.message}".casefold())
                                        and (scope is None or (e.path, e.sheet_name) == scope)])
        return kind, data

    def _apply_local_filter(self) -> None:
        if self.local_filter_job is not None:
            self.after_cancel(self.local_filter_job)
        self.local_filter_job = None
        if self.search_cancel_event is not None:
            return
        filtered = self._filtered_result_source()
        if filtered is None:
            return
        kind, data = filtered
        self._filter_rendering = True
        try:
            {"cells": self._render_cell_matches, "files": self._render_file_results,
             "sheets": self._render_results, "json": self._render_json_validation_results}[kind](data)
        finally:
            self._filter_rendering = False

    def _record_result_source(self, kind, data, count, summary) -> None:
        if not self._filter_rendering:
            self.result_source = (kind, data)
        self.result_count_label.configure(text=f"检索结果  {count}")
        narrowed = self.local_filter_var.get().strip() or self.local_sheet_scope
        suffix = " · 结果内筛选中，Esc 恢复全部" if narrowed else ""
        self.result_summary_label.configure(text=summary + suffix)

    def _export_results(self) -> None:
        from result_export import export_result_rows
        filtered = self._filtered_result_source()
        if filtered is None or self.search_cancel_event is not None or self.tree_render_job is not None:
            self.status_label.configure(text="请等待检索及结果显示完成后导出")
            return
        destination = filedialog.asksaveasfilename(parent=self, title="导出当前筛选结果", defaultextension=".xlsx", filetypes=[("Excel 工作簿", "*.xlsx")], initialfile="检索结果.xlsx")
        if not destination:
            return
        kind, data = filtered
        headers = ("文件路径", "工作表", "单元格", "内容", "说明")
        if kind == "cells":
            rows = [(str(m.path), m.sheet_name, format_excel_cell_address(m.row_index, m.column_index), m.value, "") for m in data]
        elif kind == "files":
            rows = [(str(r.path), "", "", r.file_name, "文件夹" if r.is_directory else r.path.suffix) for r in data]
        elif kind == "json":
            rows = [(str(e.path), e.sheet_name, format_excel_cell_address(e.row_index, e.column_index), e.value, e.message) for e in data.errors]
        else:
            rows = [(str(r.path), sheet if not description else "", "", "", description or status)
                    for r, sheet, status, description in self._worksheet_display_rows(data)]
        self.status_label.configure(text=f"正在导出 {len(rows)} 条结果…")
        def worker():
            try:
                count = export_result_rows(Path(destination), headers, rows)
            except Exception as exc:
                self._post_to_ui(lambda message=str(exc): messagebox.showerror("导出失败", message, parent=self))
            else:
                self._post_to_ui(lambda: self.status_label.configure(text=f"已导出 {count} 条结果：{destination}"))
        Thread(target=worker, daemon=True).start()

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

        updated_history = normalize_folder_history(self.folder_history, str(folder_path))
        if updated_history == self.folder_history:
            return

        self.folder_history = updated_history
        self._refresh_folder_history_menu()

    def _refresh_folder_history_menu(self) -> None:
        """刷新文件夹历史下拉框中的选项。"""

        if self.folder_dropdown_active:
            self.folder_history_refresh_pending = True
            return

        self.folder_entry.configure(values=list(dict.fromkeys([*self.pinned_folders, *self.folder_history])))
        self._refresh_pin_button()
        self._sync_folder_dropdown_width()

    def _sync_folder_dropdown_width(self) -> None:
        """让历史下拉列表宽度尽量贴近文件夹输入框宽度。"""

        if self.folder_dropdown_active:
            self.folder_history_refresh_pending = True
            return

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
        self.folder_dropdown_active = True
        try:
            dropdown_menu.tk_popup(menu_x, menu_y)
        finally:
            self.folder_dropdown_active = False
            if self.folder_history_refresh_pending and not self.is_closing:
                self.folder_history_refresh_pending = False
                self._refresh_folder_history_menu()

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
        for folder_path in dict.fromkeys([*self.pinned_folders, *self.folder_history]):
            label = ("★  " if folder_path in self.pinned_folders else "") + folder_path
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
        self._refresh_filter_chips()
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
        self._refresh_filter_chips()
        self._save_cache_from_inputs()
        self.status_label.configure(text="检索设置已保存")

    def run_selected_search(self) -> None:
        """根据下拉框当前选项分发到对应检索功能。"""

        if self.search_cancel_event is not None:
            return
        search_mode = self.search_mode_var.get()
        if search_mode == "单元格检索":
            self.start_cell_search()
            return
        if search_mode == "文件检索":
            self.start_file_search()
            return
        if search_mode == "JSON格式检查":
            self.start_json_validation()
            return
        self.start_scan()

    def _on_search_mode_changed(self, selected_mode: str) -> None:
        """切换检索类型，并同步已打开的检索设置窗口。"""

        self.search_mode_var.set(selected_mode)
        self._refresh_filter_chips()
        self.search_settings_dialog.refresh(selected_mode)
        if selected_mode == "JSON格式检查":
            self.keyword_entry.configure(placeholder_text="JSON 格式检查无需输入内容：自动检查第 4、6 行识别出的 JSON 列")
        else:
            placeholders = {"工作表检索": "输入工作簿或工作表关键词 · Enter 检索", "单元格检索": "输入单元格内容 · Enter 检索", "文件检索": "输入文件名或路径关键词 · Enter 检索"}
            self.keyword_entry.configure(placeholder_text=placeholders.get(selected_mode, "输入关键词"))
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
        if not self.current_keyword:
            messagebox.showinfo("缺少搜索内容", "请先在输入框里填写要搜索的内容。")
            return
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
        if not keyword:
            messagebox.showinfo("缺少搜索内容", "请先在输入框里填写要搜索的内容。")
            return
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

    def start_json_validation(self) -> None:
        """检查非 `$` 工作表中由第 4、6 行识别出的 JSON 数据列。"""

        folder_text = self.folder_entry.get().strip()
        if not folder_text:
            messagebox.showinfo("缺少文件夹", "请先选择或输入一个文件夹地址。")
            return

        folder_path = Path(folder_text)
        if not folder_path.exists() or not folder_path.is_dir():
            messagebox.showerror("文件夹无效", "请输入有效的文件夹地址。")
            return

        self._remember_folder_path(folder_path)
        self._save_cache_from_inputs()
        cancel_event = self._begin_search()
        self.status_label.configure(text="正在检查非 $ 工作表中的 JSON 数据列，请稍等")
        self._clear_results()

        worker = Thread(
            target=self._validate_json_in_background,
            args=(folder_path, cancel_event),
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

        tracker = self.search_tracker
        try:
            results = scan_excel_workbooks(
                folder_path,
                keyword,
                cancel_event=cancel_event,
                exact_match=exact_match,
                tracker=tracker,
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
            self._post_search_callback(cancel_event, lambda: self._search_cancelled(lambda: self._render_results(results)))
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
        tracker = self.search_tracker
        try:
            matches = search_excel_cells(
                folder_path,
                keyword,
                disabled_sheet_marker=disabled_marker,
                data_filter_config=data_filter_config,
                cancel_event=cancel_event,
                issues=issues,
                exact_match=exact_match,
                tracker=tracker,
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
            self._post_search_callback(cancel_event, lambda: self._search_cancelled(lambda: self._render_cell_matches(matches)))
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

        tracker = self.search_tracker
        try:
            results = search_files(
                folder_path,
                keyword,
                cancel_event=cancel_event,
                suffixes=suffixes,
                exact_match=exact_match,
                include_folders=include_folders,
                tracker=tracker,
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
            self._post_search_callback(cancel_event, lambda: self._search_cancelled(lambda: self._render_file_results(results)))
            return
        self._post_search_callback(cancel_event, lambda results=results: self._file_search_finished(results))

    def _validate_json_in_background(self, folder_path: Path, cancel_event: Event) -> None:
        """后台检查 JSON 数据列，并把错误列表切回主线程渲染。"""

        tracker = self.search_tracker
        try:
            report = validate_excel_json_columns(
                folder_path,
                cancel_event=cancel_event,
                max_errors=DEFAULT_MAX_JSON_VALIDATION_ERRORS,
                tracker=tracker,
            )
        except ExcelReadError as exc:
            message = str(exc)
            if cancel_event.is_set():
                self._post_search_callback(cancel_event, self._search_cancelled)
            else:
                self._post_search_callback(
                    cancel_event,
                    lambda message=message: self._json_validation_failed(message),
                )
            return
        except Exception as exc:
            message = f"发生未预期错误：{exc}"
            self._post_search_callback(
                cancel_event,
                lambda message=message: self._json_validation_failed(message),
            )
            return

        if cancel_event.is_set():
            self._post_search_callback(cancel_event, lambda: self._search_cancelled(lambda: self._render_json_validation_results(report)))
            return
        self._post_search_callback(
            cancel_event,
            lambda report=report: self._json_validation_finished(report),
        )

    def _begin_search(self) -> Event:
        """创建任务和限频进度回调，旧任务的快照不能写入新界面。"""

        cancel_event = Event()
        self.search_cancel_event = cancel_event
        self.search_tracker = SearchTracker(
            lambda snapshot: self._post_search_callback(
                cancel_event, lambda snapshot=snapshot: self._show_search_progress(snapshot),
            )
        )
        self.current_keyword = self.keyword_entry.get().strip()
        self.current_exact_match = self.exact_match_var.get()
        self.failure_details_button.grid_remove()
        self._set_search_controls_running(True)
        return cancel_event

    def _show_search_progress(self, snapshot: SearchSnapshot) -> None:
        if self.search_cancel_event is None or self.search_cancel_event.is_set():
            return
        name = Path(snapshot.current_path).name
        if len(name) > 44:
            name = name[:41] + "…"
        self.status_label.configure(text=(
            f"已扫描 {snapshot.scanned_count} 项，找到 {snapshot.result_count} 条，"
            f"耗时 {snapshot.elapsed_seconds:.1f} 秒 · {name}"
        ))

    def _completion_status(self, text: str, limit_reached: bool = False) -> str:
        tracker = self.search_tracker
        parts = [text]
        if tracker is not None:
            snapshot = tracker.snapshot()
            parts.append(f"已扫描 {snapshot.scanned_count} 项，耗时 {snapshot.elapsed_seconds:.1f} 秒")
            limit_reached = snapshot.stop_reason == "limit"
            if tracker.issues:
                parts.append(f"跳过 {len(tracker.issues)} 处无法读取的文件或目录")
                self.failure_details_button.configure(text=f"失败明细 ({len(tracker.issues)})")
                self.failure_details_button.grid()
        if limit_reached:
            parts.append("达到结果上限，已停止检索，结果可能不完整")
        return "；".join(parts)

    def show_failure_details(self) -> None:
        if self.search_tracker is None or not self.search_tracker.issues:
            return
        window = ctk.CTkToplevel(self)
        window.title("检索失败明细")
        window.geometry("820x480")
        window.transient(self)
        textbox = ctk.CTkTextbox(window, wrap="word")
        textbox.pack(fill="both", expand=True, padx=12, pady=12)
        textbox.insert("1.0", "\n\n".join(
            f"{issue.path}\n{issue.message}" for issue in self.search_tracker.issues
        ))
        textbox.configure(state="disabled")

    def cancel_current_search(self) -> None:
        """响应取消按钮，请求后台线程尽快停止。"""

        if self.search_cancel_event is None:
            return

        self.search_cancel_event.set()
        if self.cancel_search_button is not None:
            self.cancel_search_button.configure(state="disabled", text="取消中...")
        self.status_label.configure(text="正在取消检索，请稍等")

    def _search_cancelled(self, partial_renderer: Callable[[], None] | None = None) -> None:
        """后台线程确认取消后恢复界面状态。"""

        if self.search_tracker is not None:
            self.search_tracker.finish("cancelled")
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        if partial_renderer is not None:
            partial_renderer()
        self.status_label.configure(text=self._completion_status("已取消检索，保留已找到的结果；结果可能不完整"))

    def _scan_failed(self, message: str) -> None:
        """显示工作表检索失败信息。"""

        if self.search_tracker is not None:
            self.search_tracker.finish("failed")
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="检索失败")
        messagebox.showerror("检索失败", message)

    def _scan_finished(self, results: list[ExcelWorkbookInfo]) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_results(results)
        row_count = self._count_display_rows(results)
        self.status_label.configure(text=self._completion_status(
            f"工作表检索完成：{len(results)} 个工作簿，{row_count} 条结果",
        ))

    def _cell_search_failed(self, message: str) -> None:
        """显示单元格检索失败信息。"""

        if self.search_tracker is not None:
            self.search_tracker.finish("failed")
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="单元格搜索失败")
        messagebox.showerror("单元格搜索失败", message)

    def _cell_search_finished(
        self, matches: list[ExcelCellMatch], keyword: str, issues: list[ExcelCellSearchIssue],
    ) -> None:
        self.current_keyword = keyword
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_cell_matches(matches)
        self.status_label.configure(text=self._completion_status(f"单元格检索完成：找到 {len(matches)} 条匹配"))

    def _file_search_failed(self, message: str) -> None:
        """显示文件检索失败信息。"""

        if self.search_tracker is not None:
            self.search_tracker.finish("failed")
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="文件检索失败")
        messagebox.showerror("文件检索失败", message)

    def _file_search_finished(self, results: list[FileSearchResult]) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_file_results(results)
        file_count = sum(not result.is_directory for result in results)
        folder_count = sum(result.is_directory for result in results)
        count_text = f"文件检索完成：找到 {file_count} 个文件"
        if folder_count:
            count_text += f"，{folder_count} 个文件夹"
        self.status_label.configure(text=self._completion_status(count_text))

    def _json_validation_failed(self, message: str) -> None:
        """显示 JSON 格式检查失败信息。"""

        if self.search_tracker is not None:
            self.search_tracker.finish("failed")
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self.status_label.configure(text="JSON 格式检查失败")
        messagebox.showerror("JSON 格式检查失败", message)

    def _json_validation_finished(self, report: JsonValidationReport) -> None:
        self.search_cancel_event = None
        self._set_search_controls_running(False)
        self._render_json_validation_results(report)
        text = (
            f"JSON 检查完成：识别 {report.json_column_count} 列，检查 {report.checked_cell_count} 个单元格，"
            f"发现 {len(report.errors)} 处格式错误"
        )
        self.status_label.configure(text=self._completion_status(text, report.limit_reached))

    def _set_search_controls_running(self, is_running: bool) -> None:
        """切换检索中/空闲状态下按钮和选项的可用性。"""

        self._refresh_filter_chips()
        self.local_filter_entry.configure(state="disabled" if is_running else "normal")
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

    def _cancel_tree_render(self) -> None:
        self.tree_render_serial += 1
        if self.tree_render_job is not None:
            self.after_cancel(self.tree_render_job)
            self.tree_render_job = None

    def _render_tree_rows(self, rows: Iterator[None]) -> None:
        self._cancel_tree_render()
        serial = self.tree_render_serial

        def insert_batch() -> None:
            self.tree_render_job = None
            if self.is_closing or serial != self.tree_render_serial:
                return
            deadline = perf_counter() + 0.008
            for _ in range(200):
                try:
                    next(rows)
                except StopIteration:
                    self._update_result_group_controls_visibility()
                    self._schedule_action_button_layout()
                    return
                if perf_counter() >= deadline:
                    break
            self._schedule_action_button_layout()
            self.tree_render_job = self.after(1, insert_batch)

        insert_batch()

    def _render_cell_matches(self, matches: list[ExcelCellMatch]) -> None:
        """按工作簿、工作表两级目录展示单元格命中。"""

        self._clear_results()
        self._record_result_source("cells", list(matches), len(matches),
            f"{len({m.path for m in matches})} 个工作簿 · {len({(m.path, m.sheet_name) for m in matches})} 张工作表")
        self._show_table_results()
        self._configure_result_columns(
            CELL_RESULT_COLUMNS,
            tree_heading="工作簿 / 工作表 / 位置",
            tree_width=260,
            tree_min_width=180,
            tree_stretch=True,
        )

        matches_by_workbook: dict[Path, list[ExcelCellMatch]] = {}
        for match in matches:
            matches_by_workbook.setdefault(match.path, []).append(match)

        def rows() -> Iterator[None]:
            for workbook_path, workbook_matches in matches_by_workbook.items():
                # 工作簿和工作表都是分类目录；只有具体单元格命中行会打开 Excel。
                workbook_item_id = self.result_tree.insert(
                    "",
                    "end",
                    text=workbook_path.name,
                    values=(f"{len(workbook_matches)} 条匹配",),
                    open=True,
                    tags=("cell-workbook",),
                )
                self.item_targets[workbook_item_id] = ResultTarget(workbook_path)

                matches_by_sheet: dict[str, list[ExcelCellMatch]] = {}
                for match in workbook_matches:
                    matches_by_sheet.setdefault(match.sheet_name, []).append(match)

                for sheet_name, sheet_matches in matches_by_sheet.items():
                    sheet_item_id = self.result_tree.insert(
                        workbook_item_id,
                        "end",
                        text=sheet_name,
                        values=(f"{len(sheet_matches)} 条匹配",),
                        open=True,
                        tags=("cell-sheet",),
                    )
                    self.item_targets[sheet_item_id] = ResultTarget(workbook_path, sheet_name)

                    for match in sheet_matches:
                        item_id = self.result_tree.insert(
                            sheet_item_id,
                            "end",
                            text=format_excel_cell_address(match.row_index, match.column_index),
                            values=(format_cell_value(match.value),),
                        )
                        self.item_targets[item_id] = ResultTarget(
                            match.path,
                            match.sheet_name,
                            match.row_index,
                            match.column_index,
                        )
                        self.item_raw_values[item_id] = str(match.value)
                        yield

        self._render_tree_rows(rows())
        self._update_result_group_controls_visibility()
        self._schedule_action_button_layout()

    def _render_json_validation_results(self, report: JsonValidationReport) -> None:
        """以表格展示 JSON 格式错误；双击错误行可打开并定位到对应单元格。"""

        self._clear_results()
        self._record_result_source("json", report, len(report.errors), f"检查 {report.checked_cell_count} 个单元格 · JSON 格式错误")
        self._show_table_results()
        self._configure_result_columns(
            JSON_RESULT_COLUMNS,
            tree_heading="工作簿",
            tree_width=260,
            tree_min_width=180,
            tree_stretch=True,
        )

        def rows() -> Iterator[None]:
            for error in report.errors:
                item_id = self.result_tree.insert(
                    "",
                    "end",
                    text=error.workbook_name,
                    values=(
                        error.sheet_name,
                        format_excel_cell_address(error.row_index, error.column_index),
                        format_cell_value(error.value),
                        error.message,
                    ),
                )
                self.item_targets[item_id] = ResultTarget(
                    error.path,
                    error.sheet_name,
                    error.row_index,
                    error.column_index,
                )
                self.item_raw_values[item_id] = str(error.value)
                yield

        self._render_tree_rows(rows())
        self._update_result_group_controls_visibility()
        self._schedule_action_button_layout()

    def _render_file_results(self, results: list[FileSearchResult]) -> None:
        """保存文件检索结果，并按当前视图渲染。"""

        results = list(results)
        self._clear_results()
        self._record_result_source("files", results, len(results), f"{sum(not r.is_directory for r in results)} 个文件 · {sum(r.is_directory for r in results)} 个文件夹")
        self.file_thumbnail_results = list(results)
        self._render_file_result_view()

        if not results:
            self.status_label.configure(text="没有找到匹配的文件")

    def _on_file_result_view_changed(self, _selected_view: str) -> None:
        """在缩略图和表格之间切换当前文件检索结果。"""

        if self.result_source is None or self.result_source[0] != "files":
            return
        self._render_file_result_view()
        self._save_cache_from_inputs()

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

        self._cancel_tree_render()
        self.result_group_controls.grid_remove()
        self.file_thumbnail_frame.clear()
        self.item_targets.clear()
        self.item_raw_values.clear()
        self.selected_target = None
        self.selected_value = ""
        self.selection_label.configure(text="选择结果可查看详情")
        self._sync_result_actions()
        self.preview_pane.clear()
        items = self.result_tree.get_children()
        if items:
            self.result_tree.delete(*items)

    def _render_file_results_as_thumbnails(self) -> None:
        """按逻辑位置展示全部分组，只绘制当前视口的卡片。"""

        self.file_thumbnail_frame.set_results(self.file_thumbnail_results)

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

        def rows() -> Iterator[None]:
            for folder_path, folder_results in grouped_results.items():
                folder_item_id = self.result_tree.insert(
                    "",
                    "end",
                    text=str(folder_path),
                    values=("", "", f"{len(folder_results)} 项"),
                    open=True,
                )
                self.item_targets[folder_item_id] = ResultTarget(folder_path, is_directory=True)
                for result in folder_results:
                    item_type = self._get_file_thumbnail_style(result)[2]
                    item_id = self.result_tree.insert(
                        folder_item_id,
                        "end",
                        text="",
                        values=(result.file_name, item_type, str(result.path)),
                    )
                    self.item_targets[item_id] = ResultTarget(result.path, is_directory=result.is_directory)
                    self.item_raw_values[item_id] = result.file_name
                    yield

        self._render_tree_rows(rows())
        self._update_result_group_controls_visibility()

    def _group_file_results_by_folder(self) -> dict[Path, list[FileSearchResult]]:
        """保持检索排序，将同一所在文件夹的结果连续分组。"""

        grouped_results: dict[Path, list[FileSearchResult]] = {}
        for result in self.file_thumbnail_results:
            grouped_results.setdefault(result.path.parent, []).append(result)
        return grouped_results

    @staticmethod
    def _get_file_thumbnail_style(result: FileSearchResult) -> tuple[str, str, str]:
        """按文件类型返回缩略图图标、主色和说明文字。"""

        if result.is_directory:
            return "📁", "#b7791f", "文件夹"

        suffix = result.path.suffix.lower()
        if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".csv"}:
            return "📊", "#217346", "Excel / 表格文件"
        if suffix in {".doc", ".docx", ".rtf", ".txt", ".md", ".pdf"}:
            return "📄", ui.ACCENT, "文档"
        if suffix in IMAGE_FILE_SUFFIXES or suffix == ".svg":
            return "🖼", "#7c3aed", "图片"
        if suffix in {".mp4", ".avi", ".mkv", ".mov", ".wmv"}:
            return "🎬", "#be185d", "视频"
        if suffix in {".mp3", ".wav", ".flac", ".aac", ".m4a"}:
            return "🎵", "#0f766e", "音频"
        if suffix in {".zip", ".rar", ".7z", ".tar", ".gz"}:
            return "🗜", "#92400e", "压缩包"
        return "📄", ui.TEXT, "文件"

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

    def _open_result_item(self, item_id: str) -> None:
        """双击和 Enter 与底部主按钮一致；分组使用展开箭头控制。"""

        target = self.item_targets.get(item_id)
        if target is None:
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

        target = replace(target, file_path=Path(os.path.abspath(os.fspath(target.file_path))))
        if target in self.opening_targets:
            return
        self.opening_targets.add(target)
        self._sync_result_actions()
        self.status_label.configure(text=f"正在打开：{target.file_path.name}")
        worker = Thread(target=self._open_target_in_background, args=(target,), daemon=True)
        worker.start()

    def _open_target_in_background(self, target: ResultTarget) -> None:
        """在后台打开目标，避免 Excel COM 或系统关联程序阻塞 Tk 主线程。"""

        try:
            result = open_desktop_target(
                target.file_path, target.sheet_name, target.row_index, target.column_index,
                reveal=target.open_containing_folder,
            )
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

        self._post_to_ui(lambda target=target, result=result: self._open_target_finished(target, result))

    def _open_target_finished(self, target: ResultTarget, result: OpenResult) -> None:
        """目标打开完成后更新状态。"""

        self.opening_targets.discard(target)
        self._sync_result_actions()
        if not result.hwnd:
            text = f"已提交打开请求，未确认目标窗口：{target.file_path.name}"
        elif result.foreground:
            action = "定位" if result.located and (target.open_containing_folder or target.sheet_name) else "打开"
            text = f"已{action}并切到前台：{target.file_path.name}"
        else:
            text = f"已打开：{target.file_path.name}；窗口当前未在前台"
        if result.hwnd and not result.located:
            text += "；未能完成目标定位"
        if result.message:
            text += f"；{result.message}"
        self.status_label.configure(text=text)

    def _open_target_failed(self, target: ResultTarget, message: str) -> None:
        """显示后台打开文件失败信息。"""

        self.opening_targets.discard(target)
        self._sync_result_actions()
        self.status_label.configure(text="打开失败")
        messagebox.showerror("打开失败", f"无法打开文件：{target.file_path}\n{message}")

    def _clear_results(self) -> None:
        """清空表格或缩略图结果、行目标映射和悬浮操作按钮。"""

        self._cancel_tree_render()
        if self.local_filter_job is not None:
            self.after_cancel(self.local_filter_job)
            self.local_filter_job = None
        self.selected_target = None
        self.selected_value = ""
        self._sync_result_actions()
        self.preview_pane.clear()
        self.selection_label.configure(text="选择结果可查看详情")
        self.item_raw_values.clear()
        if not self._filter_rendering:
            self.result_source = None
            self.local_sheet_scope = None
            if self.local_filter_var.get():
                self.local_filter_var.set("")
                if self.local_filter_job is not None:
                    self.after_cancel(self.local_filter_job)
                    self.local_filter_job = None
            self.result_count_label.configure(text="检索结果  0")
            self.result_summary_label.configure(text="等待检索结果…" if self.search_cancel_event else "选择目录，输入关键词即可开始")
        self.action_item_ids.clear()
        self.result_group_controls.grid_remove()
        self._hide_result_grid_lines()
        self.file_thumbnail_results.clear()
        self.file_thumbnail_frame.clear()
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
        items = self.result_tree.get_children()
        if items:
            self.result_tree.delete(*items)

    def _render_results(self, results: list[ExcelWorkbookInfo]) -> None:
        """工作表渲染、结果计数和导出共用同一份可见行规则。"""
        self._clear_results()
        if not self._filter_rendering:
            self.result_sheet_enabled_only = self.only_enabled_sheets_var.get()
            self.result_disabled_marker = self.disabled_sheet_marker
        self._record_result_source("sheets", list(results), self._count_display_rows(results), f"{len(results)} 个工作簿 · 工作表结果")
        self._show_table_results()
        self._configure_result_columns(WORKSHEET_RESULT_COLUMNS)

        def rows() -> Iterator[None]:
            for result, sheet, status, description in self._worksheet_display_rows(results):
                item = self.result_tree.insert("", "end", values=(result.workbook_name, sheet, status, ""))
                valid_sheet = not description
                self.item_targets[item] = ResultTarget(result.path, sheet if valid_sheet else None)
                self.item_raw_values[item] = description or sheet
                if valid_sheet:
                    self.action_item_ids.add(item)
                yield
        self._render_tree_rows(rows())

    def _handle_result_click(self, event: object) -> str | None:
        """工作表操作列只执行引用查找，文件打开使用统一操作栏。"""

        column_key = self._get_column_key(self.result_tree.identify_column(event.x))
        if column_key != "action":
            return None

        item_id = self.result_tree.identify_row(event.y)
        if not item_id:
            return "break"
        if not self._action_button_is_available(item_id):
            return "break"

        self.result_tree.selection_set(item_id)
        self.result_tree.focus(item_id)
        self._find_references_for_item(item_id)
        return "break"

    def _handle_result_double_click(self, event: object) -> str:
        """处理双击事件，默认打开对应目标。"""

        if self.result_tree.identify_element(event.x, event.y) == "Treeitem.indicator":
            return "break"
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
            self._open_result_item(item_id)
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

        return (
            item_id in self.action_item_ids
            and self.result_tree.exists(item_id)
            and item_id not in self.reference_searching_items
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
            values[status_index] = get_sheet_enabled_status(target.sheet_name, self.result_disabled_marker)
            self.result_tree.item(item_id, values=values)

    def _count_display_rows(self, results: list[ExcelWorkbookInfo]) -> int:
        """统计工作表检索完成后实际展示的行数。"""

        return sum(1 for _ in self._worksheet_display_rows(results))

    def _worksheet_display_rows(self, results):
        query = self.local_filter_var.get().strip().casefold()
        for result in results:
            if result.error:
                if self.local_sheet_scope is None and (not query or query in f"{result.path} {result.error}".casefold()):
                    yield result, "无法读取", "读取失败", result.error
                continue
            names = self._get_display_sheet_names(result)
            if names:
                for name in names:
                    yield result, name, get_sheet_enabled_status(name, self.result_disabled_marker), ""
            elif not self.result_sheet_enabled_only and not query and self.local_sheet_scope is None:
                yield result, "无匹配工作表", "", "无匹配工作表"

    def _get_display_sheet_names(self, result: ExcelWorkbookInfo) -> list[str]:
        names = self._get_search_sheet_names(result)
        scope = self.local_sheet_scope
        if scope is not None:
            if result.path != scope[0]:
                return []
            names = [name for name in names if name == scope[1]]
        query = self.local_filter_var.get().strip().casefold()
        if query and query not in str(result.path).casefold():
            names = [name for name in names if query in name.casefold()]
        return names

    def _get_search_sheet_names(self, result: ExcelWorkbookInfo) -> list[str]:
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

        if not self.result_sheet_enabled_only:
            return sheet_names
        return [
            sheet_name
            for sheet_name in sheet_names
            if get_sheet_enabled_status(sheet_name, self.result_disabled_marker) == "启用"
        ]

    def open_selected_item(self) -> None:
        """Enter 与底部主按钮打开相同的目标。"""

        selected_items = self.result_tree.selection()
        if not selected_items:
            messagebox.showinfo("未选择", "请先在列表中选择一个工作簿或工作表。")
            return
        self._open_result_item(selected_items[0])

    def _load_cached_inputs(self) -> None:
        """启动时读取上次文件夹、关键词、历史地址和设置。"""

        cache_data = load_cache_data()
        self.pinned_folders = normalize_folder_history(cache_data.get("pinned_folders"))[:10]
        self.preview_enabled = cache_data.get("preview_enabled", True) is not False
        self.density_var.set("紧凑" if cache_data.get("result_density") == "紧凑" else "舒适")
        self.file_result_view_var.set("表格" if cache_data.get("file_result_view") == "表格" else "缩略图")
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
            self.keyword_entry.delete(0, "end")
            self.keyword_entry.insert(0, keyword)

        self._refresh_pin_button()
        self._refresh_filter_chips()
        ttk.Style(self).configure("Treeview", rowheight=28 if self.density_var.get() == "紧凑" else 36)
        self._sync_preview_visibility()

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
                "pinned_folders": self.pinned_folders,
                "preview_enabled": self.preview_enabled,
                "result_density": self.density_var.get(),
                "file_result_view": self.file_result_view_var.get(),
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
        """先取消任务再销毁窗口；关闭时不再刷新预览控件。"""

        if self.is_closing:
            return
        self.is_closing = True
        try:
            self._cancel_tree_render()
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
        finally:
            try:
                for job in self.tk.call("after", "info"):
                    self.tk.call("after", "cancel", job)
            finally:
                # 子组件的 destroy 负责停止工作线程和释放图片；不再调用 clear。
                self.destroy()


def get_sheet_enabled_status(sheet_name: str, disabled_marker: str) -> str:
    """根据工作表名称前缀判断启用状态。"""

    if not sheet_name:
        return ""
    marker = disabled_marker.strip()
    return "未启用" if marker and sheet_name.startswith(marker) else "启用"


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
