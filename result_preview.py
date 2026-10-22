"""异步、只读的结果预览。预览任务只读取当前选中文件，不参与检索。"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from queue import Empty, Full, Queue
import stat
from threading import Event, Thread
from tkinter import TclError, ttk
from typing import Callable

import customtkinter as ctk
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps
import ui_theme as theme


IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff", ".ico"}
XLSX_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
MAX_PREVIEW_TEXT = 100_000


class PreviewCancelled(Exception):
    """选中项已改变，不再需要本次预览。"""


@dataclass(frozen=True)
class PreviewRow:
    index: int
    values: tuple[object, ...]


@dataclass(frozen=True)
class CellPreviewData:
    first_column: int
    headers: tuple[str, ...]
    rows: tuple[PreviewRow, ...]
    field_name: str = ""


@dataclass(frozen=True)
class FilePreviewData:
    is_directory: bool
    size: int
    modified: str
    image: Image.Image | None = None
    dimensions: tuple[int, int] | None = None
    image_error: str = ""


def _check_cancelled(cancel_event: Event) -> None:
    if cancel_event.is_set():
        raise PreviewCancelled()


def _display_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ").removesuffix(" 00:00:00")
    return str(value)


def _preview_bounds(row_index: int, column_index: int) -> tuple[int, int, int, int]:
    if not 1 <= row_index <= 1_048_576 or not 1 <= column_index <= 16_384:
        raise ValueError("单元格位置超出 Excel 范围")
    return max(1, row_index - 2), min(1_048_576, row_index + 2), max(1, column_index - 2), min(16_384, column_index + 2)


def _read_cell_preview(
    path: Path, sheet_name: str, row_index: int, column_index: int, cancel_event: Event,
) -> CellPreviewData:
    """仅收集目标周围至多 5×5 个值；工作簿与行迭代器均在取消时关闭。"""

    first_row, last_row, first_column, last_column = _preview_bounds(row_index, column_index)
    _check_cancelled(cancel_event)
    nearby: list[PreviewRow] = []
    header_values: list[tuple[object, ...]] = []
    suffix = path.suffix.lower()

    def collect(index: int, values: tuple[object, ...]) -> None:
        if index <= 5:
            header_values.append(values)
        if first_row <= index <= last_row:
            nearby.append(PreviewRow(index, values))

    if suffix in XLSX_SUFFIXES:
        workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        row_iterator = None
        try:
            _check_cancelled(cancel_event)
            worksheet = workbook[sheet_name]
            # One sequential iterator avoids reopening the worksheet XML for the header.
            row_iterator = worksheet.iter_rows(
                min_row=1, max_row=last_row, min_col=first_column,
                max_col=last_column, values_only=True,
            )
            for index, values in enumerate(row_iterator, start=1):
                _check_cancelled(cancel_event)
                collect(index, tuple(values))
        finally:
            if row_iterator is not None:
                row_iterator.close()
            workbook.close()
    elif suffix == ".xls":
        import xlrd

        workbook = xlrd.open_workbook(str(path), on_demand=True)
        try:
            _check_cancelled(cancel_event)
            worksheet = workbook.sheet_by_name(sheet_name)
            indices = sorted(set(range(1, min(5, worksheet.nrows) + 1)) | set(range(first_row, min(last_row, worksheet.nrows) + 1)))
            for index in indices:
                _check_cancelled(cancel_event)
                values: list[object] = []
                for column in range(first_column - 1, last_column):
                    if column >= worksheet.ncols:
                        values.append(None)
                        continue
                    cell = worksheet.cell(index - 1, column)
                    value = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        value = xlrd.xldate_as_datetime(value, workbook.datemode)
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        value = bool(value)
                    elif cell.ctype == xlrd.XL_CELL_ERROR:
                        value = xlrd.error_text_from_code.get(value, "#ERROR")
                    values.append(value)
                collect(index, tuple(values))
        finally:
            workbook.release_resources()
    elif suffix == ".xlsb":
        from pyxlsb import open_workbook

        with open_workbook(str(path)) as workbook:
            with workbook.get_sheet(sheet_name) as worksheet:
                row_iterator = worksheet.rows(sparse=True)
                try:
                    for row in row_iterator:
                        _check_cancelled(cancel_event)
                        if not row:
                            continue
                        index = row[0].r + 1
                        if index > last_row:
                            break
                        if index > 5 and index < first_row:
                            continue
                        values_by_column = {cell.c + 1: cell.v for cell in row if first_column <= cell.c + 1 <= last_column}
                        collect(index, tuple(values_by_column.get(column) for column in range(first_column, last_column + 1)))
                finally:
                    row_iterator.close()
    else:
        raise ValueError(f"暂不支持 {suffix or '此格式'} 的单元格预览")

    _check_cancelled(cancel_event)
    width = last_column - first_column + 1
    header_names: list[str] = []
    for offset in range(width):
        # Config sheets commonly use one of their first five rows for field names.
        header_name = next((str(values[offset]).strip() for values in header_values if offset < len(values) and isinstance(values[offset], str) and str(values[offset]).strip()), "")
        header_names.append(header_name)
    field_name = header_names[column_index - first_column]
    return CellPreviewData(first_column, tuple(header_names), tuple(nearby), field_name)


def _read_file_preview(path: Path, cancel_event: Event) -> FilePreviewData:
    _check_cancelled(cancel_event)
    metadata = path.stat()
    is_directory = stat.S_ISDIR(metadata.st_mode)
    modified = datetime.fromtimestamp(metadata.st_mtime).strftime("%Y-%m-%d %H:%M")
    preview_image = None
    dimensions = None
    image_error = ""
    if not is_directory and path.suffix.lower() in IMAGE_SUFFIXES:
        try:
            with Image.open(path) as original:
                _check_cancelled(cancel_event)
                with ImageOps.exif_transpose(original) as oriented:
                    dimensions = oriented.size
                    oriented.thumbnail((580, 420), Image.Resampling.LANCZOS)
                    preview_image = oriented.convert("RGBA")
        except PreviewCancelled:
            raise
        except (OSError, ValueError, Image.DecompressionBombError) as exc:
            image_error = f"无法预览此图片：{exc}"
    _check_cancelled(cancel_event)
    return FilePreviewData(is_directory, metadata.st_size, modified, preview_image, dimensions, image_error)


def _format_file_size(size: int) -> str:
    amount = float(size)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if amount < 1024 or unit == "TB":
            return f"{int(amount)} B" if unit == "B" else f"{amount:.1f} {unit}"
        amount /= 1024
    return str(size)


def _set_label_image(label: ctk.CTkLabel, image: ctk.CTkImage | None, text: str) -> None:
    """先解除原生 Tk 图片引用，再替换 CTkImage，最后修改文字。

    CustomTkinter 6.0 的 image=None 不会清除原生 Label 的 image，且一次
    configure 调用会先处理 text。先显式清除原生引用，避免旧 PhotoImage 已
    释放时更新文字引发 ``image ... doesn't exist``。
    """

    label._label.configure(image="")
    label.configure(image=image)
    label.configure(text=text)


class ResultPreviewPane(ctk.CTkFrame):
    """当前选中结果的检查器。I/O、Excel 解析与图像解码均在单个后台线程执行。"""

    def __init__(
        self, parent: object, on_open: Callable[[], None],
        on_directory: Callable[[], None], on_copy_path: Callable[[], None],
    ) -> None:
        super().__init__(parent, width=350, fg_color=theme.SURFACE, corner_radius=9, border_width=1, border_color=theme.BORDER)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self._generation = 0
        self._closed = Event()
        self._cancel_event = Event()
        self._requests: Queue[tuple[int, Event, str, tuple[object, ...]] | None] = Queue(maxsize=1)
        self._responses: Queue[tuple[int, str, object]] = Queue(maxsize=2)
        self._debounce_job: str | None = None
        self._poll_job: str | None = None
        self._selection_key: tuple[object, ...] | None = None
        self._image: ctk.CTkImage | None = None
        self._current_row = 0
        self._keyword = ""

        header = ctk.CTkFrame(self, fg_color="transparent", height=48)
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(8, 3))
        header.grid_columnconfigure(0, weight=1)
        self.title_label = ctk.CTkLabel(header, text="结果预览", font=("Microsoft YaHei UI", 14, "bold"), text_color=theme.TEXT, anchor="w")
        self.title_label.grid(row=0, column=0, sticky="w", pady=8)
        ctk.CTkLabel(header, text="只读", font=("Microsoft YaHei UI", 11), text_color=theme.SUBTLE).grid(row=0, column=1)
        ctk.CTkFrame(self, height=1, fg_color=theme.DIVIDER, corner_radius=0).grid(row=1, column=0, sticky="ew", padx=1)

        self.body = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0, scrollbar_button_color=theme.BORDER, scrollbar_button_hover_color=theme.ACCENT_HOVER)
        self.body.grid(row=2, column=0, sticky="nsew", padx=(14, 6), pady=(12, 3))
        self.body.grid_columnconfigure(0, weight=1)
        self.empty_label = ctk.CTkLabel(self.body, text="选择一条结果\n在这里查看完整内容与相关信息", text_color=theme.MUTED, font=("Microsoft YaHei UI", 12), justify="center", height=160, wraplength=260)
        self.empty_label.grid(row=0, column=0, sticky="ew", pady=20)

        self.cell_frame = ctk.CTkFrame(self.body, fg_color="transparent")
        self.cell_frame.grid_columnconfigure(0, weight=1)
        self.workbook_label = self._metadata_label(self.cell_frame, 0)
        self.sheet_label = self._metadata_label(self.cell_frame, 1)
        self.field_label = self._metadata_label(self.cell_frame, 2)
        ctk.CTkLabel(self.cell_frame, text="完整内容", font=("Microsoft YaHei UI", 11), text_color=theme.MUTED, anchor="w").grid(row=3, column=0, sticky="ew", pady=(18, 7))
        self.content_box = ctk.CTkTextbox(self.cell_frame, height=140, fg_color=theme.INSET_BG, text_color=theme.TEXT, border_width=1, border_color=theme.BORDER, corner_radius=7, font=("Microsoft YaHei UI", 12), wrap="word", spacing1=3, spacing3=4)
        self.content_box.grid(row=4, column=0, sticky="ew")
        self.content_box.configure(state="disabled")
        self.content_box.tag_config("keyword", background=theme.HIGHLIGHT_BG, foreground=theme.HIGHLIGHT_TEXT)
        self.nearby_label = ctk.CTkLabel(self.cell_frame, text="相关数据", font=("Microsoft YaHei UI", 11), text_color=theme.MUTED, anchor="w")
        self.nearby_label.grid(row=5, column=0, sticky="ew", pady=(16, 7))
        style = ttk.Style(self)
        style.configure("Preview.Treeview", font=("Microsoft YaHei UI", 10), rowheight=28, background=theme.TABLE_BG, fieldbackground=theme.TABLE_BG, foreground=theme.TEXT, borderwidth=0)
        style.configure("Preview.Treeview.Heading", font=("Microsoft YaHei UI", 10), background=theme.HEADER_BG, foreground=theme.HEADING_TEXT, relief="flat")
        style.map("Preview.Treeview", background=[("selected", theme.SELECTION_BG)], foreground=[("selected", theme.SELECTION_TEXT)])
        self.nearby_table_frame = ctk.CTkFrame(self.cell_frame, width=200, height=190, fg_color="transparent", corner_radius=0)
        self.nearby_table_frame.grid(row=6, column=0, sticky="ew")
        self.nearby_table_frame.grid_propagate(False)
        self.nearby_table_frame.grid_columnconfigure(0, weight=1)
        self.nearby_table_frame.grid_rowconfigure(0, weight=1)
        self.nearby_tree = ttk.Treeview(self.nearby_table_frame, height=5, show="tree headings", style="Preview.Treeview", selectmode="browse")
        self.nearby_tree.grid(row=0, column=0, sticky="nsew")
        self.nearby_horizontal_scrollbar = ttk.Scrollbar(self.nearby_table_frame, orient="horizontal", command=self.nearby_tree.xview)
        self.nearby_horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        self.nearby_tree.configure(xscrollcommand=self.nearby_horizontal_scrollbar.set)
        self.nearby_tree.heading("#0", text="行")
        self.nearby_tree.column("#0", width=44, minwidth=44, stretch=False, anchor="center")
        self.cell_status_label = ctk.CTkLabel(self.cell_frame, text="", font=("Microsoft YaHei UI", 11), text_color=theme.SUBTLE, anchor="w", justify="left", wraplength=280)
        self.cell_status_label.grid(row=7, column=0, sticky="ew", pady=(7, 0))

        self.file_frame = ctk.CTkFrame(self.body, fg_color="transparent")
        self.file_frame.grid_columnconfigure(0, weight=1)
        self.file_image_frame = ctk.CTkFrame(self.file_frame, fg_color=theme.INSET_BG, corner_radius=7, border_width=1, border_color=theme.DIVIDER, height=210)
        self.file_image_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self.file_image_frame.grid_propagate(False)
        self.file_image_frame.grid_columnconfigure(0, weight=1)
        self.file_image_frame.grid_rowconfigure(0, weight=1)
        self.file_image_label = ctk.CTkLabel(self.file_image_frame, text="", text_color=theme.SUBTLE, font=("Microsoft YaHei UI", 15), wraplength=260, justify="center")
        self.file_image_label.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.file_name_label = ctk.CTkLabel(self.file_frame, text="", font=("Microsoft YaHei UI", 13, "bold"), text_color=theme.TEXT, anchor="w", justify="left", wraplength=280)
        self.file_name_label.grid(row=1, column=0, sticky="ew", pady=(0, 9))
        self.file_metadata_label = ctk.CTkLabel(self.file_frame, text="", font=("Microsoft YaHei UI", 11), text_color=theme.MUTED, anchor="w", justify="left", wraplength=280)
        self.file_metadata_label.grid(row=2, column=0, sticky="ew")

        self.path_frame = ctk.CTkFrame(self.body, fg_color="transparent")
        self.path_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.path_frame, text="文件位置", text_color=theme.MUTED, font=("Microsoft YaHei UI", 11), anchor="w").grid(row=0, column=0, sticky="w")
        self.copy_path_button = ctk.CTkButton(self.path_frame, text="复制", width=44, height=24, fg_color="transparent", hover_color=theme.HOVER_BG, text_color=theme.ACCENT, font=("Microsoft YaHei UI", 11), command=on_copy_path)
        self.copy_path_button.grid(row=0, column=1, sticky="e")
        self.path_box = ctk.CTkTextbox(self.path_frame, height=64, font=("Segoe UI", 11), fg_color=theme.INSET_BG, text_color=theme.MUTED, corner_radius=6, wrap="word")
        self.path_box.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 0))
        self.path_box.configure(state="disabled")

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=3, column=0, sticky="ew", padx=18, pady=(10, 4))
        footer.grid_columnconfigure(0, weight=1)
        self.open_button = ctk.CTkButton(footer, text="打开所选结果", height=35, fg_color=theme.ACCENT, hover_color=theme.ACCENT_HOVER, text_color=theme.ON_ACCENT, font=("Microsoft YaHei UI", 12, "bold"), command=on_open)
        self.open_button.grid(row=0, column=0, sticky="ew", padx=(0, 7))
        self.directory_button = ctk.CTkButton(footer, text="目录", width=64, height=35, fg_color=theme.SURFACE, border_width=1, border_color=theme.BORDER, hover_color=theme.HOVER_BG, text_color=theme.TEXT, font=("Microsoft YaHei UI", 11), command=on_directory)
        self.directory_button.grid(row=0, column=1)
        ctk.CTkLabel(self, text="单击预览 · 双击打开 · Ctrl+C 复制内容", font=("Microsoft YaHei UI", 10), text_color=theme.SUBTLE, anchor="w").grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 9))
        self.bind("<Configure>", self._on_resize, add="+")
        self.clear()
        self._worker_thread = Thread(target=self._worker, name="result-preview", daemon=True)
        self._worker_thread.start()
        self._poll_job = self.after(40, self._poll_results)

    @staticmethod
    def _metadata_label(parent: object, row: int) -> ctk.CTkLabel:
        label = ctk.CTkLabel(parent, text="", font=("Microsoft YaHei UI", 11), text_color=theme.MUTED, anchor="w", justify="left", wraplength=280)
        label.grid(row=row, column=0, sticky="ew", pady=3)
        return label

    def show_cell(self, path: Path, sheet_name: str, row_index: int, column_index: int, value: object, keyword: str = "") -> None:
        path = Path(path)
        key = ("cell", str(path), sheet_name, row_index, column_index, _display_text(value), keyword)
        if key == self._selection_key:
            return
        self._new_selection(key)
        self._current_row = row_index
        self._keyword = keyword
        self.title_label.configure(text="单元格预览")
        self.empty_label.grid_remove()
        self._set_file_image(None, "")
        self.file_frame.grid_remove()
        self.cell_frame.grid(row=0, column=0, sticky="ew")
        self.path_frame.grid(row=1, column=0, sticky="ew", pady=(18, 3))
        self.workbook_label.configure(text=f"工作簿    {path.name}")
        self.sheet_label.configure(text=f"工作表    {sheet_name}   /   {get_column_letter(column_index)}{row_index}")
        self.field_label.configure(text="字段        正在读取…")
        self._set_text(self.content_box, _display_text(value))
        self._highlight_keyword(keyword)
        self._set_text(self.path_box, str(path))
        self.nearby_tree.delete(*self.nearby_tree.get_children())
        self.nearby_label.configure(text="相关数据")
        self.cell_status_label.configure(text="正在读取附近数据…")
        self.open_button.configure(text="在 Excel 中定位", state="normal")
        self.directory_button.configure(state="normal")
        self.copy_path_button.configure(state="normal")
        self._schedule("cell", (path, sheet_name, row_index, column_index))

    def show_file(self, path: Path) -> None:
        path = Path(path)
        key = ("file", str(path))
        if key == self._selection_key:
            return
        self._new_selection(key)
        self.title_label.configure(text="文件预览")
        self.empty_label.grid_remove()
        self.cell_frame.grid_remove()
        self.file_frame.grid(row=0, column=0, sticky="ew")
        self.path_frame.grid(row=1, column=0, sticky="ew", pady=(18, 3))
        self._set_file_image(None, "正在准备预览…")
        self.file_name_label.configure(text=path.name or str(path))
        self.file_metadata_label.configure(text="正在读取文件信息…")
        self._set_text(self.path_box, str(path))
        self.open_button.configure(text="打开文件", state="normal")
        self.directory_button.configure(state="normal")
        self.copy_path_button.configure(state="normal")
        self._schedule("file", (path,))

    def _new_selection(self, key: tuple[object, ...] | None) -> None:
        self._generation += 1
        self._cancel_event.set()
        self._cancel_event = Event()
        self._selection_key = key
        if self._debounce_job is not None:
            self.after_cancel(self._debounce_job)
            self._debounce_job = None
        self._discard_queue(self._requests)
        self._discard_queue(self._responses)
        self.body._parent_canvas.yview_moveto(0)

    @staticmethod
    def _discard_queue(queue: Queue) -> None:
        while True:
            try:
                queue.get_nowait()
            except Empty:
                return

    def _schedule(self, kind: str, args: tuple[object, ...]) -> None:
        generation, cancel_event = self._generation, self._cancel_event

        def submit() -> None:
            self._debounce_job = None
            if not self._closed.is_set() and generation == self._generation:
                self._discard_queue(self._requests)
                self._requests.put_nowait((generation, cancel_event, kind, args))

        self._debounce_job = self.after(150, submit)

    def _worker(self) -> None:
        while not self._closed.is_set():
            try:
                request = self._requests.get(timeout=0.2)
            except Empty:
                continue
            if request is None:
                return
            generation, cancel_event, kind, args = request
            try:
                _check_cancelled(cancel_event)
                result = _read_cell_preview(*args, cancel_event) if kind == "cell" else _read_file_preview(*args, cancel_event)
                _check_cancelled(cancel_event)
            except PreviewCancelled:
                continue
            except Exception as exc:
                result = str(exc)
                kind = "error"
            if not self._closed.is_set() and not cancel_event.is_set():
                try:
                    self._responses.put_nowait((generation, kind, result))
                except Full:
                    pass

    def _poll_results(self) -> None:
        self._poll_job = None
        if self._closed.is_set():
            return
        while True:
            try:
                generation, kind, result = self._responses.get_nowait()
            except Empty:
                break
            if generation != self._generation:
                continue
            if kind == "cell":
                self._show_cell_data(result)
            elif kind == "file":
                self._show_file_data(result)
            elif self._selection_key and self._selection_key[0] == "cell":
                self.field_label.configure(text="字段        暂不可用")
                self.cell_status_label.configure(text=f"附近数据暂不可用：{result}\n仍可查看命中内容或在 Excel 中定位。")
            else:
                self._set_file_image(None, "预览暂不可用")
                self.file_metadata_label.configure(text=f"无法读取文件信息：{result}")
        self._poll_job = self.after(40, self._poll_results)

    def _show_cell_data(self, data: CellPreviewData) -> None:
        self.field_label.configure(text=f"字段        {data.field_name or '未找到字段名称'}")
        columns = tuple(f"column{index}" for index in range(len(data.headers)))
        self.nearby_tree.configure(columns=columns)
        for offset, column in enumerate(columns):
            letter = get_column_letter(data.first_column + offset)
            heading = f"{letter} · {data.headers[offset]}" if data.headers[offset] else letter
            self.nearby_tree.heading(column, text=heading)
            self.nearby_tree.column(column, width=86, minwidth=76, stretch=False, anchor="w")
        self.nearby_tree.xview_moveto(0)
        self.nearby_tree.delete(*self.nearby_tree.get_children())
        for row in data.rows:
            item = self.nearby_tree.insert("", "end", text=str(row.index), values=tuple(_display_text(value).replace("\n", " ")[:150] for value in row.values))
            if row.index == self._current_row:
                self.nearby_tree.selection_set(item)
        if data.rows:
            self.nearby_label.configure(text=f"相关数据    第 {data.rows[0].index}–{data.rows[-1].index} 行")
            self.cell_status_label.configure(text="附近数据为磁盘中已保存的值。")
        else:
            self.cell_status_label.configure(text="附近没有可预览的数据。")

    def _show_file_data(self, data: FilePreviewData) -> None:
        suffix = Path(self._selection_key[1]).suffix.upper().lstrip(".") if self._selection_key else ""
        kind = "文件夹" if data.is_directory else f"{suffix} 文件" if suffix else "文件"
        details = [f"类型    {kind}"]
        if not data.is_directory:
            details.append(f"大小    {_format_file_size(data.size)}")
        if data.dimensions:
            details.append(f"尺寸    {data.dimensions[0]} × {data.dimensions[1]} 像素")
        details.append(f"修改    {data.modified}")
        self.file_metadata_label.configure(text="\n\n".join(details))
        if data.image is not None:
            width = max(100, self.winfo_width() - 58)
            scale = min(width / data.image.width, 188 / data.image.height, 1)
            size = (max(1, round(data.image.width * scale)), max(1, round(data.image.height * scale)))
            image = ctk.CTkImage(light_image=data.image, dark_image=data.image, size=size)
            self._set_file_image(image, "")
        else:
            text = data.image_error or ("▱\n文件夹" if data.is_directory else f"{suffix or 'FILE'}\n使用默认程序打开此文件")
            self._set_file_image(None, text)
        self.open_button.configure(text="打开文件夹" if data.is_directory else "打开文件")

    def _set_file_image(self, image: ctk.CTkImage | None, text: str) -> None:
        _set_label_image(self.file_image_label, image, text)
        # Hold the previous image until the native Label has detached from it.
        self._image = image

    @staticmethod
    def _set_text(widget: ctk.CTkTextbox, text: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        if len(text) > MAX_PREVIEW_TEXT:
            text = text[:MAX_PREVIEW_TEXT] + "\n… 内容较长，后续请在 Excel 中查看。"
        widget.insert("1.0", text)
        widget.configure(state="disabled")
        widget.yview_moveto(0)

    def _highlight_keyword(self, keyword: str) -> None:
        self.content_box.tag_remove("keyword", "1.0", "end")
        if not keyword:
            return
        content = self.content_box.get("1.0", "end-1c")
        index = 0
        # Use Tcl's character indices so astral Unicode does not shift the highlight.
        for _ in range(200):
            found = self.content_box._textbox.search(keyword, f"1.0+{index}c", stopindex="end", nocase=True)
            if not found:
                break
            count = self.content_box._textbox.tk.call("string", "length", keyword)
            end = f"{found}+{count}c"
            self.content_box.tag_add("keyword", found, end)
            index = int(self.content_box._textbox.count("1.0", end, "chars")[0])
            if index >= len(content) * 2:
                break

    def _on_resize(self, event: object) -> None:
        if getattr(event, "widget", None) is not self:
            return
        width = max(100, self.winfo_width() - 55)
        for label in (self.workbook_label, self.sheet_label, self.field_label, self.file_name_label, self.file_metadata_label, self.cell_status_label):
            label.configure(wraplength=width)

    def clear(self) -> None:
        if self._closed.is_set():
            return
        self._new_selection(None)
        self._set_file_image(None, "")
        self.nearby_tree.delete(*self.nearby_tree.get_children())
        self._set_text(self.content_box, "")
        self._set_text(self.path_box, "")
        self.title_label.configure(text="结果预览")
        self.cell_frame.grid_remove()
        self.file_frame.grid_remove()
        self.path_frame.grid_remove()
        self.empty_label.grid(row=0, column=0, sticky="ew", pady=20)
        self.open_button.configure(text="打开所选结果", state="disabled")
        self.directory_button.configure(state="disabled")
        self.copy_path_button.configure(state="disabled")

    def destroy(self) -> None:
        self._closed.set()
        self._cancel_event.set()
        for job in (self._debounce_job, self._poll_job):
            if job is not None:
                try:
                    self.after_cancel(job)
                except TclError:
                    pass
        self._discard_queue(self._requests)
        try:
            self._requests.put_nowait(None)
        except Full:
            pass
        self._discard_queue(self._responses)
        try:
            self._set_file_image(None, "")
        except TclError:
            # An ancestor may already have destroyed the underlying Tk label.
            self._image = None
        super().destroy()
