"""文件分组的虚拟滚动视图：只绘制可见卡片，后台生成缩略图。"""

from bisect import bisect_left, bisect_right
from collections import OrderedDict
from pathlib import Path
from queue import Empty, Full, Queue
from threading import Event, Thread
from tkinter import Canvas, font as tkfont, ttk
from typing import Callable

import customtkinter as ctk
from PIL import Image, ImageTk, UnidentifiedImageError

from file_search import FileSearchResult
import ui_theme as theme


IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
CARD_HEIGHT = 192
HEADER_HEIGHT = 36
CACHE_LIMIT = 256
THUMBNAIL_SIZE = (256, 118)


class FileResultView(ctk.CTkFrame):
    def __init__(
        self, parent: object, open_result: Callable[[FileSearchResult], None],
        get_style: Callable[[FileSearchResult], tuple[str, str, str]], columns: int = 4,
        *, on_select: Callable[[FileSearchResult], None] | None = None,
        on_context_menu: Callable[[FileSearchResult, object], None] | None = None,
        on_copy: Callable[[FileSearchResult], None] | None = None,
    ) -> None:
        super().__init__(parent, fg_color=theme.TABLE_BG, corner_radius=0)
        self.open_result = open_result
        self.get_style = get_style
        self.columns = max(1, columns)
        self.on_select = on_select
        self.on_context_menu = on_context_menu
        self.on_copy = on_copy
        self._layout_columns = self.columns
        self._groups: list[tuple[Path, list[int]]] = []
        self._group_counts: dict[str, int] = {}
        self._name_font = tkfont.Font(self, family="Microsoft YaHei UI", size=9)
        self._detail_font = tkfont.Font(self, family="Segoe UI", size=8)
        self._group_font = tkfont.Font(self, family="Microsoft YaHei UI", size=9)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.canvas = Canvas(self, background=theme.TABLE_BG, highlightthickness=0, yscrollincrement=32, takefocus=True)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=self._on_scroll)
        self.canvas.bind("<Configure>", self._on_resize)
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Button-4>", lambda _event: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Button-5>", lambda _event: self.canvas.yview_scroll(3, "units"))
        self.canvas.bind("<ButtonRelease-1>", self._on_click)
        self.canvas.bind("<Double-1>", self._on_double_click)
        self.canvas.bind("<Button-3>", self._on_context_menu)
        self.canvas.bind("<KeyPress>", self._on_key)
        self.canvas.bind("<FocusIn>", lambda _event: self._schedule_draw())
        self.canvas.bind("<FocusOut>", lambda _event: self._schedule_draw())
        self.results: list[FileSearchResult] = []
        self.rows: list[tuple[int, str | None, list[int]]] = []
        self.row_positions: list[int] = []
        self.item_positions: list[int] = []
        self.selected_index = 0
        self.cards: list[FileSearchResult] = []
        self.generation = 0
        self._closed = Event()
        self._draw_job: str | None = None
        self._poll_job: str | None = self.after(40, self._poll_images)
        self._tasks: Queue[tuple[int, Path, bool]] = Queue(maxsize=32)
        self._loaded: Queue[tuple] = Queue(maxsize=32)
        self._pending: set[Path] = set()
        self._cache: OrderedDict[Path, Image.Image | None] = OrderedDict()
        self._metadata: dict[Path, tuple[int | None, tuple[int, int] | None]] = {}
        self._photos: dict[Path, ImageTk.PhotoImage] = {}
        self._photo_cache: OrderedDict[tuple[Path, int, int], ImageTk.PhotoImage] = OrderedDict()
        for _ in range(2):
            Thread(target=self._image_worker, daemon=True).start()

    def set_results(self, results: list[FileSearchResult]) -> None:
        self.clear()
        self.results = list(results)
        grouped: dict[Path, list[int]] = {}
        for index, result in enumerate(self.results):
            grouped.setdefault(result.path.parent, []).append(index)
        self._groups = list(grouped.items())
        self._group_counts = {str(folder): len(indices) for folder, indices in self._groups}
        self._reflow_rows(self._columns_for_width(self.canvas.winfo_width()))
        self.canvas.yview_moveto(0)
        self._draw_visible()

    def _columns_for_width(self, width: int) -> int:
        return self.columns if width <= 1 else max(1, min(self.columns, width // 170))

    def _reflow_rows(self, columns: int) -> None:
        self._layout_columns = columns
        self.rows.clear()
        self.item_positions = [0] * len(self.results)
        y = 4
        for folder, indices in self._groups:
            self.rows.append((y, str(folder), []))
            y += HEADER_HEIGHT
            for start in range(0, len(indices), columns):
                row_indices = indices[start:start + columns]
                self.rows.append((y, None, row_indices))
                for index in row_indices:
                    self.item_positions[index] = y
                y += CARD_HEIGHT
        self.row_positions = [row[0] for row in self.rows]
        self.canvas.configure(scrollregion=(0, 0, 1, y + 8))

    def _on_resize(self, event: object) -> None:
        columns = self._columns_for_width(event.width)
        if columns != self._layout_columns:
            top = self.canvas.canvasy(0)
            self._reflow_rows(columns)
            if self.results:
                total = self.row_positions[-1] + CARD_HEIGHT + 8
                y = self.item_positions[self.selected_index]
                height = max(240, self.canvas.winfo_height())
                if y < top:
                    top = y
                elif y + CARD_HEIGHT > top + height:
                    top = y + CARD_HEIGHT - height
                self.canvas.yview_moveto(top / total)
        self._schedule_draw()

    def get_selected_result(self) -> FileSearchResult | None:
        if 0 <= self.selected_index < len(self.results):
            return self.results[self.selected_index]
        return None

    def clear(self) -> None:
        self.generation += 1
        if self._draw_job is not None:
            self.after_cancel(self._draw_job)
            self._draw_job = None
        self.results.clear()
        self._groups.clear()
        self._group_counts.clear()
        self.rows.clear()
        self.row_positions.clear()
        self.item_positions.clear()
        self.selected_index = 0
        self.cards.clear()
        self._pending.clear()
        self._cache.clear()
        self._metadata.clear()
        self._photos.clear()
        self._photo_cache.clear()
        for queue in (self._tasks, self._loaded):
            while True:
                try:
                    queue.get_nowait()
                except Empty:
                    break
        self.canvas.delete("all")
        self.canvas.configure(scrollregion=(0, 0, 1, 1))

    def _on_scroll(self, first: str, last: str) -> None:
        self.scrollbar.set(first, last)
        self._schedule_draw()

    def _on_wheel(self, event: object) -> str:
        delta = int(event.delta)
        self.canvas.yview_scroll(-max(1, abs(delta) // 120) * (1 if delta > 0 else -1) * 3, "units")
        return "break"

    def _schedule_draw(self) -> None:
        if not self._closed.is_set() and self._draw_job is None:
            self._draw_job = self.after_idle(self._draw_visible)

    def _draw_visible(self) -> None:
        self._draw_job = None
        if self._closed.is_set():
            return
        width = max(1, self.canvas.winfo_width())
        if width <= 1:
            width = 960
        height = max(240, self.canvas.winfo_height())
        top = self.canvas.canvasy(0)
        start = max(0, bisect_right(self.row_positions, top - CARD_HEIGHT) - 1)
        end = bisect_left(self.row_positions, top + height + CARD_HEIGHT)
        self.canvas.delete("content")
        self.cards.clear()
        self._photos.clear()
        if not self.results:
            self.canvas.create_text(width / 2, max(70, height / 3), text="没有找到匹配的文件或文件夹",
                                    fill=theme.MUTED, font=self._name_font, tags="content")
            return
        cell_width = width / self._layout_columns
        for y, title, indices in self.rows[start:end]:
            if title is not None:
                self.canvas.create_text(12, y + 17, text="▱", anchor="w", fill=theme.ACCENT,
                                        font=("Segoe UI", 13), tags="content")
                self.canvas.create_text(34, y + 17, text=self._fit_text(title, width - 140, self._group_font),
                                        anchor="w", fill=theme.MUTED, font=self._group_font, tags="content")
                self.canvas.create_text(width - 14, y + 17, text=f"{self._group_counts[title]} 项", anchor="e",
                                        fill=theme.SUBTLE, font=self._detail_font, tags="content")
                continue
            for column, index in enumerate(indices):
                self._draw_card(index, column * cell_width + 8, y + 4, cell_width - 16)

    def _draw_card(self, index: int, x: float, y: int, width: float) -> None:
        result = self.results[index]
        self.cards.append(result)
        symbol, color, item_type = self.get_style(result)
        tags = ("content", f"card:{index}")
        selected = index == self.selected_index
        self._rounded_rectangle(x, y, x + width, y + 180, 7, outline=theme.ACCENT if selected else theme.BORDER,
                                width=2 if selected else 1, fill=theme.SELECTION_BG if selected else theme.SURFACE, tags=tags)
        self._rounded_rectangle(x + 9, y + 9, x + width - 9, y + 127, 5,
                                fill=theme.INSET_BG, outline="", tags=tags)
        image = self._get_thumbnail(result.path, max(1, int(width - 20)), THUMBNAIL_SIZE[1], result.is_directory)
        if image is not None:
            self.canvas.create_image(x + width / 2, y + 68, image=image, tags=tags)
        else:
            self.canvas.create_text(x + width / 2, y + 65, text=symbol, fill=color,
                                    font=("Segoe UI Emoji", 35), tags=tags)
        if selected:
            self._rounded_rectangle(x + width - 30, y + 15, x + width - 14, y + 31, 3,
                                    fill=theme.ACCENT, outline="", tags=tags)
            self.canvas.create_text(x + width - 22, y + 23, text="✓", fill=theme.ON_ACCENT,
                                    font=("Segoe UI", 9), tags=tags)
        self.canvas.create_text(x + 10, y + 145, text=self._fit_text(result.file_name, width - 20, self._name_font),
                                anchor="w", fill=theme.SELECTION_TEXT if selected else theme.TEXT, font=self._name_font, tags=tags)
        size, dimensions = self._metadata.get(result.path, (None, None))
        detail = f"{dimensions[0]} × {dimensions[1]}" if dimensions else item_type
        size_text = self._format_size(size) if size is not None and not result.is_directory else ""
        detail_width = width - 20 - (55 if size_text else 0)
        self.canvas.create_text(x + 10, y + 166, text=self._fit_text(detail, detail_width, self._detail_font),
                                anchor="w", fill=theme.SUBTLE, font=self._detail_font, tags=tags)
        self.canvas.create_text(x + width - 10, y + 166, text=size_text, anchor="e",
                                fill=theme.SUBTLE, font=self._detail_font, tags=tags)

    def _rounded_rectangle(self, left: float, top: float, right: float, bottom: float,
                           radius: int, **kwargs: object) -> int:
        return self.canvas.create_polygon(
            left + radius, top, right - radius, top, right, top, right, top + radius,
            right, bottom - radius, right, bottom, right - radius, bottom,
            left + radius, bottom, left, bottom, left, bottom - radius,
            left, top + radius, left, top, smooth=True, **kwargs,
        )

    @staticmethod
    def _format_size(size: int) -> str:
        if size < 1024:
            return f"{size} B"
        if size < 1024 * 1024:
            return f"{size / 1024:.0f} KB"
        return f"{size / (1024 * 1024):.1f} MB"

    @staticmethod
    def _fit_text(value: str, width: float, font: tkfont.Font) -> str:
        if width <= 0:
            return ""
        if font.measure(value) <= width:
            return value
        low, high = 0, len(value)
        while low < high:
            middle = (low + high + 1) // 2
            if font.measure(value[:middle] + "…") <= width:
                low = middle
            else:
                high = middle - 1
        return value[:low] + "…"

    @staticmethod
    def _shorten(value: str, length: int) -> str:
        return value if len(value) <= length else value[:max(1, length - 1)] + "…"

    def _index_at(self, event: object) -> int | None:
        # Coordinates work for both actual pointer events and callers that invoke
        # selection from another widget; don't depend on the transient current tag.
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        for item in reversed(self.canvas.find_overlapping(x, y, x, y)):
            for tag in self.canvas.gettags(item):
                if tag.startswith("card:"):
                    index = int(tag.split(":", 1)[1])
                    if 0 <= index < len(self.results):
                        return index
        return None

    def _select_index(self, index: int) -> None:
        self.selected_index = index
        self.canvas.focus_set()
        self._schedule_draw()
        if self.on_select is not None:
            self.on_select(self.results[index])

    def _on_click(self, event: object) -> None:
        index = self._index_at(event)
        if index is not None:
            self._select_index(index)

    def _on_double_click(self, event: object) -> str:
        index = self._index_at(event)
        if index is not None:
            self._select_index(index)
            self.open_result(self.results[index])
        return "break"

    def _on_context_menu(self, event: object) -> str:
        index = self._index_at(event)
        if index is not None:
            self._select_index(index)
            if self.on_context_menu is not None:
                self.on_context_menu(self.results[index], event)
        return "break"

    def _on_key(self, event: object) -> str | None:
        if not self.results:
            return None
        if event.keysym.lower() == "c" and getattr(event, "state", 0) & 0x4:
            if self.on_copy is not None:
                self.on_copy(self.results[self.selected_index])
            return "break"
        if event.keysym in ("Return", "space"):
            self.open_result(self.results[self.selected_index])
            return "break"
        offsets = {"Left": -1, "Right": 1, "Up": -self._layout_columns, "Down": self._layout_columns,
                   "Home": -len(self.results), "End": len(self.results)}
        if event.keysym not in offsets:
            return None
        self._select_index(max(0, min(len(self.results) - 1, self.selected_index + offsets[event.keysym])))
        y = self.item_positions[self.selected_index]
        top = self.canvas.canvasy(0)
        height = max(240, self.canvas.winfo_height())
        total_height = self.row_positions[-1] + CARD_HEIGHT + 8
        if y < top:
            self.canvas.yview_moveto(y / total_height)
        elif y + CARD_HEIGHT > top + height:
            self.canvas.yview_moveto((y + CARD_HEIGHT - height) / total_height)
        self._schedule_draw()
        return "break"

    def _get_thumbnail(self, path: Path, width: int = THUMBNAIL_SIZE[0], height: int = THUMBNAIL_SIZE[1],
                       is_directory: bool = False) -> ImageTk.PhotoImage | None:
        if path not in self._cache:
            if path not in self._pending:
                try:
                    self._tasks.put_nowait((self.generation, path, is_directory))
                    self._pending.add(path)
                except Full:
                    pass
            return None
        self._cache.move_to_end(path)
        image = self._cache[path]
        if image is None:
            return None
        key = (path, width, height)
        if key not in self._photo_cache:
            resized = image.copy()
            resized.thumbnail((width, height), Image.Resampling.LANCZOS)
            self._photo_cache[key] = ImageTk.PhotoImage(resized, master=self.canvas)
        self._photo_cache.move_to_end(key)
        photo = self._photo_cache[key]
        self._photos[path] = photo
        while len(self._photo_cache) > CACHE_LIMIT:
            self._photo_cache.popitem(last=False)
        return photo

    def _image_worker(self) -> None:
        while not self._closed.is_set():
            try:
                generation, path, is_directory = self._tasks.get(timeout=0.2)
            except Empty:
                continue
            if generation != self.generation:
                continue
            image = None
            size = None
            dimensions = None
            try:
                if not is_directory:
                    size = path.stat().st_size
                if not is_directory and path.suffix.lower() in IMAGE_SUFFIXES:
                    with Image.open(path) as source:
                        dimensions = source.size
                        source.thumbnail(THUMBNAIL_SIZE, Image.Resampling.LANCZOS)
                        image = source.convert("RGBA")
            except (OSError, UnidentifiedImageError, ValueError, Image.DecompressionBombError):
                pass
            if generation == self.generation and not self._closed.is_set():
                while not self._closed.is_set() and generation == self.generation:
                    try:
                        self._loaded.put((generation, path, image, (size, dimensions)), timeout=0.1)
                        break
                    except Full:
                        continue

    def _poll_images(self) -> None:
        self._poll_job = None
        if self._closed.is_set():
            return
        updated = False
        for _ in range(16):
            try:
                loaded = self._loaded.get_nowait()
            except Empty:
                break
            generation, path, image = loaded[:3]
            if generation != self.generation:
                continue
            self._pending.discard(path)
            self._cache[path] = image
            self._metadata[path] = loaded[3] if len(loaded) > 3 else (None, None)
            while len(self._cache) > CACHE_LIMIT:
                old_path, _ = self._cache.popitem(last=False)
                self._metadata.pop(old_path, None)
                for key in [key for key in self._photo_cache if key[0] == old_path]:
                    self._photo_cache.pop(key, None)
            updated = True
        if updated or (self.results and len(self._pending) < 32):
            # 重试因有界队列满而暂未提交的可见图片。
            if updated or any(result.path not in self._cache
                              and result.path not in self._pending for result in self.cards):
                self._schedule_draw()
        self._poll_job = self.after(40, self._poll_images)

    def destroy(self) -> None:
        self._closed.set()
        for job in (self._draw_job, self._poll_job):
            if job is not None:
                self.after_cancel(job)
        super().destroy()
