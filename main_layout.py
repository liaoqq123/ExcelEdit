"""主窗口布局；业务处理和检索状态由 ExcelEditApp 管理。"""

import customtkinter as ctk
import ui_theme as ui
from tkinter import ttk

from file_result_view import FileResultView
from result_preview import ResultPreviewPane

MODES = {"工作表": "工作表检索", "单元格": "单元格检索", "文件": "文件检索", "JSON 检查": "JSON格式检查"}
BLUE = ui.ACCENT
INK = ui.TEXT
MUTED = ui.MUTED


def build_main_layout(app):
    app.configure(fg_color=ui.APP_BG)
    app.grid_columnconfigure(0, weight=1)
    app.grid_rowconfigure(3, weight=1)
    app.background_image_label = ctk.CTkLabel(app, text="", fg_color="transparent")
    app.background_image_label.place(x=0, y=0, relwidth=1, relheight=1)
    app.background_image_label.lower()
    app.bind("<Configure>", app._on_app_configure, add="+")

    def quiet(parent, text, command, width=84):
        return ctk.CTkButton(parent, text=text, command=command, width=width, height=30,
                             fg_color="transparent", hover_color=ui.HOVER_BG, text_color=INK,
                             border_width=1, border_color=ui.BORDER, corner_radius=6)

    heading = ctk.CTkFrame(app, fg_color=ui.TOOLBAR_BG, corner_radius=0)
    heading.grid(row=0, column=0, sticky="ew")
    heading.grid_columnconfigure(1, weight=1)
    ctk.CTkLabel(heading, text="查找配置与文件", font=("Microsoft YaHei UI", 23, "bold"), text_color=ui.TEXT).grid(row=0, column=0, padx=(24, 16), pady=(20, 14))
    ctk.CTkLabel(heading, text="目录历史 · 快速定位 · 只读预览", text_color=MUTED, anchor="w", font=("Microsoft YaHei UI", 12)).grid(row=0, column=1, sticky="w")
    app.settings_button = quiet(heading, "设置", app.open_settings, 68)
    app.settings_button.grid(row=0, column=2, padx=(8, 8))
    app.help_button = quiet(heading, "使用说明", app.open_help_url)
    app.help_button.grid(row=0, column=3, padx=(0, 24))
    app._register_translucent_strip("top", heading)

    directory = ctk.CTkFrame(app, fg_color=ui.TOOLBAR_BG, corner_radius=0)
    directory.grid(row=1, column=0, sticky="ew")
    directory.grid_columnconfigure(1, weight=1)
    ctk.CTkLabel(directory, text="工作目录", text_color=MUTED).grid(row=0, column=0, padx=(24, 14), pady=(0, 12))
    app.folder_entry = ctk.CTkComboBox(directory, variable=app.folder_path_var, values=[], command=app._select_folder_from_history,
        height=36, fg_color=ui.INPUT_BG, border_color=ui.BORDER, button_color=ui.HEADER_BG, button_hover_color=ui.HOVER_BG,
        dropdown_fg_color=ui.SURFACE, dropdown_hover_color=ui.HOVER_BG, dropdown_text_color=INK, text_color=INK)
    app.folder_entry.grid(row=0, column=1, sticky="ew", pady=(0, 12))
    app.folder_entry.bind("<FocusOut>", lambda _event: app._save_cache_from_inputs())
    app.folder_entry._open_dropdown_menu = app._open_folder_dropdown_menu
    directory.bind("<Configure>", lambda _event: app._sync_folder_dropdown_width(), add="+")
    app.after_idle(app._sync_folder_dropdown_width)
    quiet(directory, "选择目录", app.choose_folder, 98).grid(row=0, column=2, padx=(12, 8), pady=(0, 12))
    app.pin_folder_button = quiet(directory, "☆ 固定目录", app._toggle_pinned_folder, 104)
    app.pin_folder_button.grid(row=0, column=3, padx=(0, 24), pady=(0, 12))
    app._register_translucent_strip("search_input", directory)

    controls = ctk.CTkFrame(app, fg_color=ui.TOOLBAR_BG, corner_radius=0)
    controls.grid(row=2, column=0, sticky="ew")
    controls.grid_columnconfigure(1, weight=1)
    app.mode_display_var = ctk.StringVar(value="工作表")
    app.search_mode_menu = ctk.CTkSegmentedButton(controls, values=list(MODES), variable=app.mode_display_var, height=38,
        fg_color=ui.INSET_BG, selected_color=ui.SELECTION_BG, selected_hover_color=ui.HOVER_BG, unselected_color=ui.INSET_BG,
        unselected_hover_color=ui.HOVER_BG, text_color=INK, command=lambda value: app._on_search_mode_changed(MODES[value]))
    app.search_mode_menu.grid(row=0, column=0, padx=(24, 14), pady=(0, 12))
    app.keyword_entry = ctk.CTkEntry(controls, placeholder_text="输入工作簿或工作表关键词 · Enter 检索", height=40,
        border_color=ui.ACCENT, fg_color=ui.INPUT_BG, text_color=ui.TEXT, placeholder_text_color=ui.MUTED, corner_radius=7)
    app.keyword_entry.grid(row=0, column=1, sticky="ew", pady=(0, 12))
    app.keyword_entry.bind("<Return>", lambda _event: app.run_selected_search())
    app.keyword_entry.bind("<FocusOut>", lambda _event: app._save_cache_from_inputs())
    app.cancel_search_button = quiet(controls, "取消", app.cancel_current_search, 74)
    app.cancel_search_button.grid(row=0, column=2, padx=(10, 0), pady=(0, 12))
    app.cancel_search_button.grid_remove()
    app.run_search_button = ctk.CTkButton(controls, text="检索", command=app.run_selected_search, width=104, height=40,
                                       fg_color=BLUE, hover_color=ui.ACCENT_HOVER, corner_radius=7)
    app.run_search_button.grid(row=0, column=3, padx=(12, 24), pady=(0, 12))
    app.filter_chips_frame = ctk.CTkFrame(controls, fg_color="transparent", height=32)
    app.filter_chips_frame.grid(row=1, column=0, columnspan=4, padx=24, pady=(0, 14), sticky="ew")
    app._register_translucent_strip("search_action", controls)

    body = ctk.CTkFrame(app, fg_color="transparent")
    body.grid(row=3, column=0, padx=20, pady=(16, 0), sticky="nsew")
    body.grid_rowconfigure(0, weight=1)
    body.grid_columnconfigure(0, weight=1)
    app.results_body = body
    card = ctk.CTkFrame(body, fg_color=ui.SURFACE, border_width=1, border_color=ui.BORDER, corner_radius=9)
    card.grid(row=0, column=0, sticky="nsew")
    card.grid_columnconfigure(0, weight=1)
    card.grid_rowconfigure(1, weight=1)
    bar = ctk.CTkFrame(card, fg_color="transparent")
    bar.grid(row=0, column=0, padx=16, pady=(12, 8), sticky="ew")
    bar.grid_columnconfigure(1, weight=1)
    app.result_count_label = ctk.CTkLabel(bar, text="检索结果  0", text_color=INK, font=("Microsoft YaHei UI", 15, "bold"))
    app.result_count_label.grid(row=0, column=0, sticky="w")
    app.local_filter_entry = ctk.CTkEntry(bar, placeholder_text="在结果中查找 · Ctrl+F", width=186, height=30,
                                        border_color=ui.BORDER, fg_color=ui.INPUT_BG, text_color=INK, placeholder_text_color=ui.MUTED)
    app.local_filter_entry.grid(row=0, column=1, padx=(12, 8), sticky="e")
    def sync_local_input(*_args):
        value = app.local_filter_var.get()
        if app.local_filter_entry.get() != value:
            app.local_filter_entry.delete(0, "end")
            if value:
                app.local_filter_entry.insert(0, value)
        app._schedule_local_filter()
    app.local_filter_var.trace_add("write", sync_local_input)
    app.local_filter_entry.bind("<KeyRelease>", lambda _event: app.local_filter_var.set(app.local_filter_entry.get()))
    app.local_filter_entry.bind("<<Paste>>", lambda _event: app.after_idle(lambda: app.local_filter_var.set(app.local_filter_entry.get())), add="+")
    app.local_filter_entry.bind("<<Cut>>", lambda _event: app.after_idle(lambda: app.local_filter_var.set(app.local_filter_entry.get())), add="+")
    app.local_filter_entry.bind("<Escape>", lambda _event: app._reset_local_filter())
    app.density_menu = ctk.CTkOptionMenu(bar, values=["舒适", "紧凑"], variable=app.density_var, command=app._set_density,
                                      width=76, height=30, fg_color=ui.APP_BG, button_color=ui.HEADER_BG, text_color=INK)
    app.density_menu.grid(row=0, column=2, padx=(0, 8))
    quiet(bar, "导出", app._export_results, 56).grid(row=0, column=3, padx=(0, 8))
    app.preview_toggle_button = quiet(bar, "预览 ▸", app._toggle_preview, 70)
    app.preview_toggle_button.grid(row=0, column=4)
    app.result_summary_label = ctk.CTkLabel(bar, text="选择目录，输入关键词即可开始", anchor="w", text_color=MUTED, font=("Microsoft YaHei UI", 11))
    app.result_summary_label.grid(row=1, column=0, columnspan=5, sticky="w", pady=(3, 0))

    area = ctk.CTkFrame(card, fg_color=ui.TABLE_BG, corner_radius=0)
    app.result_list_frame = area
    area.grid(row=1, column=0, padx=1, sticky="nsew")
    area.grid_columnconfigure(0, weight=1)
    area.grid_rowconfigure(2, weight=1)
    app.result_group_controls = ctk.CTkFrame(area, fg_color=ui.TOOLBAR_BG, corner_radius=0, height=32)
    app.result_group_controls.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 4))
    app.result_group_controls.grid_columnconfigure(0, weight=1)
    app.collapse_all_results_button = quiet(app.result_group_controls, "全部收起", app.collapse_all_results, 82)
    app.collapse_all_results_button.grid(row=0, column=1, padx=(0, 8))
    app.expand_all_results_button = quiet(app.result_group_controls, "全部展开", app.expand_all_results, 82)
    app.expand_all_results_button.grid(row=0, column=2)
    app.result_group_controls.grid_remove()
    app.file_result_view_controls = ctk.CTkFrame(area, fg_color="transparent")
    app.file_result_view_controls.grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 6))
    app.file_result_view_controls.grid_columnconfigure(0, weight=1)
    app.file_result_count_label = ctk.CTkLabel(app.file_result_view_controls, text="", text_color=MUTED, anchor="w")
    app.file_result_count_label.grid(row=0, column=0, sticky="w")
    app.file_result_view_switch = ctk.CTkSegmentedButton(app.file_result_view_controls, values=["缩略图", "表格"],
        variable=app.file_result_view_var, command=app._on_file_result_view_changed, width=150,
        fg_color=ui.INSET_BG, selected_color=ui.SELECTION_BG, selected_hover_color=ui.HOVER_BG, unselected_color=ui.INSET_BG, text_color=INK)
    app.file_result_view_switch.grid(row=0, column=1)
    app.file_result_view_controls.grid_remove()
    app._configure_tree_style()
    app.result_tree = ttk.Treeview(area, columns=(), show="headings", selectmode="browse")
    app.result_tree.grid(row=2, column=0, sticky="nsew")
    app.result_tree.bind("<ButtonRelease-1>", app._handle_result_click)
    app.result_tree.bind("<Double-1>", app._handle_result_double_click)
    app.result_tree.bind("<<TreeviewSelect>>", app._on_result_selected)
    app.result_tree.bind("<Return>", lambda _event: app.open_selected_item())
    app.result_tree.bind("<Control-c>", lambda _event: app._copy_selected())
    app.result_tree.bind("<Button-3>", app._show_result_menu)
    app.result_tree.bind("<Configure>", lambda _event: app._schedule_action_button_layout())
    app.result_tree.bind("<Configure>", lambda _event: app._schedule_result_grid_lines(), add="+")
    app.result_tree.tag_configure("reference", foreground=ui.ACCENT)
    app.result_tree.tag_configure("cell-workbook", background=ui.GROUP_BG, foreground=ui.TEXT, font=("Microsoft YaHei UI", 10, "bold"))
    app.result_tree.tag_configure("cell-sheet", background=ui.TOOLBAR_BG, foreground=ui.MUTED)
    app.result_tree_vertical_scroll = ttk.Scrollbar(area, orient="vertical", command=app._scroll_result_tree_y)
    app.result_tree_vertical_scroll.grid(row=2, column=1, sticky="ns")
    app.result_tree_horizontal_scroll = ttk.Scrollbar(area, orient="horizontal", command=app._scroll_result_tree_x)
    app.result_tree_horizontal_scroll.grid(row=3, column=0, sticky="ew")
    app.result_tree.configure(yscrollcommand=lambda *args: app._update_scrollbar(app.result_tree_vertical_scroll, *args),
                              xscrollcommand=lambda *args: app._update_scrollbar(app.result_tree_horizontal_scroll, *args))
    app.file_thumbnail_frame = FileResultView(area, open_result=app._open_file_result, get_style=app._get_file_thumbnail_style, columns=4,
        on_select=app._on_file_selected, on_context_menu=app._show_file_menu, on_copy=lambda result: (app._on_file_selected(result), app._copy_selected("path")))
    app.file_thumbnail_frame.grid(row=2, column=0, columnspan=2, padx=4, pady=(0, 4), sticky="nsew")
    app.file_thumbnail_cards = app.file_thumbnail_frame.cards
    app.file_thumbnail_frame.grid_remove()

    selected = ctk.CTkFrame(card, fg_color=ui.INSET_BG, corner_radius=0, height=40)
    selected.grid(row=2, column=0, sticky="ew", padx=1, pady=(0, 1))
    selected.grid_columnconfigure(0, weight=1)
    app.selection_label = ctk.CTkLabel(selected, text="选择结果可查看详情", text_color=ui.MUTED, anchor="w", font=("Microsoft YaHei UI", 11))
    app.selection_label.grid(row=0, column=0, padx=16, pady=4, sticky="ew")
    quiet(selected, "复制内容", app._copy_selected, 84).grid(row=0, column=1, padx=(4, 8))
    app.selection_open_button = quiet(selected, "打开 / 定位", app._open_selected_target, 104)
    app.selection_open_button.grid(row=0, column=2, padx=(0, 12))
    app.preview_pane = ResultPreviewPane(body, on_open=app._open_selected_target, on_directory=app._open_selected_directory,
                                        on_copy_path=lambda: app._copy_selected("path"))
    app.preview_pane.configure(width=330)
    app.preview_pane.grid(row=0, column=1, padx=(14, 0), sticky="nsew")
    app.preview_pane.grid_propagate(False)

    footer = ctk.CTkFrame(app, fg_color="transparent")
    footer.grid(row=4, column=0, padx=24, pady=(6, 8), sticky="ew")
    footer.grid_columnconfigure(0, weight=1)
    app.status_label = ctk.CTkLabel(footer, text="请选择目录后执行检索", text_color=ui.MUTED, anchor="w", justify="left", font=("Microsoft YaHei UI", 11))
    app.status_label.grid(row=0, column=0, sticky="ew")
    app.failure_details_button = quiet(footer, "失败明细", app.show_failure_details, 110)
    app.failure_details_button.grid(row=0, column=1, padx=8)
    app.failure_details_button.grid_remove()
    ctk.CTkLabel(footer, text="仅检索磁盘上已保存的内容", text_color=ui.SUBTLE, font=("Microsoft YaHei UI", 11)).grid(row=0, column=2, padx=(12, 0))
    footer.bind("<Configure>", lambda event: app.status_label.configure(wraplength=max(280, event.width - 340)), add="+")
    app.bind("<Control-f>", app._focus_local_filter)
    app.bind("<Control-l>", lambda _event: app.folder_entry.focus_set())
    app.bind("<Escape>", lambda _event: app._reset_local_filter())
