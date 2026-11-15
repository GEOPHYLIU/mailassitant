#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""邮件助手 —— read_eml 的可视化工作台（Tkinter，仅标准库）。

界面与命令行共用同一套逻辑（read_eml.run_batch）：
    选目录 → 跑批 → 左侧列表 → 右侧「原文 / 元信息」对照，元信息页下方可直接
    双击打开附件。

界面支持中英文切换：切到英文时，界面文字与输出表格的表头一并变英文。
用法：python read_eml_gui.py
"""

from __future__ import annotations

import argparse
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import read_eml as core

APP_DIR = Path(os.environ.get("APPDATA") or Path.home()) / "eml_workbench"
CONFIG_PATH = APP_DIR / "config.json"

# 视觉口径参照 Windows 11 设置页：浅灰底 + 白卡片 + 细描边 + 微软雅黑
FONT_UI = "Microsoft YaHei UI"
THEME = {
    "bg": "#F3F3F3",
    "card": "#FFFFFF",
    "border": "#E1E1E1",
    "text": "#1B1B1B",
    "muted": "#616161",
    "control": "#FFFFFF",
    "control_hover": "#F5F5F5",
    "accent": "#0067C0",
    "accent_hover": "#005BA8",
    "accent_disabled": "#A0C3E3",
    "select": "#E5F1FB",
    "head": "#FAFAFA",
}

# 分类的内部键（与核心引擎一致）→ 显示名
# 分类键从核心规则表派生，另加 ICS 派生的会议类与兜底的其他类
CATEGORIES = [name for name, _ in core.CATEGORY_RULES] + ["会议类", "其他类"]

# 删除原因 → 底色（原文页签里的高亮）
REASON_COLORS = {
    "套话/免责声明": "#FFF3BF",
    "历史转发/引用": "#E5DBFF",
    "签名档": "#D3F9D8",
    "收尾语/寒暄": "#FFE8CC",
    "整封无自写正文": "#FFE3E3",
}

REASON_EN = {
    "套话/免责声明": "Boilerplate/disclaimer",
    "历史转发/引用": "Quoted history",
    "签名档": "Signature",
    "收尾语/寒暄": "Closing/salutation",
    "整封无自写正文": "No original body",
}

STRINGS: dict[str, dict[str, str]] = {
    "zh": {
        "app_name": "邮件助手",
        "input_dir": "邮件目录",
        "output_dir": "输出目录",
        "choose": "选择…",
        "recursive": "递归子目录",
        "attachments": "导出附件",
        "full": "完整清单（{n} 列）",
        "rebuild": "强制重建",
        "summary": "分类汇总",
        "calendar": "日历邀请",
        "language": "界面语言",
        "start": "开始处理",
        "cancel": "取消",
        "open_output": "打开输出目录",
        "export_filtered": "导出筛选结果",
        "status_idle": "待开始",
        "status_running": "正在处理…",
        "status_cancelling": "正在取消…",
        "status_cancelled": "已取消",
        "status_failed": "运行失败",
        "status_copied": "已复制当前页内容",
        "category": "分类",
        "all": "全部",
        "search": "搜索",
        "col_category": "分类",
        "col_received": "收件时间",
        "col_from": "发件人",
        "col_subject": "主题",
        "count": "共 {total} 封，当前显示 {shown} 封",
        "tab_raw": "原文",
        "tab_meta": "元信息",
        "copy": "复制当前页内容",
        "legend_none": "删除：无（正文未做删减）",
        "legend": "删除：{items}",
        "legend_unlocated": "{reason}(未定位)",
        "no_subject": "(无主题)",
        "no_date": "(无日期)",
        "no_body": "（无正文）",
        "raw_missing": "（原文不可用：源邮件已被移走）",
        "raw_failed": "读取原文失败：{error}",
        "meta_file": "文件",
        "meta_subject": "主题",
        "meta_from": "发件人",
        "meta_to": "收件人",
        "meta_cc": "抄送",
        "meta_date": "日期(原始)",
        "meta_local": "本机时间",
        "meta_category": "分类",
        "meta_chars": "正文字符",
        "meta_attachments": "附件",
        "meta_parse_error": "解析错误",
        "meta_event": "日历事件",
        "meta_basis": "依据",
        "pick_input": "选择邮件所在目录",
        "pick_output": "选择输出目录",
        "no_output_dir": "输出目录还不存在：{path}",
        "nothing_to_export": "当前没有可导出的记录",
        "export_title": "导出当前筛选结果",
        "export_filetype": "Excel 工作簿",
        "export_default": "eml_result_筛选_{n}封.xlsx",
        "export_failed": "导出失败：{error}",
        "exported": "已导出筛选结果 {n} 封 → {path}",
        "exported_status": "已导出 {n} 封",
        "need_input": "请先选择邮件目录",
        "dir_missing": "目录不存在：{path}",
        "progress_phase": "{phase} {done}/{total}",
        "cancelled_log": "已取消（已写入的部分结果保留在输出目录）",
        "failed_log": "运行失败：{error}",
        "fallback_log": "注意：主文件被占用，已另存为 {name}",
        "fallback_status": "完成（主文件被占用，结果另存 {name}）",
        "done_status": "完成：{mode}（{reason}），{n} 封",
        "loaded_log": "已载入上次结果：{n} 封（{path}）",
        "done_detail": "本轮：扫到 {files} 个文件，入表 {n} 封，重复 {dup} 封，日历 {ics} 个",
        "path_table": "清单：{path}",
        "path_summary": "汇总：{path}",
        "path_meeting": "日程：{path}",
        "path_index": "索引：{path}",
        "attach_box": "附件（双击用默认程序打开）",
        "attach_col_name": "文件名",
        "attach_col_size": "大小",
        "attach_empty": "本封没有附件",
        "attach_open_failed": "打不开这个附件：{error}",
        "attach_missing": "附件文件已不在这台机器上：{path}",
        "attach_menu_open": "打开",
        "attach_menu_folder": "在文件夹中显示",
        "attach_folder_failed": "打不开附件目录：{error}",
    },
    "en": {
        "app_name": "Mail Assistant",
        "input_dir": "Mail folder",
        "output_dir": "Output folder",
        "choose": "Browse…",
        "recursive": "Include subfolders",
        "attachments": "Export attachments",
        "full": "Full table ({n} columns)",
        "rebuild": "Force rebuild",
        "summary": "Summary",
        "calendar": "Calendar invites",
        "language": "Language",
        "start": "Start",
        "cancel": "Cancel",
        "open_output": "Open output folder",
        "export_filtered": "Export filtered",
        "status_idle": "Ready",
        "status_running": "Working…",
        "status_cancelling": "Cancelling…",
        "status_cancelled": "Cancelled",
        "status_failed": "Failed",
        "status_copied": "Current tab copied",
        "category": "Category",
        "all": "All",
        "search": "Search",
        "col_category": "Category",
        "col_received": "Received",
        "col_from": "From",
        "col_subject": "Subject",
        "count": "{total} emails, {shown} shown",
        "tab_raw": "Original",
        "tab_meta": "Details",
        "copy": "Copy current tab",
        "legend_none": "Removed: none (body untouched)",
        "legend": "Removed: {items}",
        "legend_unlocated": "{reason} (not located)",
        "no_subject": "(no subject)",
        "no_date": "(no date)",
        "no_body": "(no body)",
        "raw_missing": "(Original unavailable: source email has been moved)",
        "raw_failed": "Failed to read original: {error}",
        "meta_file": "File",
        "meta_subject": "Subject",
        "meta_from": "From",
        "meta_to": "To",
        "meta_cc": "Cc",
        "meta_date": "Date (raw)",
        "meta_local": "Local time",
        "meta_category": "Category",
        "meta_chars": "Body characters",
        "meta_attachments": "Attachments",
        "meta_parse_error": "Parse error",
        "meta_event": "Calendar event",
        "meta_basis": "matched by",
        "pick_input": "Choose the mail folder",
        "pick_output": "Choose the output folder",
        "no_output_dir": "Output folder does not exist yet: {path}",
        "nothing_to_export": "Nothing to export right now",
        "export_title": "Export filtered results",
        "export_filetype": "Excel workbook",
        "export_default": "eml_result_filtered_{n}.xlsx",
        "export_failed": "Export failed: {error}",
        "exported": "Exported {n} filtered emails → {path}",
        "exported_status": "Exported {n}",
        "need_input": "Choose the mail folder first",
        "dir_missing": "Folder not found: {path}",
        "progress_phase": "{phase} {done}/{total}",
        "cancelled_log": "Cancelled (partial results kept in the output folder)",
        "failed_log": "Run failed: {error}",
        "fallback_log": "Note: the main file was locked, saved as {name}",
        "fallback_status": "Done (file locked, saved as {name})",
        "done_status": "Done: {mode}, {n} emails",
        "loaded_log": "Loaded {n} previous results ({path})",
        "done_detail": "Run: {files} files scanned, {n} emails, {dup} duplicates, {ics} calendar invites",
        "path_table": "Table: {path}",
        "path_summary": "Summary: {path}",
        "path_meeting": "Meetings: {path}",
        "path_index": "Index: {path}",
        "attach_box": "Attachments (double-click to open)",
        "attach_col_name": "File name",
        "attach_col_size": "Size",
        "attach_empty": "This email has no attachments",
        "attach_open_failed": "Could not open the attachment: {error}",
        "attach_missing": "The attachment is no longer on this machine: {path}",
        "attach_menu_open": "Open",
        "attach_menu_folder": "Show in folder",
        "attach_folder_failed": "Could not open the attachment folder: {error}",
    },
}


def enable_dpi_awareness() -> None:
    """Windows 高分屏下避免界面发虚。"""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def load_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(data: dict) -> None:
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def default_out_dir() -> Path:
    """双击 exe 时的工作目录不可靠，默认落在程序自己旁边。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "eml_output"
    return Path(__file__).resolve().parent / "eml_output"


def open_folder(path: Path, select: Path | None = None) -> None:
    """打开文件夹；给了 select 时在资源管理器里选中该文件（不直接执行它）。"""
    if sys.platform == "win32":
        if select is not None and select.exists():
            subprocess.run(["explorer", "/select,", str(select)], check=False)
            return
        os.startfile(str(path))  # noqa: S606 - 用户主动点击打开自己指定的目录
        return
    subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", str(path)], check=False)


class Workbench:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.config = load_config()
        self.lang = "en" if str(self.config.get("lang", "zh")).lower() == "en" else "zh"
        self.records: list[dict] = []
        self.visible: list[dict] = []
        self.trace_cache: dict[str, list[dict]] = {}
        self.worker: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self.events: queue.Queue = queue.Queue()
        self.last_result: dict = {}
        self.sort_state: dict[str, bool] = {}
        self.search_job: str | None = None
        self.selected_path: str = ""
        self._attachment_rows: list[dict] = []

        self.root.title(self.L["app_name"])
        self._apply_geometry()
        self.root.minsize(900, 600)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self._build_ui()
        self._load_previous_results()
        self.pump_events()

    # ---------------------------------------------------------------- 语言

    @property
    def L(self) -> dict[str, str]:
        return STRINGS[self.lang]

    def category_text(self, key: str, lang: str | None = None) -> str:
        """分类的内部键 → 当前语言的显示名。"""
        return core._category_label(key, lang or self.lang)

    def set_language(self, choice: str) -> None:
        lang = "en" if choice == "English" else "zh"
        if lang == self.lang:
            return
        keyword = self.var_search.get()
        category = self.var_category_key()
        self.lang = lang
        self.config["lang"] = lang
        save_config(self.config)
        for child in self.root.winfo_children():
            child.destroy()
        self.trace_cache.clear()
        self._build_ui()
        self.var_search.set(keyword)
        self.var_category.set(self.category_text(category) if category else self.L["all"])
        self.root.title(self.L["app_name"])
        self.apply_filter()
        self._restore_selection()
        record = self.current_record()
        if record is not None:
            self.show_record(record)
        else:
            self._refresh_attachments(None)

    def _restore_selection(self) -> None:
        for index, record in enumerate(self.visible):
            if record.get("file") == self.selected_path:
                self.tree.selection_set(str(index))
                self.tree.see(str(index))
                return

    # ---------------------------------------------------------------- 界面

    def _apply_geometry(self) -> None:
        """窗口不超出屏幕：换过屏幕后保存的尺寸可能比当前屏幕还大。"""
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        max_w = screen_w - 100   # 留出窗口边框与阴影，不然右边缘会压出屏幕
        max_h = screen_h - 120
        saved = self.config.get("geometry") or ""
        if saved and "x" in saved:
            try:
                size, x, y = saved.replace("-", "+-").split("+")
                width, height = (int(v) for v in size.split("x"))
                pos_x, pos_y = int(x), int(y)
            except ValueError:
                width = height = 0
            else:
                width = min(width, max_w)
                height = min(height, max_h)
                if width >= 900 and height >= 560:
                    pos_x = max(0, min(pos_x, screen_w - width - 20))
                    pos_y = max(0, min(pos_y, screen_h - height - 60))
                    self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
                    return
        width = min(1180, max(900, max_w))
        height = min(740, max(600, max_h))
        self.root.geometry(f"{width}x{height}")

    def _card(self, parent, padding=(16, 12, 16, 12)):
        """白卡片 + 1px 浅描边：外框 tk.Frame 上色，内层 ttk.Frame 铺内容。"""
        holder = tk.Frame(parent, bg=THEME["border"], highlightthickness=0)
        inner = ttk.Frame(holder, padding=padding, style="Card.TFrame")
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        return holder, inner

    def _build_ui(self) -> None:
        L = self.L
        self.root.configure(bg=THEME["bg"])
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        font = (FONT_UI, 9)
        # 参照 Windows 11 设置页：浅灰底、白卡片、细描边、微软雅黑/9 号正文
        style.configure("TFrame", background=THEME["card"])
        style.configure("Card.TFrame", background=THEME["card"])
        style.configure("Bg.TFrame", background=THEME["bg"])
        style.configure("TLabel", font=font, background=THEME["card"],
                        foreground=THEME["text"])
        style.configure("Muted.TLabel", font=font, background=THEME["card"],
                        foreground=THEME["muted"])
        style.configure("Title.TLabel", font=(FONT_UI, 11, "bold"),
                        background=THEME["card"], foreground=THEME["text"])
        style.configure("TButton", font=font, padding=(12, 5),
                        background=THEME["control"], foreground=THEME["text"],
                        bordercolor=THEME["border"], relief="flat", focusthickness=0)
        style.map("TButton",
                  background=[("active", THEME["control_hover"]),
                              ("disabled", THEME["control"])],
                  foreground=[("disabled", THEME["muted"])])
        style.configure("Accent.TButton", font=(FONT_UI, 9), padding=(14, 5),
                        background=THEME["accent"], foreground="#FFFFFF",
                        bordercolor=THEME["accent"], relief="flat", focusthickness=0)
        style.map("Accent.TButton",
                  background=[("active", THEME["accent_hover"]),
                              ("disabled", THEME["accent_disabled"])],
                  foreground=[("disabled", "#FFFFFF")])
        style.configure("TCheckbutton", font=font, background=THEME["card"],
                        foreground=THEME["text"], focusthickness=0)
        style.map("TCheckbutton", background=[("active", THEME["card"])])
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground=THEME["text"],
                        bordercolor=THEME["border"], lightcolor=THEME["border"],
                        darkcolor=THEME["border"], insertcolor=THEME["text"], padding=(6, 4))
        style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF",
                        bordercolor=THEME["border"], arrowcolor=THEME["text"], padding=(6, 3))
        style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")])
        style.configure("Treeview", rowheight=26, font=font,
                        background="#FFFFFF", fieldbackground="#FFFFFF",
                        foreground=THEME["text"], bordercolor=THEME["border"])
        style.map("Treeview",
                  background=[("selected", THEME["select"])],
                  foreground=[("selected", THEME["text"])])
        style.configure("Treeview.Heading", font=(FONT_UI, 9, "bold"),
                        background=THEME["head"], foreground=THEME["text"],
                        relief="flat", bordercolor=THEME["border"], padding=(6, 5))
        style.map("Treeview.Heading", background=[("active", THEME["head"]),
                                                  ("pressed", THEME["head"])])
        style.configure("TNotebook", background=THEME["card"], borderwidth=0,
                        tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab", font=font, padding=(14, 6),
                        background=THEME["bg"], foreground=THEME["text"],
                        bordercolor=THEME["border"])
        style.map("TNotebook.Tab",
                  background=[("selected", "#FFFFFF")],
                  foreground=[("selected", THEME["text"])])
        style.configure("TLabelframe", background=THEME["card"],
                        bordercolor=THEME["border"], relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", font=font, background=THEME["card"],
                        foreground=THEME["muted"])
        style.configure("TSeparator", background=THEME["border"])
        style.configure("TProgressbar", background=THEME["accent"],
                        troughcolor=THEME["bg"], bordercolor=THEME["border"],
                        lightcolor=THEME["accent"], darkcolor=THEME["accent"])
        # 滚动条与分栏分隔条：去掉两端箭头、细窄浅色滑块，贴近 Win11 观感
        for orient, key in (("Vertical", "ns"), ("Horizontal", "we")):
            name = f"{orient}.TScrollbar"
            element = f"{orient}.Scrollbar"
            style.layout(name, [
                (f"{element}.trough", {
                    "children": [(f"{element}.thumb", {"expand": "1", "sticky": key})],
                    "sticky": key,
                })
            ])
            style.configure(name, troughcolor=THEME["card"], background="#C9C9C9",
                            bordercolor=THEME["card"], lightcolor="#C9C9C9",
                            darkcolor="#C9C9C9", gripcount=0, arrowcolor=THEME["card"])
            style.map(name,
                      background=[("active", "#A6A6A6"), ("pressed", "#8C8C8C")])
        style.configure("TPanedwindow", background=THEME["border"], sashthickness=6)
        style.configure("Sash", background=THEME["border"], bordercolor=THEME["border"],
                        lightcolor=THEME["border"], darkcolor=THEME["border"])

        self.root.rowconfigure(1, weight=1)
        self.root.rowconfigure(2, weight=0)
        self.root.columnconfigure(0, weight=1)

        # ---- 顶部：目录与开关（白卡片）
        top_card, top = self._card(self.root, padding=(16, 12, 16, 12))
        top_card.grid(row=0, column=0, sticky="we", padx=12, pady=(12, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text=L["input_dir"]).grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.var_input = tk.StringVar(value=self.config.get("input_dir") or "")
        ttk.Entry(top, textvariable=self.var_input).grid(row=0, column=1, sticky="we")
        ttk.Button(top, text=L["choose"], command=self.pick_input).grid(row=0, column=2, padx=6)

        ttk.Label(top, text=L["output_dir"]).grid(row=1, column=0, sticky="w", padx=(0, 6),
                                                 pady=(4, 0))
        self.var_out = tk.StringVar(value=self.config.get("output_dir") or str(default_out_dir()))
        ttk.Entry(top, textvariable=self.var_out).grid(row=1, column=1, sticky="we", pady=(4, 0))
        ttk.Button(top, text=L["choose"], command=self.pick_output).grid(
            row=1, column=2, padx=6, pady=(4, 0))

        opts = ttk.Frame(top)
        opts.grid(row=2, column=0, columnspan=3, sticky="we", pady=(8, 0))
        self.var_recursive = tk.BooleanVar(value=bool(self.config.get("recursive", False)))
        self.var_attach = tk.BooleanVar(value=bool(self.config.get("attachments", True)))
        self.var_full = tk.BooleanVar(value=bool(self.config.get("full", False)))
        self.var_rebuild = tk.BooleanVar(value=bool(self.config.get("rebuild", False)))
        self.var_summary = tk.BooleanVar(value=bool(self.config.get("summary", True)))
        self.var_calendar = tk.BooleanVar(value=bool(self.config.get("calendar", True)))
        ttk.Checkbutton(opts, text=L["recursive"], variable=self.var_recursive).pack(side="left")
        ttk.Checkbutton(opts, text=L["attachments"], variable=self.var_attach).pack(
            side="left", padx=(12, 0))
        ttk.Checkbutton(opts, text=L["summary"], variable=self.var_summary).pack(
            side="left", padx=(12, 0))
        ttk.Checkbutton(opts, text=L["calendar"], variable=self.var_calendar).pack(
            side="left", padx=(12, 0))
        ttk.Checkbutton(opts, text=L["full"].format(n=len(core.CSV_FIELDS)),
                       variable=self.var_full).pack(
            side="left", padx=(12, 0))
        ttk.Checkbutton(opts, text=L["rebuild"], variable=self.var_rebuild).pack(
            side="left", padx=(12, 0))
        self.var_lang = tk.StringVar(value="English" if self.lang == "en" else "中文")
        lang_box = ttk.Combobox(opts, textvariable=self.var_lang, values=["中文", "English"],
                                width=8, state="readonly")
        lang_box.pack(side="right")
        ttk.Label(opts, text=L["language"]).pack(side="right", padx=(12, 4))
        lang_box.bind("<<ComboboxSelected>>", lambda _e: self.set_language(self.var_lang.get()))

        actions = ttk.Frame(top)
        actions.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.btn_start = ttk.Button(actions, text=L["start"], command=self.start,
                                    style="Accent.TButton")
        self.btn_start.pack(side="left")
        self.btn_cancel = ttk.Button(actions, text=L["cancel"], command=self.cancel,
                                     state="disabled")
        self.btn_cancel.pack(side="left", padx=(8, 0))
        ttk.Button(actions, text=L["open_output"], command=self.open_output).pack(
            side="left", padx=(8, 0))
        ttk.Button(actions, text=L["export_filtered"], command=self.export_filtered).pack(
            side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(top, mode="determinate", maximum=100)
        self.progress.grid(row=4, column=0, columnspan=3, sticky="we", pady=(12, 0))
        self.var_status = tk.StringVar(value=L["status_idle"])
        ttk.Label(top, textvariable=self.var_status, style="Muted.TLabel").grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # ---- 中部：左列表 + 右详情
        pane = ttk.Panedwindow(self.root, orient="horizontal")
        pane.grid(row=1, column=0, sticky="nsew", padx=6, pady=(6, 6))

        left_card, left = self._card(pane, padding=(14, 12, 14, 12))
        left.rowconfigure(2, weight=1)
        left.columnconfigure(0, weight=1)

        filters = ttk.Frame(left)
        filters.grid(row=0, column=0, sticky="we")
        ttk.Label(filters, text=L["category"]).pack(side="left")
        self.var_category = tk.StringVar(value=L["all"])
        combo = ttk.Combobox(filters, textvariable=self.var_category, width=12, state="readonly",
                             values=[L["all"]] + [self.category_text(k) for k in CATEGORIES])
        combo.pack(side="left", padx=(4, 10))
        combo.bind("<<ComboboxSelected>>", lambda _e: self.apply_filter())
        ttk.Label(filters, text=L["search"]).pack(side="left")
        self.var_search = tk.StringVar()
        entry = ttk.Entry(filters, textvariable=self.var_search)
        entry.pack(side="left", fill="x", expand=True, padx=(4, 0))
        entry.bind("<KeyRelease>", self.on_search_changed)

        self.var_count = tk.StringVar(value=L["count"].format(total=0, shown=0))
        ttk.Label(left, textvariable=self.var_count, style="Muted.TLabel").grid(
            row=1, column=0, sticky="w", pady=(6, 4))

        columns = ("category", "received", "from", "subject")
        self.tree = ttk.Treeview(left, columns=columns, show="headings", selectmode="browse",
                                 height=6)
        headings = {"category": (L["col_category"], 92), "received": (L["col_received"], 118),
                    "from": (L["col_from"], 150), "subject": (L["col_subject"], 260)}
        for key, (text, width) in headings.items():
            self.tree.heading(key, text=text, command=lambda k=key: self.sort_by(k))
            self.tree.column(key, width=width, minwidth=50, anchor="w",
                             stretch=(key == "subject"))
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        hscroll = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hscroll.set)
        self.tree.grid(row=2, column=0, sticky="nsew")
        scroll.grid(row=2, column=1, sticky="ns")
        hscroll.grid(row=3, column=0, sticky="we")
        self.tree.tag_configure("error", foreground="#C92A2A")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        right_card, right = self._card(pane, padding=(14, 12, 14, 12))
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        self.notebook = ttk.Notebook(right)
        self.notebook.grid(row=0, column=0, columnspan=2, sticky="nsew")

        self.text_raw = self._make_readonly_tab(L["tab_raw"])
        self.text_meta, self.attach_tree, self.attach_box = self._make_meta_tab(L["tab_meta"])

        detail_actions = ttk.Frame(right)
        detail_actions.grid(row=1, column=0, columnspan=2, sticky="we", pady=(8, 0))
        ttk.Button(detail_actions, text=L["copy"], command=self.copy_current).pack(side="left")
        self.var_legend = tk.StringVar(value="")
        ttk.Label(detail_actions, textvariable=self.var_legend,
                  style="Muted.TLabel").pack(side="left", padx=(12, 0))

        pane.add(left_card, weight=3)
        pane.add(right_card, weight=4)
        self.pane = pane
        self.root.after_idle(self._place_sash)

        # ---- 底部：日志（白卡片）
        log_card, log_frame = self._card(self.root, padding=(14, 8, 14, 10))
        log_card.grid(row=2, column=0, sticky="we", padx=12, pady=(0, 12))
        log_frame.columnconfigure(0, weight=1)
        self.log = ScrolledText(log_frame, height=4, wrap="word",
                                font=("Consolas", 9), state="disabled",
                                bg="#FFFFFF", fg=THEME["text"], relief="flat",
                                borderwidth=0, highlightthickness=1,
                                highlightbackground=THEME["border"],
                                padx=6, pady=4)
        self.log.grid(row=0, column=0, sticky="we")
        self._style_text_scrollbars(self.log)

    def _make_readonly_tab(self, title: str) -> ScrolledText:
        frame = ttk.Frame(self.notebook)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        widget = ScrolledText(frame, wrap="word", font=(FONT_UI, 9),
                              height=4, width=40, undo=False, state="disabled",
                              bg="#FFFFFF", fg=THEME["text"], relief="flat",
                              borderwidth=0, highlightthickness=1,
                              highlightbackground=THEME["border"], padx=8, pady=6)
        widget.grid(row=0, column=0, sticky="nsew")
        self._style_text_scrollbars(widget)
        self.notebook.add(frame, text=title)
        return widget

    def _make_meta_tab(self, title: str):
        """元信息页签：上面是字段文本，下面是可以双击打开的附件列表。

        两块放进竖直分栏，用户可以把更多高度分给字段文本或附件列表。
        """
        L = self.L
        frame = ttk.Frame(self.notebook)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        split = ttk.Panedwindow(frame, orient="vertical")
        split.grid(row=0, column=0, sticky="nsew")

        text_holder = ttk.Frame(split)
        text_holder.rowconfigure(0, weight=1)
        text_holder.columnconfigure(0, weight=1)
        widget = ScrolledText(text_holder, wrap="word", font=(FONT_UI, 9),
                              height=4, width=40, undo=False, state="disabled",
                              bg="#FFFFFF", fg=THEME["text"], relief="flat",
                              borderwidth=0, highlightthickness=1,
                              highlightbackground=THEME["border"], padx=8, pady=6)
        widget.grid(row=0, column=0, sticky="nsew")
        self._style_text_scrollbars(widget)

        box = ttk.LabelFrame(split, text=L["attach_box"], padding=(6, 4))
        box.columnconfigure(0, weight=1)
        tree = ttk.Treeview(box, columns=("name", "size"), show="headings",
                            selectmode="browse", height=3)
        tree.heading("name", text=L["attach_col_name"])
        tree.heading("size", text=L["attach_col_size"])
        tree.column("name", width=260, minwidth=120, anchor="w", stretch=True)
        tree.column("size", width=90, minwidth=60, anchor="e", stretch=False)
        tree.grid(row=0, column=0, sticky="we")
        bar = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=bar.set)
        bar.grid(row=0, column=1, sticky="ns")
        tree.bind("<Double-1>", self.open_attachment)
        tree.bind("<Return>", self.open_attachment)
        tree.bind("<Button-3>", self._attachment_menu)
        self.attach_menu = tk.Menu(tree, tearoff=0)
        self.attach_menu.add_command(label=L["attach_menu_open"],
                                     command=lambda: self.open_attachment(None))
        self.attach_menu.add_command(label=L["attach_menu_folder"],
                                     command=self.show_attachment_folder)

        split.add(text_holder, weight=3)
        split.add(box, weight=2)
        self.meta_split = split
        split.after_idle(lambda: self._place_meta_sash(split))

        self.notebook.add(frame, text=title)
        return widget, tree, box

    # -------------------------------------------------------------- 附件

    def _refresh_attachments(self, record: dict | None) -> None:
        tree = getattr(self, "attach_tree", None)
        if tree is None:
            return
        tree.delete(*tree.get_children())
        self._attachment_rows = []
        items = (record or {}).get("attachments") or []
        if not items:
            tree.insert("", "end", iid="none", values=(self.L["attach_empty"], ""),
                        tags=("empty",))
            tree.tag_configure("empty", foreground="#888")
            return
        tree.tag_configure("empty", foreground="#888")
        for index, item in enumerate(items):
            if not isinstance(item, dict):
                continue
            name = item.get("filename", "")
            size = item.get("size_bytes", 0)
            try:
                size_text = self._size_text(int(size))
            except (TypeError, ValueError):
                size_text = ""
            iid = f"a{index}"
            tree.insert("", "end", iid=iid, values=(name, size_text))
            self._attachment_rows.append({"iid": iid, "path": item.get("saved_path", ""),
                                          "error": item.get("save_error", "")})

    @staticmethod
    def _size_text(size: int) -> str:
        if size >= 1024 * 1024:
            return f"{size / (1024 * 1024):.1f} MB"
        if size >= 1024:
            return f"{size / 1024:.1f} KB"
        return f"{size} B"

    def _selected_attachment(self) -> dict | None:
        tree = getattr(self, "attach_tree", None)
        if tree is None:
            return None
        selection = tree.selection()
        if not selection:
            return None
        iid = selection[0]
        for item in getattr(self, "_attachment_rows", []):
            if item["iid"] == iid:
                return item
        return None

    def open_attachment(self, _event=None) -> None:
        item = self._selected_attachment()
        if not item:
            return
        path = Path(item["path"]) if item["path"] else None
        if path is None or not path.is_file():
            reason = item.get("error")
            self.var_status.set(self.L["attach_open_failed"].format(error=reason) if reason
                                else self.L["attach_missing"].format(path=item["path"] or "-"))
            return
        try:
            os.startfile(str(path))
        except OSError as exc:
            self.var_status.set(self.L["attach_open_failed"].format(error=exc))

    def show_attachment_folder(self) -> None:
        item = self._selected_attachment()
        if not item or not item["path"]:
            return
        path = Path(item["path"])
        try:
            if path.is_file():
                subprocess.Popen(["explorer", "/select,", str(path)])
            elif path.parent.is_dir():
                os.startfile(str(path.parent))
            else:
                self.var_status.set(self.L["attach_missing"].format(path=str(path)))
        except OSError as exc:
            self.var_status.set(self.L["attach_folder_failed"].format(error=exc))

    def _attachment_menu(self, event) -> None:
        row = self.attach_tree.identify_row(event.y)
        if row:
            self.attach_tree.selection_set(row)
            self.attach_menu.tk_popup(event.x_root, event.y_root)
            self.attach_menu.grab_release()

    # ---------------------------------------------------------------- 数据

    def _place_sash(self) -> None:
        """左右分栏按比例落位：只靠 weight 时右侧详情会被挤出窗口。"""
        pane = getattr(self, "pane", None)
        if pane is None:
            return
        width = pane.winfo_width()
        if width < 200:
            self.root.after(80, self._place_sash)
            return
        try:
            self.pane.sashpos(0, int(width * 0.56))
        except tk.TclError:
            pass

    def _style_text_scrollbars(self, widget) -> None:
        """ScrolledText 自带的滚动条是 tk.Scrollbar，单独调成细窄浅色样式。"""
        for bar in (getattr(widget, "vbar", None), getattr(widget, "hbar", None)):
            if bar is None:
                continue
            try:
                bar.configure(width=11, borderwidth=0, relief="flat",
                              elementborderwidth=0, highlightthickness=0,
                              troughcolor="#FFFFFF", background="#C9C9C9",
                              activebackground="#A6A6A6")
            except tk.TclError:
                pass

    def _place_meta_sash(self, split) -> None:
        """元信息页：字段文本占六成，附件列表占四成。"""
        height = split.winfo_height()
        if height < 80:
            split.after(80, lambda: self._place_meta_sash(split))
            return
        try:
            split.sashpos(0, int(height * 0.58))
        except tk.TclError:
            pass

    def _load_previous_results(self) -> None:
        """冷启动：把上次跑出的索引读回列表，否则界面每次都是空的。"""
        out_dir = Path(self.var_out.get() or default_out_dir())
        index_path = out_dir / core.INDEX_FILE
        records: list[dict] = []
        source = ""
        if index_path.is_file():
            try:
                payload = json.loads(index_path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                payload = {}
            records = payload.get("records") or []
            source = str(index_path)
        if not records:
            # 兼容旧输出目录：早期版本写的是 eml_result.jsonl
            legacy = out_dir / "eml_result.jsonl"
            if legacy.is_file():
                try:
                    with legacy.open(encoding="utf-8") as handle:
                        for line in handle:
                            line = line.strip()
                            if not line:
                                continue
                            try:
                                records.append(json.loads(line))
                            except json.JSONDecodeError:
                                continue
                except OSError:
                    records = []
                source = str(legacy)
        if not records:
            return
        self.records = core.sort_records(records)
        for record in self.records:
            if not record.get("category"):
                record["category"], record["category_hit"] = core.classify(record)
        self.apply_filter()
        self.append_log(self.L["loaded_log"].format(n=len(records), path=source))

    def var_category_key(self) -> str:
        """当前下拉选项 → 分类内部键；“全部”返回 ""。"""
        shown = self.var_category.get()
        for key in CATEGORIES:
            if shown in (key, self.category_text(key, "zh"), self.category_text(key, "en")):
                return key
        return ""

    def apply_filter(self) -> None:
        category = self.var_category_key()
        keyword = (self.var_search.get() or "").strip().lower()
        visible = []
        for record in self.records:
            if category and record.get("category") != category:
                continue
            if keyword:
                haystack = " ".join(str(record.get(key, "")) for key in
                                    ("subject", "from", "from_addr", "to", "body_text")).lower()
                if keyword not in haystack:
                    continue
            visible.append(record)
        self.visible = visible
        self.refresh_tree()

    def refresh_tree(self) -> None:
        selected = self.current_record()
        self.tree.delete(*self.tree.get_children())
        for index, record in enumerate(self.visible):
            tags = ("error",) if record.get("parse_error") else ()
            self.tree.insert("", "end", iid=str(index), tags=tags, values=(
                self.category_text(record.get("category", "")),
                core._received_text(record),
                record.get("from", ""),
                record.get("subject", "") or self.L["no_subject"],
            ))
        self.var_count.set(self.L["count"].format(total=len(self.records),
                                                  shown=len(self.visible)))
        if selected is None or self.visible.count(selected) == 0:
            self.show_record(None)
        else:
            for index, record in enumerate(self.visible):
                if record is selected:
                    self.tree.selection_set(str(index))
                    break

    def on_search_changed(self, _event=None) -> None:
        if self.search_job is not None:
            self.root.after_cancel(self.search_job)
        self.search_job = self.root.after(300, self.apply_filter)

    def sort_by(self, key: str) -> None:
        reverse = not self.sort_state.get(key, False)
        self.sort_state[key] = reverse

        def value(record: dict):
            if key == "received":
                return record.get("date", "") or ""
            if key == "category":
                return record.get("category", "")
            if key == "from":
                return record.get("from", "")
            return record.get("subject", "")

        self.records.sort(key=value, reverse=reverse)
        self.apply_filter()

    def current_record(self) -> dict | None:
        selection = self.tree.selection()
        if not selection:
            return None
        try:
            return self.visible[int(selection[0])]
        except (IndexError, ValueError):
            return None

    def on_select(self, _event=None) -> None:
        record = self.current_record()
        if record is not None:
            self.selected_path = record.get("file", "")
        self.show_record(record)

    def show_record(self, record: dict | None) -> None:
        if record is None:
            self._set_text(self.text_raw, "")
            self._set_text(self.text_meta, "")
            self.var_legend.set("")
            self._refresh_attachments(None)
            return
        self._render_raw(record)
        self._set_text(self.text_meta, self._meta_text(record))
        self._refresh_attachments(record)

    def _meta_text(self, record: dict) -> str:
        L = self.L
        attachments = record.get("attachments") or []
        from_addr = record.get("from_addr") or ""
        sender = record.get("from", "")
        if from_addr and from_addr not in sender:
            sender = f"{sender} <{from_addr}>"
        category = self.category_text(record.get("category", ""))
        basis = record.get("category_hit", "")
        rows = [
            (L["meta_file"], record.get("file", "")),
            (L["meta_subject"], record.get("subject", "")),
            (L["meta_from"], sender),
            (L["meta_to"], record.get("to", "")),
            (L["meta_cc"], record.get("cc", "")),
            (L["meta_date"], record.get("date", "")),
            (L["meta_local"], core._received_text(record) or L["no_date"]),
            ("Message-ID", record.get("message_id", "")),
            (L["meta_category"], f"{category} ({L['meta_basis']}: {basis})"
             if self.lang == "en" else f"{category}（{L['meta_basis']}：{basis}）"),
            (L["meta_chars"], str(record.get("body_chars", 0))),
            (L["meta_attachments"], str(len(attachments))),
        ]
        width = max(len(label) for label, _ in rows) + 2
        lines = []
        for label, value in rows:
            text_value = str(value)
            sep = ": " if self.lang == "en" else "："
            if len(text_value) > 46:  # 长值另起一行，免得标签被挤成两行
                lines.append(f"{label}{sep}")
                lines.append(f"    {text_value}")
            else:
                lines.append(f"{label.ljust(width)}{sep}{text_value}")
        # 附件明细在页签下方的列表里（可双击打开），这里只报总数
        if record.get("parse_error"):
            lines.append(f"{L['meta_parse_error']}: {record['parse_error']}")
        for event in record.get("ics_events") or []:
            lines.append(f"{L['meta_event']}: {event.get('summary', '')} / {event.get('start', '')}")
            join_info = core._join_info(event)
            if join_info:
                lines.append(f"    {join_info}")
        for event in record.get("body_events") or []:
            start = event.get("start", "") or ("time not given" if self.lang == "en"
                                                else "正文未给出时间")
            lines.append(f"{L['meta_event']}: {event.get('summary', '')} / {start}")
            join_info = core._join_info(event)
            if join_info:
                lines.append(f"    {join_info}")
        return "\n".join(lines)

    def _render_raw(self, record: dict) -> None:
        """原文页签：显示清洗前正文，并把被删掉的段落按原因上色。"""
        path = record.get("file", "")
        widget = self.text_raw
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        for tag, color in REASON_COLORS.items():
            widget.tag_configure(tag, background=color)

        trace = None
        raw_text = ""
        cached = self.trace_cache.get(path)
        if cached is not None:
            raw_text, trace = cached
        elif path and Path(path).is_file():
            collected: list[dict] = []
            try:
                parsed = core.read_eml_file(path, trace=collected)
                raw_text = parsed.get("body_raw") or parsed.get("body_text") or ""
            except Exception as exc:
                raw_text = self.L["raw_failed"].format(
                    error=f"{type(exc).__name__}: {exc}")
            trace = collected
            self.trace_cache[path] = (raw_text, trace)
        else:
            raw_text = record.get("body_text") or ""
            trace = []
        widget.insert("1.0", raw_text or self.L["raw_missing"])

        cursor = 0
        counts: dict[str, int] = {}
        for item in trace or []:
            text = item.get("text", "")
            if not text:
                continue
            position = raw_text.find(text, cursor)
            if position < 0:
                position = raw_text.find(text)
            if position < 0:
                key = self.L["legend_unlocated"].format(reason=self.reason_text(item["reason"]))
                counts[key] = counts.get(key, 0) + 1
                continue
            tag = item["reason"] if item["reason"] in REASON_COLORS else "整封无自写正文"
            widget.tag_add(tag, f"1.0+{position}c", f"1.0+{position + len(text)}c")
            label = self.reason_text(item["reason"])
            counts[label] = counts.get(label, 0) + 1
            cursor = position + len(text)
        widget.configure(state="disabled")

        if counts:
            sep = "，" if self.lang == "zh" else ", "
            self.var_legend.set(self.L["legend"].format(
                items=sep.join(f"{k}×{v}" for k, v in counts.items())))
        else:
            self.var_legend.set(self.L["legend_none"])

    def reason_text(self, reason: str) -> str:
        return REASON_EN.get(reason, reason) if self.lang == "en" else reason

    def _set_text(self, widget: ScrolledText, value: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", value or "")
        widget.configure(state="disabled")

    # ---------------------------------------------------------------- 动作

    def pick_input(self) -> None:
        initial = self.var_input.get() or str(Path.home())
        chosen = filedialog.askdirectory(title=self.L["pick_input"], initialdir=initial)
        if chosen:
            self.var_input.set(chosen)

    def pick_output(self) -> None:
        initial = self.var_out.get() or str(default_out_dir())
        chosen = filedialog.askdirectory(title=self.L["pick_output"], initialdir=initial)
        if chosen:
            self.var_out.set(chosen)

    def open_output(self) -> None:
        out_dir = Path(self.var_out.get() or default_out_dir())
        if not out_dir.exists():
            messagebox.showinfo(self.L["app_name"],
                                self.L["no_output_dir"].format(path=out_dir))
            return
        open_folder(out_dir)

    def copy_current(self) -> None:
        widget = self.text_meta if self.notebook.index(self.notebook.select()) == 1 \
            else self.text_raw
        content = widget.get("1.0", "end").rstrip("\n")
        if not content:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.var_status.set(self.L["status_copied"])

    def export_filtered(self) -> None:
        if not self.visible:
            messagebox.showinfo(self.L["app_name"], self.L["nothing_to_export"])
            return
        out_dir = Path(self.var_out.get() or default_out_dir())
        default_name = self.L["export_default"].format(n=len(self.visible))
        target = filedialog.asksaveasfilename(
            title=self.L["export_title"], defaultextension=".xlsx",
            initialdir=str(out_dir), initialfile=default_name,
            filetypes=[(self.L["export_filetype"], "*.xlsx")])
        if not target:
            return
        try:
            if self.var_full.get():
                core.write_full_table(self.visible, target, with_body=True, lang=self.lang)
            else:
                core.write_brief_table(self.visible, target, lang=self.lang)
        except Exception as exc:
            messagebox.showerror(self.L["app_name"],
                                 self.L["export_failed"].format(
                                     error=f"{type(exc).__name__}: {exc}"))
            return
        self.append_log(self.L["exported"].format(n=len(self.visible), path=target))
        self.var_status.set(self.L["exported_status"].format(n=len(self.visible)))

    def start(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        input_dir = (self.var_input.get() or "").strip()
        if not input_dir:
            messagebox.showinfo(self.L["app_name"], self.L["need_input"])
            return
        if not Path(input_dir).exists():
            messagebox.showerror(self.L["app_name"],
                                 self.L["dir_missing"].format(path=input_dir))
            return
        out_dir = (self.var_out.get() or str(default_out_dir())).strip()
        self.cancel_event.clear()
        self.btn_start.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self.progress.configure(value=0)
        self.var_status.set(self.L["status_running"])

        args = argparse.Namespace(
            paths=[input_dir], out=out_dir,
            recursive=self.var_recursive.get(), pattern="*.eml",
            attachments=self.var_attach.get(), no_body=False,
            full=self.var_full.get(), no_summary=not self.var_summary.get(),
            no_calendar=not self.var_calendar.get(),
            quiet=False, rebuild=self.var_rebuild.get(), lang=self.lang,
        )

        def run() -> None:
            try:
                result = core.run_batch(
                    args,
                    log=lambda text: self.events.put(("log", text)),
                    progress=lambda phase, done, total: self.events.put(
                        ("progress", (phase, done, total))),
                    cancelled=self.cancel_event.is_set,
                )
                self.events.put(("done", result))
            except core.BatchCancelled:
                self.events.put(("cancelled", None))
            except Exception as exc:  # 把异常带回主线程显示，避免静默崩掉
                self.events.put(("failed", f"{type(exc).__name__}: {exc}"))

        self.worker = threading.Thread(target=run, daemon=True)
        self.worker.start()

    def cancel(self) -> None:
        if self.worker and self.worker.is_alive():
            self.cancel_event.set()
            self.var_status.set(self.L["status_cancelling"])

    def pump_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self.append_log(str(payload))
                elif kind == "progress":
                    phase, done, total = payload
                    percent = int(done * 100 / total) if total else 0
                    self.progress.configure(value=percent)
                    self.var_status.set(self.L["progress_phase"].format(
                        phase=phase, done=done, total=total))
                elif kind == "done":
                    self.on_done(payload)
                elif kind == "cancelled":
                    self.append_log(self.L["cancelled_log"])
                    self.var_status.set(self.L["status_cancelled"])
                    self.finish_buttons()
                elif kind == "failed":
                    self.append_log(self.L["failed_log"].format(error=payload))
                    self.var_status.set(self.L["status_failed"])
                    messagebox.showerror(self.L["app_name"],
                                         self.L["failed_log"].format(error=payload))
                    self.finish_buttons()
        except queue.Empty:
            pass
        self.root.after(120, self.pump_events)

    def on_done(self, result: dict) -> None:
        self.last_result = result or {}
        self.finish_buttons()
        if not result or not result.get("ok"):
            self.var_status.set(result.get("error", self.L["status_idle"]) if result
                                else self.L["status_idle"])
            return
        self.records = core.sort_records(result.get("records") or [])
        self.trace_cache.clear()
        self.apply_filter()
        self.progress.configure(value=100)
        fallback = self._detect_fallback(Path(result.get("out_dir", "")))
        if fallback:
            self.append_log(self.L["fallback_log"].format(name=fallback.name))
            self.var_status.set(self.L["fallback_status"].format(name=fallback.name))
        else:
            if self.lang == "en":
                mode_text = core.MODE_EN.get(result.get("mode", ""), result.get("mode", ""))
                self.var_status.set(self.L["done_status"].format(
                    mode=mode_text, n=len(self.records)))
            else:
                self.var_status.set(self.L["done_status"].format(
                    mode=result.get("mode", ""), reason=result.get("reason", ""),
                    n=len(self.records)))
        self.append_log(self.L["done_detail"].format(
            files=result.get("file_count", 0), n=len(self.records),
            dup=len(result.get("duplicates") or []), ics=result.get("ics_count", 0)))
        for key, path in (("path_table", result.get("table_path")),
                          ("path_summary", result.get("summary_path")),
                          ("path_meeting", result.get("meeting_path")),
                          ("path_index", result.get("index_path"))):
            if path:
                self.append_log("  " + self.L[key].format(path=path))

    def _detect_fallback(self, out_dir: Path) -> Path | None:
        """xlsx 被 Excel 占用时程序会另存 _new，这里把这件事显式告诉用户。"""
        if not out_dir or not out_dir.exists():
            return None
        main_file = out_dir / "eml_result.xlsx"
        candidates = sorted(out_dir.glob("eml_result_new*.xlsx"),
                            key=lambda item: item.stat().st_mtime, reverse=True)
        if not candidates:
            return None
        newest = candidates[0]
        if main_file.exists() and newest.stat().st_mtime <= main_file.stat().st_mtime:
            return None
        return newest

    def finish_buttons(self) -> None:
        self.btn_start.configure(state="normal")
        self.btn_cancel.configure(state="disabled")

    def append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def on_close(self) -> None:
        self.cancel_event.set()
        save_config({
            "input_dir": self.var_input.get(),
            "output_dir": self.var_out.get(),
            "recursive": self.var_recursive.get(),
            "attachments": self.var_attach.get(),
            "full": self.var_full.get(),
            "rebuild": self.var_rebuild.get(),
            "summary": self.var_summary.get(),
            "calendar": self.var_calendar.get(),
            "lang": self.lang,
            "geometry": self.root.geometry(),
        })
        self.root.destroy()


def main() -> int:
    enable_dpi_awareness()
    root = tk.Tk()
    Workbench(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
