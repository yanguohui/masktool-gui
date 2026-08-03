"""
Tkinter 图形界面
================

选用 Tkinter + ttk 而非 PyQt6，理由：
  · ttk 在 Windows 上走 vista 原生主题，观感即系统标准控件，符合"简洁专业"诉求；
  · 标准库自带，PyInstaller 单文件产物约 12 MB，PyQt6 方案通常 60 MB 以上，
    且 Qt 插件目录在单文件模式下极易出问题；
  · 冷启动更快，对"双击即用"的工具型软件体验更好。

界面所有耗时操作均在工作线程执行，通过 queue 回主线程刷新，保证不卡顿。
"""

from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import (
    BooleanVar, Listbox, StringVar, Tk, filedialog, messagebox, ttk,
)

from app import settings
from app.engine import (
    MODE_LABELS, MODES, REPORT_ONLY_EXTS, SUPPORTED_EXTS,
    SENSITIVITY_LEVELS, SENSITIVITY_KEYS, SENSITIVITY_DEFAULT,
    FileResult, MaskEngine, MaskToolNotFound, ToolInfo, locate_mask_tool,
)

APP_TITLE = "本地文档脱敏工具"
APP_VERSION = "1.0.0"

IS_WINDOWS = os.name == "nt"

INSTALL_HELP = (
    "未检测到脱敏核心 mask-tool。\n\n"
    "请按以下步骤安装（只需一次）：\n"
    "  1. 安装 Python 3.10 及以上版本，安装时勾选 “Add Python to PATH”\n"
    "  2. 下载 mask-tool 源码：\n"
    "     git clone https://github.com/ZagooYWX/mask-tool.git\n"
    "  3. 进入目录执行：\n"
    "     pip install -e .\n\n"
    "安装完成后点击“重新检测”。\n"
    "若已安装但仍无法识别，可点击“手动指定…”选择 mask-tool 可执行文件。"
)

# 用户词库类别：界面中文标签 -> mask-tool 词库类别键
LEX_CATEGORIES = [
    ("公司名", "company"),
    ("人名", "person"),
    ("项目名", "project"),
    ("机构 / 单位", "government"),
    ("地名", "location"),
    ("金额", "amount"),
    ("自定义", "custom"),
]


# --------------------------------------------------------------------------
# 平台适配
# --------------------------------------------------------------------------


def enable_dpi_awareness() -> None:
    """Windows 高分屏下避免界面发虚。"""
    if not IS_WINDOWS:
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # type: ignore[attr-defined]
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()       # type: ignore[attr-defined]
    except Exception:
        pass


def open_in_explorer(path: Path) -> None:
    """在系统文件管理器中定位文件/目录。"""
    try:
        if IS_WINDOWS:
            if path.is_file():
                subprocess.run(["explorer", "/select,", str(path)], check=False)
            else:
                os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            args = ["open", "-R", str(path)] if path.is_file() else ["open", str(path)]
            subprocess.run(args, check=False)
        else:
            target = path if path.is_dir() else path.parent
            subprocess.run(["xdg-open", str(target)], check=False)
    except Exception:
        messagebox.showwarning("提示", f"无法自动打开，请手动前往：\n{path}")


def ui_font() -> tuple[str, int]:
    if IS_WINDOWS:
        return ("Microsoft YaHei UI", 9)
    if sys.platform == "darwin":
        return ("PingFang SC", 13)
    return ("Noto Sans CJK SC", 10)


def human_size(p: Path) -> str:
    try:
        n = p.stat().st_size
    except OSError:
        return "-"
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0
    return "-"


class Tooltip:
    """轻量悬浮提示：鼠标悬停在控件上时显示一段说明文字。"""

    def __init__(self, widget, text: str, wraplength: int = 360):
        self.widget = widget
        self.text = text
        self.wraplength = wraplength
        self._win: "tk.Toplevel | None" = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)
        widget.bind("<ButtonPress>", self._hide)

    def _show(self, _event=None) -> None:
        if self._win or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        except Exception:
            return
        win = tk.Toplevel(self.widget)
        win.wm_overrideredirect(True)
        win.wm_geometry(f"+{x}+{y}")
        fam, size = ui_font()
        lbl = ttk.Label(
            win, text=self.text, justify="left",
            background="#ffffe0", foreground="#222222",
            relief="solid", borderwidth=1,
            font=(fam, max(size - 1, 8)),
            padding=(7, 5), wraplength=self.wraplength,
        )
        lbl.pack()
        try:
            win.lift()
        except Exception:
            pass
        self._win = win

    def _hide(self, _event=None) -> None:
        if self._win is not None:
            try:
                self._win.destroy()
            except Exception:
                pass
            self._win = None

    def set_text(self, text: str) -> None:
        self.text = text


# --------------------------------------------------------------------------
# 主窗口
# --------------------------------------------------------------------------


class MaskApp:
    def __init__(self, root: Tk):
        self.root = root
        self.cfg = settings.load()

        self.files: list[Path] = []
        self.results: dict[str, FileResult] = {}
        self.engine: MaskEngine | None = None
        self.tool: ToolInfo | None = None
        self.worker: threading.Thread | None = None
        self.msg_q: queue.Queue = queue.Queue()
        self.running = False
        self.last_output_dir: Path | None = None

        self._build_window()
        self._build_widgets()
        self._bind_events()

        self.root.after(60, self._pump)
        self.detect_engine()

    # -- 窗口 -------------------------------------------------------------

    def _build_window(self) -> None:
        self.root.title(f"{APP_TITLE}  v{APP_VERSION}")
        self.root.geometry("740x660")
        self.root.minsize(680, 600)

        style = ttk.Style()
        for theme in ("vista", "winnative", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break

        fam, size = ui_font()
        self.root.option_add("*Font", (fam, size))
        style.configure("Treeview", rowheight=24, font=(fam, size))
        style.configure("Treeview.Heading", font=(fam, size, "bold"))
        style.configure("Hint.TLabel", foreground="#666666")
        style.configure("Ok.TLabel", foreground="#1a7f37")
        style.configure("Err.TLabel", foreground="#c1121f")
        style.configure("Run.TButton", font=(fam, size, "bold"))

        # 居中
        self.root.update_idletasks()
        w, h = 740, 660
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 3
        self.root.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")

    def _build_widgets(self) -> None:
        pad = {"padx": 10, "pady": 6}
        outer = ttk.Frame(self.root, padding=(12, 10, 12, 10))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        # ---- 引擎状态条 ----
        bar = ttk.Frame(outer)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        bar.columnconfigure(1, weight=1)

        ttk.Label(bar, text="脱敏核心：").grid(row=0, column=0, sticky="w")
        self.lbl_engine = ttk.Label(bar, text="正在检测…", style="Hint.TLabel")
        self.lbl_engine.grid(row=0, column=1, sticky="w")
        self.btn_detect = ttk.Button(bar, text="重新检测", width=10,
                                     command=self.detect_engine)
        self.btn_detect.grid(row=0, column=2, padx=(6, 0))
        self.btn_locate = ttk.Button(bar, text="手动指定…", width=11,
                                     command=self.pick_tool_path)
        self.btn_locate.grid(row=0, column=3, padx=(6, 0))

        # ---- 文件列表 ----
        box = ttk.LabelFrame(outer, text=" 待处理文件 ", padding=(8, 6))
        box.grid(row=1, column=0, sticky="nsew")
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)

        cols = ("name", "size", "status", "detail")
        self.tree = ttk.Treeview(box, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("name", text="文件名")
        self.tree.heading("size", text="大小")
        self.tree.heading("status", text="状态")
        self.tree.heading("detail", text="说明")
        self.tree.column("name", width=240, anchor="w")
        self.tree.column("size", width=70, anchor="e", stretch=False)
        self.tree.column("status", width=60, anchor="center", stretch=False)
        self.tree.column("detail", width=240, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)

        self.tree.tag_configure("ok", foreground="#1a7f37")
        self.tree.tag_configure("err", foreground="#c1121f")
        self.tree.tag_configure("warn", foreground="#b26a00")
        self.tree.tag_configure("busy", foreground="#0b5cad")

        side = ttk.Frame(box)
        side.grid(row=0, column=2, sticky="ns", padx=(8, 0))
        ttk.Button(side, text="添加文件", width=11, command=self.add_files).pack(pady=2)
        ttk.Button(side, text="添加文件夹", width=11, command=self.add_folder).pack(pady=2)
        ttk.Button(side, text="移除所选", width=11, command=self.remove_selected).pack(pady=2)
        ttk.Button(side, text="清空列表", width=11, command=self.clear_files).pack(pady=2)

        self.lbl_hint = ttk.Label(
            box,
            text="支持 .docx / .pdf / .xlsx / .pptx；可一次添加多个文件批量处理",
            style="Hint.TLabel",
        )
        self.lbl_hint.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # ---- 选项 ----
        opt = ttk.LabelFrame(outer, text=" 处理选项 ", padding=(8, 6))
        opt.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        opt.columnconfigure(1, weight=1)

        ttk.Label(opt, text="脱敏模式：").grid(row=0, column=0, sticky="w")
        self.var_mode = StringVar(value=self.cfg["mode"])
        self.cmb_mode = ttk.Combobox(
            opt, textvariable=self.var_mode, state="readonly",
            values=[MODE_LABELS[m] for m in MODES], width=42,
        )
        self.cmb_mode.set(MODE_LABELS.get(self.cfg["mode"], MODE_LABELS["smart"]))
        self.cmb_mode.grid(row=0, column=1, sticky="w", padx=(0, 8))
        self.cmb_mode.bind("<<ComboboxSelected>>", lambda _e: self._sync_mode_dependent())

        # ---- 检测灵敏度（仅 smart 模式可见/生效） ----
        self.lbl_sens_title = ttk.Label(opt, text="检测灵敏度：")
        self.lbl_sens_title.grid(row=1, column=0, sticky="w", pady=(6, 0))
        sens_row = ttk.Frame(opt)
        sens_row.grid(row=1, column=1, columnspan=2, sticky="w", pady=(6, 0))
        self.var_sens = StringVar(value=self.cfg.get("sensitivity") or SENSITIVITY_DEFAULT)
        self._sens_radios: list[ttk.Radiobutton] = []
        self._sens_tooltips: dict[str, Tooltip] = {}
        for key in SENSITIVITY_KEYS:
            lvl = SENSITIVITY_LEVELS[key]
            rb = ttk.Radiobutton(
                sens_row, text=lvl["label"], value=key,
                variable=self.var_sens, command=self._on_sens_change,
            )
            rb.pack(side="left", padx=(0, 8))
            self._sens_tooltips[key] = Tooltip(rb, lvl["desc"])
            self._sens_radios.append(rb)

        # 当前档位释义（常驻显示，强化悬浮提示）
        self.lbl_sens_desc = ttk.Label(
            opt, text="", style="Hint.TLabel", wraplength=660, justify="left",
        )
        self.lbl_sens_desc.grid(row=2, column=0, columnspan=3, sticky="w", pady=(2, 0))
        # 这三个控件仅在 smart 模式下显示，由 _sync_mode_dependent 控制
        self._sens_widgets = [self.lbl_sens_title, sens_row, self.lbl_sens_desc]

        self.var_mapping = BooleanVar(value=bool(self.cfg["save_mapping"]))
        ttk.Checkbutton(opt, text="导出映射表（用于还原）",
                        variable=self.var_mapping).grid(row=0, column=2, sticky="e")

        ttk.Label(opt, text="输出位置：").grid(row=3, column=0, sticky="w", pady=(6, 0))
        outrow = ttk.Frame(opt)
        outrow.grid(row=3, column=1, columnspan=2, sticky="ew", pady=(6, 0))
        outrow.columnconfigure(1, weight=1)

        self.var_outmode = StringVar(value=self.cfg["output_mode"])
        ttk.Radiobutton(outrow, text="原文件所在目录", value="source",
                        variable=self.var_outmode,
                        command=self._sync_out_state).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(outrow, text="指定目录", value="custom",
                        variable=self.var_outmode,
                        command=self._sync_out_state).grid(row=1, column=0, sticky="w")

        self.var_outdir = StringVar(value=self.cfg["custom_output"])
        self.ent_outdir = ttk.Entry(outrow, textvariable=self.var_outdir)
        self.ent_outdir.grid(row=1, column=1, sticky="ew", padx=(8, 6))
        self.btn_outdir = ttk.Button(outrow, text="浏览…", width=8,
                                     command=self.pick_outdir)
        self.btn_outdir.grid(row=1, column=2)

        # ---- 用户词库 ----
        self._build_lexicon_widgets(outer)

        # ---- 进度与操作 ----
        foot = ttk.Frame(outer)
        foot.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        foot.columnconfigure(0, weight=1)
        foot.rowconfigure(1, minsize=34)

        self.pb = ttk.Progressbar(foot, mode="determinate", maximum=100)
        self.pb.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.lbl_status = ttk.Label(
            foot, text="就绪，请添加文件", style="Hint.TLabel",
            anchor="w",
        )
        self.lbl_status.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        btns = ttk.Frame(foot)
        btns.grid(row=0, column=1, rowspan=2, sticky="e")
        self.btn_open = ttk.Button(btns, text="打开输出文件夹", width=17,
                                   command=self.open_output, state="disabled")
        self.btn_open.pack(side="left", padx=(0, 6))
        self.btn_cancel = ttk.Button(btns, text="取消", width=9,
                                     command=self.cancel_run, state="disabled")
        self.btn_cancel.pack(side="left", padx=(0, 6))
        self.btn_run = ttk.Button(btns, text="开始脱敏", width=13,
                                  style="Run.TButton", command=self.start_run)
        self.btn_run.pack(side="left")

        self._sync_out_state()
        self._sync_mode_dependent()

    def _bind_events(self) -> None:
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # -- 引擎检测 ---------------------------------------------------------

    def detect_engine(self) -> None:
        self.btn_detect.configure(state="disabled")
        self.btn_run.configure(state="disabled")
        self.lbl_engine.configure(text="正在检测…", style="Hint.TLabel")

        manual = self.cfg.get("mask_tool_path") or ""

        def work():
            try:
                info = locate_mask_tool(manual or None)
                self.msg_q.put(("engine_ok", info))
            except MaskToolNotFound as exc:
                self.msg_q.put(("engine_fail", str(exc)))
            except Exception:
                self.msg_q.put(("engine_fail", traceback.format_exc(limit=2)))

        threading.Thread(target=work, daemon=True).start()

    def pick_tool_path(self) -> None:
        types = [("可执行文件", "*.exe"), ("所有文件", "*.*")] if IS_WINDOWS \
            else [("所有文件", "*.*")]
        p = filedialog.askopenfilename(title="选择 mask-tool 可执行文件", filetypes=types)
        if p:
            self.cfg["mask_tool_path"] = p
            settings.save(self.cfg)
            self.detect_engine()

    # -- 文件管理 ---------------------------------------------------------

    def _initial_dir(self) -> str:
        d = self.cfg.get("last_dir") or ""
        return d if d and Path(d).is_dir() else str(Path.home())

    def add_files(self) -> None:
        pattern = " ".join(f"*{e}" for e in SUPPORTED_EXTS)
        paths = filedialog.askopenfilenames(
            title="选择要脱敏的文档",
            initialdir=self._initial_dir(),
            filetypes=[
                ("支持的文档", pattern),
                ("Word 文档", "*.docx"),
                ("PDF 文档", "*.pdf"),
                ("Excel 工作簿", "*.xlsx"),
                ("PowerPoint 演示文稿", "*.pptx"),
                ("所有文件", "*.*"),
            ],
        )
        if paths:
            self.cfg["last_dir"] = str(Path(paths[0]).parent)
            self._add([Path(p) for p in paths])

    def add_folder(self) -> None:
        d = filedialog.askdirectory(title="选择文件夹（含子目录）",
                                    initialdir=self._initial_dir())
        if not d:
            return
        self.cfg["last_dir"] = d
        found: list[Path] = []
        try:
            for p in sorted(Path(d).rglob("*")):
                if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS \
                        and not p.name.startswith("~$"):
                    found.append(p)
        except OSError as exc:
            messagebox.showerror("读取失败", f"无法读取该文件夹：\n{exc}")
            return
        if not found:
            messagebox.showinfo("没有找到文件",
                                "该文件夹（含子目录）中没有可处理的文档。\n"
                                "支持的格式：.docx / .pdf / .xlsx / .pptx")
            return
        self._add(found)

    def _add(self, paths: list[Path]) -> None:
        exist = {str(p) for p in self.files}
        skipped = 0
        added = 0
        for p in paths:
            if p.suffix.lower() not in SUPPORTED_EXTS:
                skipped += 1
                continue
            if str(p) in exist:
                continue
            self.files.append(p)
            exist.add(str(p))
            added += 1
            tag = "warn" if p.suffix.lower() in REPORT_ONLY_EXTS else ""
            note = "PDF 仅输出检测报告" if p.suffix.lower() in REPORT_ONLY_EXTS else ""
            self.tree.insert("", "end", iid=str(p),
                             values=(p.name, human_size(p), "待处理", note),
                             tags=(tag,) if tag else ())
        self._refresh_counter()
        if skipped:
            messagebox.showwarning(
                "部分文件已跳过",
                f"有 {skipped} 个文件格式不受支持，已自动跳过。\n"
                "仅支持 .docx / .pdf / .xlsx / .pptx",
            )
        if added:
            self.lbl_status.configure(text=f"已添加 {added} 个文件")

    def remove_selected(self) -> None:
        if self.running:
            return
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self.files = [f for f in self.files if str(f) != iid]
            self.results.pop(iid, None)
        self._refresh_counter()

    def clear_files(self) -> None:
        if self.running:
            return
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self.results.clear()
        self.pb["value"] = 0
        self.lbl_status.configure(text="就绪", style="Hint.TLabel")
        self._refresh_counter()

    def _refresh_counter(self) -> None:
        n = len(self.files)
        parent = self.tree.master
        if isinstance(parent, ttk.LabelFrame):
            parent.configure(text=f" 待处理文件（{n}） " if n else " 待处理文件 ")

    # -- 选项 -------------------------------------------------------------

    def _sync_out_state(self) -> None:
        custom = self.var_outmode.get() == "custom"
        state = "normal" if custom else "disabled"
        self.ent_outdir.configure(state=state)
        self.btn_outdir.configure(state=state)

    def _on_sens_change(self) -> None:
        key = self.var_sens.get()
        self.cfg["sensitivity"] = key
        settings.save(self.cfg)
        self._sync_mode_dependent()

    def _sync_mode_dependent(self) -> None:
        """根据当前模式联动灵敏度控件的可见性与说明。

        检测灵敏度仅对 smart 模式有意义：strict 只用用户词库、aggressive
        已是最高召回，因此这两种模式下隐藏灵敏度设置，避免误导。
        """
        smart = self._current_mode() == "smart"
        if smart:
            for w in self._sens_widgets:
                w.grid()
            key = self.var_sens.get()
            lvl = SENSITIVITY_LEVELS.get(key, SENSITIVITY_LEVELS[SENSITIVITY_DEFAULT])
            self.lbl_sens_desc.configure(text=lvl["desc"])
        else:
            for w in self._sens_widgets:
                w.grid_remove()

    # -- 用户词库 ---------------------------------------------------------

    def _build_lexicon_widgets(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(
            parent, text=" 用户词库（strict 模式主要依赖此词典） ", padding=(8, 6)
        )
        frame.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        frame.columnconfigure(2, weight=1)

        self.var_lex_cat = StringVar(value=LEX_CATEGORIES[0][0])
        ttk.Label(frame, text="类别：").grid(row=0, column=0, sticky="w")
        self.cmb_lex_cat = ttk.Combobox(
            frame, textvariable=self.var_lex_cat, state="readonly",
            values=[c[0] for c in LEX_CATEGORIES], width=12,
        )
        self.cmb_lex_cat.grid(row=0, column=1, sticky="w", padx=(0, 8))

        self.PLACEHOLDER_LEX = "输入词，逗号分隔，回车添加"
        self.ent_lex = ttk.Entry(frame, width=26)
        self.ent_lex.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        self.ent_lex.insert(0, self.PLACEHOLDER_LEX)
        self.ent_lex.configure(foreground="#999999")
        self.ent_lex.bind("<FocusIn>", self._lex_entry_focus_in)
        self.ent_lex.bind("<FocusOut>", self._lex_entry_focus_out)
        self.ent_lex.bind("<Return>", lambda e: self._add_lex_word())

        ttk.Button(frame, text="添加", width=8,
                   command=self._add_lex_word).grid(row=0, column=3, padx=(0, 6))
        ttk.Button(frame, text="导入TXT", width=9,
                   command=self._import_lex_txt).grid(row=0, column=4, padx=(0, 6))
        ttk.Button(frame, text="清空", width=8,
                   command=self._clear_lex).grid(row=0, column=5)

        list_row = ttk.Frame(frame)
        list_row.grid(row=1, column=0, columnspan=6, sticky="ew", pady=(6, 0))
        list_row.columnconfigure(0, weight=1)

        self.lst_lex = Listbox(list_row, height=3, selectmode="extended",
                               exportselection=False)
        self.lst_lex.grid(row=0, column=0, sticky="nsew")
        sb_lex = ttk.Scrollbar(list_row, orient="vertical",
                               command=self.lst_lex.yview)
        sb_lex.grid(row=0, column=1, sticky="ns")
        self.lst_lex.configure(yscrollcommand=sb_lex.set)
        self.lst_lex.bind("<Delete>", lambda e: self._remove_lex_selected())

        self.lbl_lex_count = ttk.Label(list_row, text="已收录 0 个词",
                                       style="Hint.TLabel")
        self.lbl_lex_count.grid(row=1, column=0, sticky="w", pady=(2, 0))
        ttk.Button(list_row, text="删除所选", width=10,
                   command=self._remove_lex_selected).grid(
            row=1, column=1, sticky="e", padx=(6, 0))

        self._refresh_lex_list()

    def _lex_cat_key(self) -> str:
        label = self.var_lex_cat.get()
        for lbl, key in LEX_CATEGORIES:
            if lbl == label:
                return key
        return "custom"

    def _lex_entry_focus_in(self, _e=None) -> None:
        if self.ent_lex.get() == self.PLACEHOLDER_LEX:
            self.ent_lex.delete(0, "end")
            self.ent_lex.configure(foreground="#000000")

    def _lex_entry_focus_out(self, _e=None) -> None:
        if not self.ent_lex.get().strip():
            self.ent_lex.insert(0, self.PLACEHOLDER_LEX)
            self.ent_lex.configure(foreground="#999999")

    @staticmethod
    def _split_words(text: str) -> list[str]:
        import re
        parts = re.split(r"[\n,，、;；\t]+", text)
        return [w.strip() for w in parts if w.strip()]

    def _add_lex_word(self) -> None:
        text = self.ent_lex.get().strip()
        if not text or text == self.PLACEHOLDER_LEX:
            return
        cat = self._lex_cat_key()
        words = self._split_words(text)
        if not words:
            return
        current = self.cfg.setdefault("user_lexicon", {})
        existing = list(current.get(cat, []))
        added = 0
        for w in words:
            if w not in existing:
                existing.append(w)
                added += 1
        if added:
            current[cat] = existing
            settings.save(self.cfg)
            self._refresh_lex_list()
        self.ent_lex.delete(0, "end")
        self._lex_entry_focus_out()

    def _import_lex_txt(self) -> None:
        p = filedialog.askopenfilename(
            title="导入词库（TXT，每行一个词，可带类别前缀）",
            initialdir=self._initial_dir(),
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if not p:
            return
        try:
            data = Path(p).read_text(encoding="utf-8-sig", errors="ignore")
        except OSError as exc:
            messagebox.showerror("读取失败", f"无法读取文件：\n{exc}")
            return
        parsed = self._parse_lex_txt(data)
        if not parsed:
            messagebox.showinfo("提示", "文件中没有可导入的词。")
            return
        current = self.cfg.setdefault("user_lexicon", {})
        total_added = 0
        cat_labels = {k: lbl for lbl, k in LEX_CATEGORIES}
        summary = []
        for cat, words in parsed.items():
            existing = list(current.get(cat, []))
            before = len(existing)
            for w in words:
                if w not in existing:
                    existing.append(w)
            added = len(existing) - before
            if added:
                current[cat] = existing
                total_added += added
                summary.append(f"「{cat_labels.get(cat, cat)}」+{added}")
        if total_added:
            settings.save(self.cfg)
            self._refresh_lex_list()
        messagebox.showinfo(
            "导入完成",
            f"已导入 {total_added} 个新词：\n" + "；".join(summary),
        )

    def _parse_lex_txt(self, text: str) -> dict[str, list[str]]:
        """解析 TXT 词库文件，返回 {类别 key: [词...]}。

        规则：
          · 以 ``#`` 或 ``//`` 开头的行是注释，忽略；
          · 行内可带类别前缀（公司名 / 人名 / 项目名 / 机构 / 单位 /
            地名 / 金额 / 自定义，或英文 company / person / ...），
            前缀后接中文冒号 ``：`` 或英文冒号 ``:``，再写词；
          · 没有前缀的行，归入界面下拉框「当前选中的类别」；
          · 同一行可用 逗号、顿号、分号、制表符 分隔多个词，
            也可以一行一个词。
        """
        label_to_key = {lbl: k for lbl, k in LEX_CATEGORIES}
        default_cat = self._lex_cat_key()
        result: dict[str, list[str]] = {}
        for raw in text.splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or line.startswith("//"):
                continue
            m = re.match(r"^([\u4e00-\u9fffA-Za-z_/ ]+?)[：:]\s*(.*)$", line)
            if m:
                prefix = m.group(1).strip()
                rest = m.group(2).strip()
                cat = label_to_key.get(prefix) or (
                    prefix.lower() if prefix.lower() in label_to_key.values() else None
                )
                if cat is None:
                    # 前缀不匹配任何类别，整行作为词归入默认类别
                    words = self._split_words(line)
                    cat = default_cat
                else:
                    words = self._split_words(rest) if rest else []
            else:
                cat = default_cat
                words = self._split_words(line)
            if words:
                result.setdefault(cat, []).extend(words)
        return result

    def _remove_lex_selected(self) -> None:
        sel = self.lst_lex.curselection()
        if not sel:
            return
        to_remove: list[tuple[str, str]] = []
        label_to_key = {lbl: k for lbl, k in LEX_CATEGORIES}
        for i in sel:
            item = self.lst_lex.get(i)
            if ":" in item:
                cat_lbl, word = item.split(":", 1)
                key = label_to_key.get(cat_lbl.strip())
                if key:
                    to_remove.append((key, word.strip()))
        if not to_remove:
            return
        current = self.cfg.get("user_lexicon", {})
        for key, word in to_remove:
            if key in current and word in current[key]:
                current[key].remove(word)
                if not current[key]:
                    del current[key]
        settings.save(self.cfg)
        self._refresh_lex_list()

    def _clear_lex(self) -> None:
        if not self.cfg.get("user_lexicon"):
            return
        if not messagebox.askyesno("确认清空", "确定要清空全部用户词库吗？"):
            return
        self.cfg["user_lexicon"] = {}
        settings.save(self.cfg)
        self._refresh_lex_list()

    def _refresh_lex_list(self) -> None:
        self.lst_lex.delete(0, "end")
        total = 0
        lex = self.cfg.get("user_lexicon", {})
        label_of = {k: lbl for lbl, k in LEX_CATEGORIES}
        for key, words in lex.items():
            cat_lbl = label_of.get(key, key)
            for w in words:
                self.lst_lex.insert("end", f"{cat_lbl}: {w}")
                total += 1
        self.lbl_lex_count.configure(text=f"已收录 {total} 个词")

    def pick_outdir(self) -> None:
        d = filedialog.askdirectory(title="选择输出目录",
                                    initialdir=self.var_outdir.get() or self._initial_dir())
        if d:
            self.var_outdir.set(d)

    def _current_mode(self) -> str:
        label = self.cmb_mode.get()
        for m, text in MODE_LABELS.items():
            if text == label:
                return m
        return "smart"

    # -- 运行 -------------------------------------------------------------

    def _validate(self) -> tuple[bool, str]:
        if self.engine is None:
            return False, "脱敏核心尚未就绪，请先完成 mask-tool 的安装与检测。"
        if not self.files:
            return False, "请先添加需要脱敏的文档。"
        missing = [f.name for f in self.files if not f.is_file()]
        if missing:
            return False, "以下文件已不存在，请重新添加：\n" + "\n".join(missing[:8])
        if self.var_outmode.get() == "custom":
            d = self.var_outdir.get().strip()
            if not d:
                return False, "已选择“指定目录”，请先浏览选择输出位置。"
            try:
                Path(d).mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                return False, f"输出目录不可用：\n{exc.strerror or exc}"
        return True, ""

    def start_run(self) -> None:
        ok, err = self._validate()
        if not ok:
            messagebox.showwarning("无法开始", err)
            return

        mode = self._current_mode()
        save_mapping = bool(self.var_mapping.get())
        custom = self.var_outmode.get() == "custom"
        outdir = Path(self.var_outdir.get().strip()) if custom else None
        sensitivity = self.var_sens.get() or SENSITIVITY_DEFAULT

        self.cfg.update({
            "mode": mode,
            "sensitivity": sensitivity,
            "output_mode": self.var_outmode.get(),
            "custom_output": self.var_outdir.get().strip(),
            "save_mapping": save_mapping,
        })
        settings.save(self.cfg)

        # 重置列表状态
        self.results.clear()
        for f in self.files:
            self.tree.item(str(f), values=(f.name, human_size(f), "排队中", ""), tags=())

        self._set_running(True)
        self.pb.configure(maximum=len(self.files), value=0)

        if self.tool is None:
            messagebox.showerror("缺少脱敏核心", INSTALL_HELP)
            return

        # 用当前用户词库重建引擎（词库改动实时生效）
        lex = self.cfg.get("user_lexicon") or {}
        engine = MaskEngine(self.tool, user_lexicon=lex)
        self.engine = engine
        files = list(self.files)
        suffix = self.cfg.get("suffix_tag") or "_脱敏"

        def out_dir_for(src: Path) -> Path:
            return outdir if outdir else src.parent

        def work():
            try:
                engine.process_batch(
                    files,
                    out_dir_for=out_dir_for,
                    mode=mode,
                    save_mapping=save_mapping,
                    suffix_tag=suffix,
                    sensitivity=sensitivity,
                    on_progress=lambda i, n, p: self.msg_q.put(("progress", (i, n, p))),
                    on_result=lambda r: self.msg_q.put(("result", r)),
                )
                self.msg_q.put(("done", outdir or (files[0].parent if files else None)))
            except Exception:
                self.msg_q.put(("crash", traceback.format_exc()))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def cancel_run(self) -> None:
        if self.engine and self.running:
            self.engine.cancel()
            self.lbl_status.configure(text="正在取消，请等待当前文件结束…")
            self.btn_cancel.configure(state="disabled")

    def _set_running(self, running: bool) -> None:
        self.running = running
        state = "disabled" if running else "normal"
        for w in (self.btn_run, self.btn_detect, self.btn_locate, self.cmb_mode):
            w.configure(state="disabled" if running else
                        ("readonly" if w is self.cmb_mode else "normal"))
        self.btn_cancel.configure(state="normal" if running else "disabled")
        if not running:
            self.btn_run.configure(state="normal" if self.engine else "disabled")
        _ = state

    # -- 消息泵 -----------------------------------------------------------

    def _pump(self) -> None:
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                self._handle(kind, payload)
        except queue.Empty:
            pass
        self.root.after(60, self._pump)

    def _handle(self, kind: str, payload) -> None:
        if kind == "engine_ok":
            self.tool = payload
            self.engine = MaskEngine(payload, user_lexicon=self.cfg.get("user_lexicon") or {})
            ver = f"  {payload.version}" if payload.version else ""
            self.lbl_engine.configure(text=f"✓ 已就绪{ver}", style="Ok.TLabel")
            self.btn_detect.configure(state="normal")
            self.btn_run.configure(state="normal")
            self.lbl_status.configure(text="就绪，请添加文件")

        elif kind == "engine_fail":
            self.tool = None
            self.engine = None
            self.lbl_engine.configure(text="✕ 未检测到 mask-tool", style="Err.TLabel")
            self.btn_detect.configure(state="normal")
            self.btn_run.configure(state="disabled")
            self.lbl_status.configure(text="请先安装脱敏核心")
            messagebox.showerror("缺少脱敏核心", INSTALL_HELP)

        elif kind == "progress":
            i, n, p = payload
            self.pb.configure(value=i - 1)
            self.lbl_status.configure(text=f"正在处理 {i}/{n}：{p.name}")
            if self.tree.exists(str(p)):
                self.tree.item(str(p),
                               values=(p.name, human_size(p), "处理中", "…"),
                               tags=("busy",))
                self.tree.see(str(p))

        elif kind == "result":
            r: FileResult = payload
            iid = str(r.source)
            tag = "ok" if r.ok and not r.report_only else ("warn" if r.ok else "err")
            if self.tree.exists(iid):
                self.tree.item(iid, values=(r.source.name, human_size(r.source),
                                            r.status_text, r.message), tags=(tag,))
            self.results[iid] = r
            self.pb.configure(value=self.pb["value"] + 1)
            if r.ok and r.output:
                self.last_output_dir = r.output.parent

        elif kind == "done":
            self._finish(payload)

        elif kind == "crash":
            self._set_running(False)
            self.lbl_status.configure(text="处理异常中止")
            messagebox.showerror(
                "程序内部错误",
                "处理过程中出现未预期的问题，本次任务已中止。\n"
                "文档未被修改，可重试或减少一次处理的文件数量。",
            )

    def _finish(self, outdir: Path | None) -> None:
        self._set_running(False)
        self.pb.configure(value=self.pb["maximum"])

        total = len(self.results)
        ok = sum(1 for r in self.results.values() if r.ok)
        fail = total - ok
        report_only = sum(1 for r in self.results.values() if r.ok and r.report_only)
        cancelled = total < len(self.files)

        if outdir and Path(outdir).is_dir():
            self.last_output_dir = Path(outdir)
        self.btn_open.configure(state="normal" if self.last_output_dir else "disabled")

        head = "处理已取消" if cancelled else "处理完成"
        self.lbl_status.configure(
            text=f"{head}：成功 {ok} 个，失败 {fail} 个（共 {total} 个）",
            style="Err.TLabel" if fail else "Ok.TLabel",
        )

        lines = [f"共处理 {total} 个文件", f"  成功：{ok} 个"]
        if report_only:
            lines.append(f"  其中 PDF 仅生成检测报告：{report_only} 个")
        if fail:
            lines.append(f"  失败：{fail} 个（详见列表“说明”列）")
        if cancelled:
            lines.append("\n任务被手动取消，剩余文件未处理。")
        if self.last_output_dir:
            lines.append(f"\n输出位置：\n{self.last_output_dir}")

        body = "\n".join(lines)
        if fail:
            messagebox.showwarning(head, body)
        else:
            messagebox.showinfo(head, body)

    # -- 其他 -------------------------------------------------------------

    def _on_double_click(self, _event) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        r = self.results.get(sel[0])
        if r and r.ok and r.output and r.output.exists():
            open_in_explorer(r.output)
        else:
            p = Path(sel[0])
            if p.exists():
                open_in_explorer(p)

    def open_output(self) -> None:
        if self.last_output_dir and self.last_output_dir.is_dir():
            open_in_explorer(self.last_output_dir)
        else:
            messagebox.showinfo("提示", "暂无可打开的输出目录。")

    def on_close(self) -> None:
        if self.running:
            if not messagebox.askyesno("确认退出", "任务正在进行中，确定要退出吗？"):
                return
            if self.engine:
                self.engine.cancel()
        settings.save(self.cfg)
        self.root.destroy()


# --------------------------------------------------------------------------


def run() -> None:
    enable_dpi_awareness()
    root = Tk()

    def on_error(exc_type, exc_value, exc_tb):
        """兜底：任何未捕获异常都转成友好中文弹窗，绝不显示 Python 堆栈。"""
        try:
            messagebox.showerror(
                "程序出现问题",
                "很抱歉，程序遇到了一个未预期的问题。\n\n"
                "您的文档没有被修改。请重试；若反复出现，"
                "请尝试重启程序或减少一次处理的文件数量。",
            )
        except Exception:
            pass

    root.report_callback_exception = on_error  # type: ignore[assignment]
    MaskApp(root)
    root.mainloop()
