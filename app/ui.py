"""
Tkinter 图形界面（小清新风格）
===========================

选用 Tkinter + ttk 而非 PyQt6，理由：
  · ttk 在 Windows 上观感稳定，且可通过 clam 主题做轻量自定义；
  · 标准库自带，PyInstaller 单文件产物体积小、冷启动快，符合“双击即用”。

本文件在保留全部功能逻辑（脱敏、选项、词库、标记漏报、白名单、strict 模式等）
的前提下，对视觉层做了“小清新”全面升级：
  · 自定义标题栏（非系统默认）、浅蓝/薄荷绿背景、柔和蓝/浅珊瑚点缀、深灰文本；
  · 圆角卡片分区、圆角按钮、悬停/点击动效、薄荷绿进度条、彩色状态标签；
  · 列表上方虚线“拖拽上传”区（点击亦可选择文件）、行交替底色与悬停高亮；
  · 圆角错误/提示弹窗（带“我知道了”按钮）。

所有耗时操作均在工作线程执行，通过 queue 回主线程刷新，保证不卡顿。
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
import tkinter.font as tkfont
from tkinter import (
    BooleanVar, Listbox, StringVar, Tk, filedialog, messagebox, ttk,
)

from app import settings
from app.engine import (
    MODE_LABELS, MODES, REPORT_ONLY_EXTS, SUPPORTED_EXTS,
    SENSITIVITY_LEVELS, SENSITIVITY_KEYS, SENSITIVITY_DEFAULT,
    FileResult, MaskEngine, MaskToolNotFound, ToolInfo, locate_mask_tool,
    load_whitelist, save_whitelist,
    set_min_confidence, set_ner_backend, ner_status,
)

APP_TITLE = "本地文档脱敏工具"
APP_VERSION = "1.0.0"
APP_NAME = "📎 文档脱敏助手"

IS_WINDOWS = os.name == "nt"

# --------------------------------------------------------------------------
# 小清新调色板
# --------------------------------------------------------------------------
BG_BLUE = "#E8F4F8"        # 浅蓝背景
BG_MINT = "#E0F2E6"        # 薄荷绿背景
ACCENT_BLUE = "#7EB8D0"    # 柔和蓝（主按钮 / 标题栏）
ACCENT_BLUE_DK = "#5C9CB8" # 柔和蓝加深（悬停 / 标题栏激活）
ACCENT_CORAL = "#F7C6B7"   # 浅珊瑚（强调按钮）
CORAL_DK = "#EBA98F"      # 浅珊瑚加深
TEXT = "#333333"           # 深灰文本（避免纯黑）
TEXT_SOFT = "#6B7B83"      # 次要文本
BORDER = "#CFE3EC"         # 浅色边框
CARD_BG = "#FFFFFF"        # 卡片底色
MINT_BAR = "#86D3A8"       # 进度条薄荷绿
HOVER_ROW = "#DCEFF6"      # 列表悬停行高亮
ROW_ALT = "#F4F8FB"        # 列表交替行底色

# 字体
_FAM0, _SZ0 = (lambda: ("Microsoft YaHei UI", 9))() if False else (None, None)


def ui_font() -> tuple[str, int]:
    if IS_WINDOWS:
        return ("Microsoft YaHei UI", 10)
    if sys.platform == "darwin":
        return ("PingFang SC", 13)
    return ("Noto Sans CJK SC", 10)


_FAM, _SZ = ui_font()
FAM = _FAM
# 正文字号：Windows 至少 10pt，其余沿用系统推荐
FSIZE = _SZ if _SZ >= 10 else 10
TITLE_FONT = (FAM, 14, "bold")
HEAD_FONT = (FAM, 11, "bold")
SUB_FONT = (FAM, FSIZE, "bold")
BODY_FONT = (FAM, FSIZE)


def _darken(hex_color: str, factor: float = 0.10) -> str:
    """将十六进制颜色按 factor 比例加深，用于 hover / 描边。"""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    r = int(r * (1 - factor))
    g = int(g * (1 - factor))
    b = int(b * (1 - factor))
    return f"#{r:02x}{g:02x}{b:02x}"


def rounded_rect(c, x1: float, y1: float, x2: float, y2: float, r: float, **kw):
    """在 Canvas 上绘制圆角矩形（平滑多边形近似），返回图形 id。"""
    pts = (
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    )
    return c.create_polygon(pts, smooth=True, **kw)


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


# --------------------------------------------------------------------------
# 圆角卡片容器
# --------------------------------------------------------------------------

class Card(tk.Frame):
    """浅色圆角卡片：在窗口底色上绘制白色圆角矩形作为可视卡片。"""

    def __init__(self, master, title: str | None = None, radius: int = 12,
                 padx: int = 16, pady: int = 14, **kw):
        super().__init__(master, bg=master.cget("bg"), bd=0,
                         highlightthickness=0, **kw)
        self.radius = radius
        self._cv = tk.Canvas(self, bg=master.cget("bg"), highlightthickness=0)
        self._cv.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._padx = padx
        self._pady = pady
        self._title_label: tk.Label | None = None
        self.bind("<Configure>", lambda e: self._draw())
        if title:
            self.set_title(title)
        self._draw()

    def _draw(self) -> None:
        w = self.winfo_width()
        h = self.winfo_height()
        self._cv.delete("all")
        if w > 4 and h > 4:
            rounded_rect(self._cv, 1, 1, w - 1, h - 1, self.radius,
                         fill=CARD_BG, outline=BORDER, width=1)

    def set_title(self, text: str) -> None:
        if self._title_label is None:
            self._title_label = tk.Label(
                self, text=text, font=HEAD_FONT, bg=CARD_BG, fg=ACCENT_BLUE_DK)
            self._title_label.place(x=self._padx, y=self._pady - 4)
        else:
            self._title_label.configure(text=text)

    def body(self) -> tk.Frame:
        """返回卡片内部的内容区（已内缩，避免压到圆角）。"""
        f = tk.Frame(self, bg=CARD_BG)
        top = self._pady + (22 if self._title_label is not None else 0)
        f.pack(fill="both", expand=True, padx=self._padx, pady=(top, self._pady))
        return f


# --------------------------------------------------------------------------
# 圆角按钮（支持图标 + 文字、悬停加深、点击缩小）
# --------------------------------------------------------------------------

class ModernButton(tk.Frame):
    def __init__(self, master, label: str = "", command=None,
                 variant: str = "primary", width: int | None = None,
                 state: str = "normal", height: int = 34, padx: int = 16,
                 radius: int = 8):
        super().__init__(master, bg=master.cget("bg"), bd=0, highlightthickness=0)
        self._label = label
        self._command = command
        self._variant = variant
        self._state = state
        self._height = height
        self._padx = padx
        self._width = width
        self._radius = radius
        self._hover = False
        self._scale = 1.0
        self._cv = tk.Canvas(self, height=height, bg=master.cget("bg"),
                             highlightthickness=0)
        self._cv.pack(fill="both", expand=True)
        self._measure()
        self._bind()
        self._draw()

    # -- 尺寸 / 颜色 -------------------------------------------------------
    def _measure(self) -> None:
        f = tkfont.Font(self, font=BODY_FONT)
        w = f.measure(self._label) if self._label else 0
        w += self._padx * 2
        if self._width:
            w = max(w, self._width)
        self._btn_w = int(max(w, 44))
        self._cv.configure(width=self._btn_w, height=self._height)

    def _base_colors(self):
        if self._variant == "primary":
            return ACCENT_BLUE, "#FFFFFF"
        if self._variant == "coral":
            return ACCENT_CORAL, TEXT
        if self._variant == "mint":
            return MINT_BAR, TEXT
        return "#FFFFFF", TEXT  # neutral

    def _draw(self) -> None:
        self._cv.delete("all")
        base, tcol = self._base_colors()
        if self._state == "disabled":
            base, tcol, edge = "#E4E9EC", "#AAB4B9", "#D6DDE1"
        else:
            edge = _darken(base, 0.15)
            if self._hover:
                base = _darken(base, 0.10)
        r = self._radius
        w = self._btn_w * self._scale
        h = self._height * self._scale
        x0 = (self._btn_w - w) / 2
        y0 = (self._height - h) / 2
        rounded_rect(self._cv, x0, y0, x0 + w, y0 + h, r,
                     fill=base, outline=edge, width=1)
        self._cv.create_text(self._btn_w / 2, self._height / 2, text=self._label,
                             font=BODY_FONT, fill=tcol, anchor="center")

    # -- 交互 ---------------------------------------------------------------
    def _bind(self) -> None:
        self._cv.bind("<Enter>", lambda e: self._enter())
        self._cv.bind("<Leave>", lambda e: self._leave())
        self._cv.bind("<ButtonPress-1>", lambda e: self._press())
        self._cv.bind("<ButtonRelease-1>", lambda e: self._release())

    def _enter(self) -> None:
        if self._state == "disabled":
            return
        self._hover = True
        self._draw()

    def _leave(self) -> None:
        self._hover = False
        self._scale = 1.0
        self._draw()

    def _press(self) -> None:
        if self._state == "disabled":
            return
        self._scale = 0.96
        self._draw()

    def _release(self) -> None:
        was = self._state
        self._scale = 1.0
        self._draw()
        if was == "disabled":
            return
        if self._command:
            self._command()

    # -- 兼容 ttk 的 configure / cget --------------------------------------
    def configure(self, **kw):
        if "state" in kw:
            self._state = kw["state"]
            self._draw()
        if "text" in kw:
            self._label = kw["text"]
            self._measure()
            self._draw()
        if "label" in kw:
            self._label = kw["label"]
            self._measure()
            self._draw()
        try:
            super().configure(
                **{k: v for k, v in kw.items()
                   if k not in ("state", "text", "label")})
        except Exception:
            pass

    config = configure

    def cget(self, key):
        if key == "state":
            return self._state
        return super().cget(key)


# --------------------------------------------------------------------------
# 轻量悬浮提示（圆角气泡）
# --------------------------------------------------------------------------

class Tooltip:
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
            x = self.widget.winfo_rootx() + 14
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except Exception:
            return
        win = tk.Toplevel(self.widget)
        win.wm_overrideredirect(True)
        win.wm_geometry(f"+{x}+{y}")
        win.lift()
        f = tkfont.Font(win, font=BODY_FONT)
        lines: list[str] = []
        cur = ""
        for ch in self.text:
            if ch == "\n":
                lines.append(cur)
                cur = ""
                continue
            if f.measure(cur + ch) > self.wraplength and cur:
                lines.append(cur)
                cur = ch
            else:
                cur += ch
        if cur:
            lines.append(cur)
        w = max((f.measure(ln) for ln in lines), default=10) + 20
        h = len(lines) * int(FSIZE * 1.5) + 14
        cv = tk.Canvas(win, width=w, height=h, bg=BG_BLUE, highlightthickness=0)
        rounded_rect(cv, 0, 0, w, h, 8, fill="#FFFDF5", outline=BORDER, width=1)
        for i, ln in enumerate(lines):
            cv.create_text(10, 8 + i * int(FSIZE * 1.5), text=ln,
                           anchor="nw", font=BODY_FONT, fill=TEXT)
        cv.pack()
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
        self._row_index = 0
        self._hover_iid: str | None = None

        self._build_window()
        self._build_widgets()
        self._bind_events()

        # 所有控件创建并布局完成后再去掉系统边框，避免子控件 pack 时窗口路径失效
        try:
            self.root.overrideredirect(True)
        except Exception:
            pass
        self._set_window_shadow()

        self.root.after(60, self._pump)
        self.detect_engine()

    # -- 窗口 -------------------------------------------------------------

    def _build_window(self) -> None:
        self.root.title(APP_TITLE)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # 全局 ttk 样式（小清新）
        style.configure(".", background=BG_BLUE, foreground=TEXT, font=BODY_FONT)
        style.configure("TLabel", background=BG_BLUE, foreground=TEXT)
        style.configure("TFrame", background=BG_BLUE)
        style.configure("TCombobox",
                        fieldbackground="#FFFFFF", foreground=TEXT,
                        bordercolor=BORDER, lightcolor=BG_BLUE,
                        darkcolor=BG_BLUE, padding=4)
        style.map("TCombobox",
                  fieldbackground=[("readonly", "#FFFFFF")],
                  background=[("readonly", ACCENT_BLUE)])
        style.configure("TEntry",
                        fieldbackground="#FFFFFF", foreground=TEXT,
                        bordercolor=BORDER, insertcolor=TEXT, padding=4)
        style.configure("TCheckbutton", background=CARD_BG, foreground=TEXT,
                        indicatorcolor="#FFFFFF", bordercolor=BORDER)
        style.map("TCheckbutton",
                  indicatorcolor=[("selected", ACCENT_BLUE)],
                  background=[("active", CARD_BG)])
        style.configure("TRadiobutton", background=CARD_BG, foreground=TEXT,
                        indicatorcolor="#FFFFFF", bordercolor=BORDER)
        style.map("TRadiobutton",
                  indicatorcolor=[("selected", ACCENT_BLUE)],
                  background=[("active", CARD_BG)])
        style.configure("Treeview", background=CARD_BG, fieldbackground=CARD_BG,
                        foreground=TEXT, rowheight=30, bordercolor=BORDER,
                        relief="flat")
        style.configure("Treeview.Heading", background=ACCENT_BLUE,
                        foreground="#FFFFFF", relief="flat", font=HEAD_FONT)
        style.map("Treeview.Heading",
                  background=[("active", ACCENT_BLUE_DK)])
        style.configure("Mint.Horizontal.TProgressbar", troughcolor="#E7F2EC",
                        background=MINT_BAR, borderwidth=0, thickness=14)
        style.configure("Hint.TLabel", foreground=TEXT_SOFT, background=BG_BLUE)
        style.configure("Ok.TLabel", foreground="#2E9E5B", background=BG_BLUE)
        style.configure("Err.TLabel", foreground="#D9534F", background=BG_BLUE)
        style.configure("TScrollbar", background=BG_BLUE, troughcolor=BG_BLUE,
                        bordercolor=BORDER, arrowcolor=ACCENT_BLUE_DK)

        self.root.geometry("940x960")
        self.root.minsize(820, 680)

        # 居中
        self.root.update_idletasks()
        w, h = 940, 960
        x = (self.root.winfo_screenwidth() - w) // 2
        y = max((self.root.winfo_screenheight() - h) // 3, 20)
        self.root.geometry(f"{w}x{h}+{max(x, 0)}+{y}")


    def _set_window_shadow(self) -> None:
        """为无边框窗口添加轻微阴影（Windows）。"""
        if not IS_WINDOWS:
            return
        try:
            import ctypes
            GCL_STYLE = -26
            CS_DROPSHADOW = 0x00020000
            self.root.update_idletasks()
            hwnd = self.root.winfo_id()
            if not hwnd:
                return
            cur = ctypes.windll.user32.GetClassLongPtrW(hwnd, GCL_STYLE)
            ctypes.windll.user32.SetClassLongPtrW(
                hwnd, GCL_STYLE, cur | CS_DROPSHADOW)
        except Exception:
            pass

    # -- 自定义标题栏 ------------------------------------------------------

    def _build_title_bar(self) -> None:
        tb = tk.Frame(self.root, bg=ACCENT_BLUE, height=44)
        tb.pack(side="top", fill="x")
        tb.pack_propagate(False)

        lbl = tk.Label(tb, text=APP_NAME, bg=ACCENT_BLUE, fg="#FFFFFF",
                       font=TITLE_FONT)
        lbl.pack(side="left", padx=16)

        close_b = ModernButton(
            tb, label="✕", variant="coral", width=28, height=28,
            radius=14, command=self.on_close)
        close_b.pack(side="right", padx=(0, 10))
        min_b = ModernButton(
            tb, label="—", variant="neutral", width=28, height=28,
            radius=14, command=self.root.iconify)
        min_b.pack(side="right", padx=(0, 4))

        self._drag_widgets = (tb, lbl)
        tb.bind("<ButtonPress-1>", self._start_drag)
        tb.bind("<B1-Motion>", self._do_drag)
        lbl.bind("<ButtonPress-1>", self._start_drag)
        lbl.bind("<B1-Motion>", self._do_drag)

    def _start_drag(self, event) -> None:
        if event.widget not in self._drag_widgets:
            return
        self._drag_x = event.x_root - self.root.winfo_x()
        self._drag_y = event.y_root - self.root.winfo_y()

    def _do_drag(self, event) -> None:
        if event.widget not in self._drag_widgets:
            return
        x = event.x_root - self._drag_x
        y = event.y_root - self._drag_y
        self.root.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    # -- 主内容（可滚动） --------------------------------------------------

    def _build_scroll(self) -> None:
        self._sb = tk.Scrollbar(self.root, orient="vertical", width=10,
                                bg=BG_BLUE, troughcolor=BG_BLUE,
                                bd=0, relief="flat")
        self._canvas = tk.Canvas(self.root, bg=BG_BLUE, highlightthickness=0,
                                 yscrollcommand=self._sb.set)
        self._sb.config(command=self._canvas.yview)
        self._canvas.pack(side="left", fill="both", expand=True,
                          padx=10, pady=6)
        self._sb.pack(side="right", fill="y")

        self.inner = tk.Frame(self._canvas, bg=BG_BLUE)
        self._canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda e: self._canvas.configure(
                            scrollregion=self._canvas.bbox("all")))
        self.inner.bind("<MouseWheel>",
                        lambda e: self._canvas.yview_scroll(
                            -int(e.delta / 120), "units"))

    # -- 各功能卡片 -------------------------------------------------------

    def _build_widgets(self) -> None:
        self._build_title_bar()
        self._build_scroll()

        # 引擎状态条（轻量卡片）
        self._build_engine_bar()

        # 待处理文件卡片
        self._build_file_card()

        # 处理选项卡片
        self._build_options_card()

        # 识别引擎（高级）
        self._build_engine_widgets(self.inner)

        # 用户词库
        self._build_lexicon_widgets(self.inner)

        # 底部固定操作栏（进度 + 状态 + 按钮）
        self._build_foot()

    def _build_engine_bar(self) -> None:
        card = Card(self.inner, radius=10, padx=14, pady=10)
        card.pack(fill="x", padx=4, pady=8)
        b = card.body()
        row = tk.Frame(b, bg=CARD_BG)
        row.pack(fill="x")
        row.columnconfigure(1, weight=1)
        tk.Label(row, text="脱敏核心：", bg=CARD_BG, fg=TEXT,
                 font=BODY_FONT).grid(row=0, column=0, sticky="w")
        self.lbl_engine = tk.Label(row, text="正在检测…", fg=TEXT_SOFT,
                                   bg=CARD_BG, font=BODY_FONT, cursor="hand2")
        self.lbl_engine.grid(row=0, column=1, sticky="w")
        self.lbl_engine.bind("<Button-1>", self._engine_popup)
        self.btn_detect = ModernButton(
            row, label="🔄 重新检测", variant="neutral", width=110,
            command=self.detect_engine)
        self.btn_detect.grid(row=0, column=2, padx=(8, 0))

    def _build_file_card(self) -> None:
        self.file_card = Card(self.inner, title="待处理文件", radius=12)
        self.file_card.pack(fill="x", padx=4, pady=8)
        b = self.file_card.body()
        b.columnconfigure(0, weight=1)

        # 工具栏（图标 + 文字）
        tool = tk.Frame(b, bg=CARD_BG)
        tool.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.btn_add = ModernButton(tool, label="📂 添加文件", variant="primary",
                                    command=self.add_files)
        self.btn_add.pack(side="left", padx=(0, 6))
        self.btn_add_folder = ModernButton(
            tool, label="📁 添加文件夹", variant="primary",
            command=self.add_folder)
        self.btn_add_folder.pack(side="left", padx=(0, 6))
        self.btn_remove = ModernButton(
            tool, label="🗑 移除所选", variant="neutral",
            command=self.remove_selected)
        self.btn_remove.pack(side="left", padx=(0, 6))
        self.btn_clear = ModernButton(
            tool, label="🧹 清空列表", variant="neutral",
            command=self.clear_files)
        self.btn_clear.pack(side="left", padx=(0, 6))
        self.btn_whitelist = ModernButton(
            tool, label="📝 白名单", variant="neutral",
            command=self.open_whitelist_editor)
        self.btn_whitelist.pack(side="left", padx=(0, 6))
        self.btn_mark_miss = ModernButton(
            tool, label="🚩 标记漏报", variant="neutral",
            command=self.open_mark_missed)
        self.btn_mark_miss.pack(side="left", padx=(0, 6))

        # 拖拽上传区（虚线框，点击亦可选择文件）
        self.drop_frame = tk.Frame(b, bg=CARD_BG)
        self.drop_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self._drop_cv = tk.Canvas(self.drop_frame, height=46, bg=CARD_BG,
                                  highlightthickness=0)
        self._drop_cv.pack(fill="x")
        self._drop_cv.bind("<Configure>", lambda e: self._draw_drop_hint())
        self._draw_drop_hint()
        self.drop_frame.bind("<Button-1>", lambda e: self.add_files())
        self._drop_cv.bind("<Button-1>", lambda e: self.add_files())
        self._setup_dragdrop(self.drop_frame)

        # 列表 + 滚动条
        list_frame = tk.Frame(b, bg=CARD_BG)
        list_frame.grid(row=2, column=0, sticky="nsew")
        b.rowconfigure(2, weight=1, minsize=160)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        cols = ("name", "size", "status", "detail")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings",
                                 selectmode="extended",
                                 style="Treeview")
        self.tree.heading("name", text="文件名")
        self.tree.heading("size", text="大小")
        self.tree.heading("status", text="状态")
        self.tree.heading("detail", text="说明")
        self.tree.column("name", width=240, anchor="w")
        self.tree.column("size", width=70, anchor="e", stretch=False)
        self.tree.column("status", width=60, anchor="center", stretch=False)
        self.tree.column("detail", width=240, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(list_frame, orient="vertical",
                           command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)

        self.tree.tag_configure("ok", foreground="#2E9E5B")
        self.tree.tag_configure("err", foreground="#D9534F")
        self.tree.tag_configure("warn", foreground="#B26A00")
        self.tree.tag_configure("busy", foreground="#0B5CAD")
        self.tree.tag_configure("odd", background=ROW_ALT)
        self.tree.tag_configure("hover", background=HOVER_ROW)
        self.tree.bind("<Motion>", self._tree_hover)

        self.lbl_hint = tk.Label(
            b,
            text="支持 .docx / .pdf / .xlsx / .pptx；可一次添加多个文件批量处理",
            fg=TEXT_SOFT, bg=CARD_BG, font=BODY_FONT)
        self.lbl_hint.grid(row=3, column=0, sticky="w", pady=(6, 0))

    def _draw_drop_hint(self) -> None:
        cv = self._drop_cv
        cv.delete("all")
        w = max(cv.winfo_width(), 10)
        h = 46
        cv.configure(width=w, height=h)
        rounded_rect(cv, 2, 2, w - 2, h - 2, 10,
                     outline=ACCENT_BLUE, width=2, dash=(6, 4), fill="#FBFEFE")
        cv.create_text(w / 2, h / 2,
                       text="⬇ 拖拽文件到此处添加（或点击此处选择）",
                       font=BODY_FONT, fill=ACCENT_BLUE_DK, anchor="center")

    def _tree_hover(self, event) -> None:
        try:
            iid = self.tree.identify_row(event.y)
        except Exception:
            return
        if iid == self._hover_iid:
            return
        # 还原上一个
        if self._hover_iid and self.tree.exists(self._hover_iid):
            tags = list(self.tree.item(self._hover_iid, "tags"))
            if "hover" in tags:
                tags.remove("hover")
                self.tree.item(self._hover_iid, tags=tuple(tags))
        # 高亮当前
        if iid:
            tags = list(self.tree.item(iid, "tags"))
            if "hover" not in tags:
                tags.append("hover")
                self.tree.item(iid, tags=tuple(tags))
        self._hover_iid = iid

    def _build_options_card(self) -> None:
        card = Card(self.inner, title="处理选项", radius=12)
        card.pack(fill="x", padx=4, pady=8)
        b = card.body()
        b.columnconfigure(1, weight=1)

        tk.Label(b, text="脱敏模式：", bg=CARD_BG, fg=TEXT,
                 font=BODY_FONT).grid(row=0, column=0, sticky="w", pady=4)
        self.var_mode = StringVar(value=self.cfg["mode"])
        self.cmb_mode = ttk.Combobox(
            b, textvariable=self.var_mode, state="readonly",
            values=[MODE_LABELS[m] for m in MODES], width=42,
            font=BODY_FONT)
        self.cmb_mode.set(MODE_LABELS.get(self.cfg["mode"],
                                          MODE_LABELS["smart"]))
        self.cmb_mode.grid(row=0, column=1, sticky="w", padx=(0, 8), pady=4)
        self.cmb_mode.bind("<<ComboboxSelected>>",
                           lambda _e: self._sync_mode_dependent())

        # 检测灵敏度（仅 smart 模式可见/生效）
        self.lbl_sens_title = tk.Label(b, text="检测灵敏度：", bg=CARD_BG,
                                       fg=TEXT, font=BODY_FONT)
        self.lbl_sens_title.grid(row=1, column=0, sticky="w", pady=(6, 0))
        sens_row = tk.Frame(b, bg=CARD_BG)
        sens_row.grid(row=1, column=1, columnspan=2, sticky="w", pady=(6, 0))
        self.var_sens = StringVar(
            value=self.cfg.get("sensitivity") or SENSITIVITY_DEFAULT)
        self._sens_radios: list[ttk.Radiobutton] = []
        self._sens_tooltips: dict[str, Tooltip] = {}
        for key in SENSITIVITY_KEYS:
            lvl = SENSITIVITY_LEVELS[key]
            rb = ttk.Radiobutton(
                sens_row, text=lvl["label"], value=key,
                variable=self.var_sens, command=self._on_sens_change)
            rb.pack(side="left", padx=(0, 8))
            self._sens_tooltips[key] = Tooltip(rb, lvl["desc"])
            self._sens_radios.append(rb)

        self.lbl_sens_hint = tk.Label(
            b, text="（鼠标悬停在灵敏度选项上查看详细说明）",
            fg=TEXT_SOFT, bg=CARD_BG, font=BODY_FONT, wraplength=660,
            justify="left")
        self.lbl_sens_hint.grid(row=2, column=0, columnspan=3, sticky="w",
                                pady=(2, 0))
        self._sens_widgets = [self.lbl_sens_title, sens_row, self.lbl_sens_hint]

        self.var_mapping = BooleanVar(value=bool(self.cfg["save_mapping"]))
        ttk.Checkbutton(b, text="导出映射表（用于还原）",
                        variable=self.var_mapping,
                        style="TCheckbutton").grid(row=0, column=2, sticky="e")

        tk.Label(b, text="输出位置：", bg=CARD_BG, fg=TEXT,
                 font=BODY_FONT).grid(row=3, column=0, sticky="w", pady=(6, 0))
        outrow = tk.Frame(b, bg=CARD_BG)
        outrow.grid(row=3, column=1, columnspan=2, sticky="ew", pady=(6, 0))
        outrow.columnconfigure(1, weight=1)

        self.var_outmode = StringVar(value=self.cfg["output_mode"])
        ttk.Radiobutton(outrow, text="原文件所在目录", value="source",
                        variable=self.var_outmode, style="TRadiobutton",
                        command=self._sync_out_state).grid(
            row=0, column=0, sticky="w")
        ttk.Radiobutton(outrow, text="指定目录", value="custom",
                        variable=self.var_outmode, style="TRadiobutton",
                        command=self._sync_out_state).grid(
            row=1, column=0, sticky="w")

        self.var_outdir = StringVar(value=self.cfg["custom_output"])
        self.ent_outdir = ttk.Entry(outrow, textvariable=self.var_outdir,
                                    font=BODY_FONT)
        self.ent_outdir.grid(row=1, column=1, sticky="ew", padx=(8, 6))
        self.btn_outdir = ModernButton(outrow, label="浏览…", variant="neutral",
                                       width=70, command=self.pick_outdir)
        self.btn_outdir.grid(row=1, column=2)

    def _build_foot(self) -> None:
        card = Card(self.root, radius=12, padx=16, pady=12)
        card.pack(side="bottom", fill="x", padx=10, pady=(0, 10))
        b = card.body()
        b.columnconfigure(0, weight=1)
        b.rowconfigure(1, minsize=30)

        self.pb = ttk.Progressbar(b, mode="determinate", maximum=100,
                                  style="Mint.Horizontal.TProgressbar")
        self.pb.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.lbl_status = tk.Label(b, text="就绪，请添加文件", fg=TEXT_SOFT,
                                   bg=CARD_BG, font=BODY_FONT, anchor="w")
        self.lbl_status.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        btns = tk.Frame(b, bg=CARD_BG)
        btns.grid(row=0, column=1, rowspan=2, sticky="e")
        self.btn_open = ModernButton(btns, label="📂 打开输出文件夹",
                                     variant="neutral", width=130,
                                     command=self.open_output, state="disabled")
        self.btn_open.pack(side="left", padx=(0, 6))
        self.btn_cancel = ModernButton(btns, label="取消", variant="neutral",
                                       width=70, command=self.cancel_run,
                                       state="disabled")
        self.btn_cancel.pack(side="left", padx=(0, 6))
        self.btn_run = ModernButton(btns, label="✨ 开始脱敏", variant="coral",
                                    width=110, command=self.start_run)
        self.btn_run.pack(side="left")

    # -- 拖拽上传（Windows 原生，失败则仅保留点击/按钮） -----------------

    def _setup_dragdrop(self, widget) -> None:
        if not IS_WINDOWS:
            return
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            shell32 = ctypes.windll.shell32
            GWL_WNDPROC = -4
            WM_DROPFILES = 0x0233
            DragQueryFileW = shell32.DragQueryFileW
            DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT,
                                       wintypes.LPWSTR, wintypes.UINT]
            DragQueryFileW.restype = wintypes.UINT
            DragFinish = shell32.DragFinish
            DragFinish.argtypes = [wintypes.HANDLE]

            hwnd = widget.winfo_id()
            oldproc = ctypes.c_void_p(
                user32.GetWindowLongPtrW(hwnd, GWL_WNDPROC))

            def wndproc(hWnd, Msg, wParam, lParam):
                if Msg == WM_DROPFILES:
                    try:
                        count = DragQueryFileW(wParam, 0xFFFFFFFF, None, 0)
                        buf = ctypes.create_unicode_buffer(1024)
                        paths = []
                        for i in range(count):
                            DragQueryFileW(wParam, i, buf, 1024)
                            paths.append(buf.value)
                        DragFinish(wParam)
                        widget.after(
                            0, lambda: self._add([Path(p) for p in paths]))
                    except Exception:
                        pass
                    return 0
                return user32.CallWindowProcW(
                    oldproc, hWnd, Msg, wParam, lParam)

            WCALLBACK = ctypes.WINFUNCTYPE(
                wintypes.LRESULT, wintypes.HWND, wintypes.UINT,
                wintypes.WPARAM, wintypes.LPARAM)
            self._drop_proc = WCALLBACK(wndproc)
            if not user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, self._drop_proc):
                return
            shell32.DragAcceptFiles(hwnd, True)
        except Exception:
            pass

    # -- 事件绑定 ---------------------------------------------------------

    def _bind_events(self) -> None:
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # -- 引擎检测 ---------------------------------------------------------

    def detect_engine(self) -> None:
        self.btn_detect.configure(state="disabled")
        self.btn_run.configure(state="disabled")
        self.lbl_engine.configure(text="正在检测…", fg=TEXT_SOFT)

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
        p = filedialog.askopenfilename(title="选择 mask-tool 可执行文件",
                                       filetypes=types)
        if p:
            self.cfg["mask_tool_path"] = p
            settings.save(self.cfg)
            self.detect_engine()

    def _engine_popup(self, event: tk.Event | None = None) -> None:
        if getattr(self, "_engine_menu_inst", None) is None:
            m = tk.Menu(self.root, tearoff=0)
            m.add_command(label="重新检测脱敏核心", command=self.detect_engine)
            m.add_separator()
            m.add_command(label="手动指定 mask-tool 路径…",
                          command=self.pick_tool_path)
            self._engine_menu_inst = m
        menu = self._engine_menu_inst
        try:
            x = event.x_root if event else self.root.winfo_pointerx()
            y = event.y_root if event else self.root.winfo_pointery()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    # -- 白名单编辑器（右上角「白名单…」按钮）-------------------------------

    def open_whitelist_editor(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("白名单管理 — 绝不脱敏的词")
        win.geometry("440x480")
        win.minsize(380, 380)
        win.configure(bg=BG_BLUE)
        win.transient(self.root)
        win.grab_set()

        items: list[str] = sorted(load_whitelist(force=True),
                                  key=lambda s: (len(s), s))
        fam, size = FAM, FSIZE

        card = Card(win, title="白名单管理 — 绝不脱敏的词", radius=12)
        card.pack(fill="both", expand=True, padx=16, pady=16)
        frm = card.body()
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(1, weight=1)

        tk.Label(
            frm,
            text="以下词条在处理时「绝不脱敏」（按整条识别结果精确匹配）：",
            fg=TEXT_SOFT, bg=CARD_BG, font=BODY_FONT,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 6))

        listbox = tk.Listbox(frm, selectmode="extended", font=(fam, size),
                             bg="#FFFFFF", fg=TEXT, relief="flat",
                             highlightbackground=BORDER, highlightthickness=1)
        listbox.grid(row=1, column=0, columnspan=2, sticky="nsew",
                     padx=6, pady=(0, 6))
        sb = ttk.Scrollbar(frm, orient="vertical", command=listbox.yview)
        sb.grid(row=1, column=2, sticky="ns", pady=(0, 6))
        listbox.configure(yscrollcommand=sb.set)
        for w in items:
            listbox.insert("end", w)

        def refresh() -> None:
            listbox.delete(0, "end")
            for w in sorted(set(items), key=lambda s: (len(s), s)):
                listbox.insert("end", w)

        add_frm = tk.Frame(frm, bg=CARD_BG)
        add_frm.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 6))
        add_frm.columnconfigure(0, weight=1)
        ent = ttk.Entry(add_frm, font=(fam, size))
        ent.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ent.focus_set()

        def do_add() -> None:
            raw = ent.get().strip()
            if not raw:
                return
            for part in re.split(r"[,，、;；\s]+", raw):
                part = part.strip()
                if part and part not in items:
                    items.append(part)
            ent.delete(0, "end")
            refresh()
            ent.focus_set()

        ent.bind("<Return>", lambda _e: do_add())
        ModernButton(add_frm, label="添加", variant="primary", width=70,
                     command=do_add).grid(row=0, column=1)

        btn_frm = tk.Frame(frm, bg=CARD_BG)
        btn_frm.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(2, 0))
        btn_frm.columnconfigure(0, weight=1)
        btn_frm.columnconfigure(1, weight=1)
        btn_frm.columnconfigure(2, weight=1)

        def do_del() -> None:
            for i in reversed(list(listbox.curselection())):
                items.remove(listbox.get(i))
            refresh()

        ModernButton(btn_frm, label="删除所选", variant="neutral", width=110,
                     command=do_del).grid(row=0, column=0, padx=(0, 4),
                                          sticky="w")

        ModernButton(btn_frm, label="取消", variant="neutral", width=90,
                     command=win.destroy).grid(row=0, column=1, padx=(0, 4))

        def do_save() -> None:
            ok = save_whitelist(items)
            if not ok:
                self._notify("error", "保存失败",
                             "无法写入 whitelist.txt（可能程序目录为只读）。")
                return
            load_whitelist(force=True)
            if getattr(self, "engine", None) is not None:
                try:
                    self.engine.refresh_whitelist()
                except Exception:
                    pass
            self._notify("info", "已保存",
                         f"白名单已更新，共 {len(set(items))} 条。"
                         f"下次脱敏自动生效。")
            win.destroy()

        ModernButton(btn_frm, label="保存", variant="coral", width=90,
                     command=do_save).grid(row=0, column=2, sticky="e")

    # -- 标记漏报（「标记漏报…」按钮）-------------------------------------

    def open_mark_missed(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("标记漏报 — 加入用户词库")
        win.geometry("520x470")
        win.minsize(420, 380)
        win.configure(bg=BG_BLUE)
        win.transient(self.root)
        win.grab_set()
        fam, size = FAM, FSIZE

        card = Card(win, title="标记漏报 — 加入用户词库", radius=12)
        card.pack(fill="both", expand=True, padx=16, pady=16)
        frm = card.body()
        frm.columnconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(2, weight=0)
        frm.rowconfigure(1, weight=1)

        tk.Label(
            frm,
            text="① 粘贴原文并选中漏报片段（或直接在下方的「片段」框输入）：",
            fg=TEXT_SOFT, bg=CARD_BG, font=BODY_FONT,
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

        txt = tk.Text(frm, height=8, wrap="word", font=(fam, size),
                      bg="#FFFFFF", fg=TEXT, relief="flat",
                      highlightbackground=BORDER, highlightthickness=1)
        txt.grid(row=1, column=0, columnspan=3, sticky="nsew")
        txt.insert("1.0",
                   "在此粘贴包含漏报实体的原文，用鼠标选中那段应被脱敏的文字…")

        tk.Label(frm, text="片段：", bg=CARD_BG, fg=TEXT,
                 font=BODY_FONT).grid(row=2, column=0, sticky="w", pady=(6, 2))
        var_frag = StringVar()
        ent_frag = ttk.Entry(frm, textvariable=var_frag, font=(fam, size))
        ent_frag.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(6, 2))

        def use_selection() -> None:
            try:
                sel = txt.get("sel.first", "sel.last").strip()
            except Exception:
                sel = ""
            if sel:
                var_frag.set(sel)

        ModernButton(frm, label="用选中内容", variant="neutral", width=110,
                     command=use_selection).grid(row=3, column=0, sticky="w",
                                                 pady=2)
        tk.Label(frm, text="实体类型：", bg=CARD_BG, fg=TEXT,
                 font=BODY_FONT).grid(row=3, column=1, sticky="e", pady=2)
        var_cat = StringVar(value=LEX_CATEGORIES[0][0])
        cmb_cat = ttk.Combobox(
            frm, textvariable=var_cat, state="readonly",
            values=[c[0] for c in LEX_CATEGORIES], width=14, font=(fam, size))
        cmb_cat.grid(row=3, column=2, sticky="e", pady=2)

        tk.Label(
            frm,
            text="② 确认片段与类型后，点「加入词库并保存」，该词下次脱敏即生效。",
            fg=TEXT_SOFT, bg=CARD_BG, font=BODY_FONT,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 2))

        btn_frm = tk.Frame(frm, bg=CARD_BG)
        btn_frm.grid(row=5, column=0, columnspan=3, sticky="e", pady=(4, 0))

        def do_add() -> None:
            word = var_frag.get().strip()
            if not word:
                self._notify("warn", "缺少片段",
                             "请先输入或选中要标记的文本片段。")
                return
            label_to_key = {lbl: k for lbl, k in LEX_CATEGORIES}
            cat = label_to_key.get(var_cat.get(), "custom")
            lex = self.cfg.setdefault("user_lexicon", {})
            cur = list(lex.get(cat, []))
            if word in cur:
                self._notify("info", "已存在",
                             f"「{word}」已在「{var_cat.get()}」词库中，"
                             f"无需重复添加。")
                win.destroy()
                return
            cur.append(word)
            lex[cat] = cur
            settings.save(self.cfg)
            try:
                self._refresh_lex_list()
            except Exception:
                pass
            self._notify(
                "info", "已加入词库",
                f"「{word}」已作为「{var_cat.get()}」加入用户词库并保存。\n"
                f"下次脱敏（尤其是 strict / smart 模式）将自动命中脱敏。")
            win.destroy()

        ModernButton(btn_frm, label="取消", variant="neutral", width=90,
                     command=win.destroy).grid(row=0, column=0, padx=(0, 6))
        ModernButton(btn_frm, label="加入词库并保存", variant="coral",
                     width=150, command=do_add).grid(row=0, column=1)

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
            self._notify("error", "读取失败", f"无法读取该文件夹：\n{exc}")
            return
        if not found:
            self._notify("info", "没有找到文件",
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
            parity = "odd" if (self._row_index % 2 == 1) else ""
            self._row_index += 1
            tag = "warn" if p.suffix.lower() in REPORT_ONLY_EXTS else ""
            note = "PDF 仅输出检测报告" \
                if p.suffix.lower() in REPORT_ONLY_EXTS else ""
            base_tags = (parity, tag) if (parity or tag) else ()
            self.tree.insert("", "end", iid=str(p),
                             values=(p.name, human_size(p), "待处理", note),
                             tags=base_tags)
        self._refresh_counter()
        if skipped:
            self._notify(
                "warn", "部分文件已跳过",
                f"有 {skipped} 个文件格式不受支持，已自动跳过。\n"
                "仅支持 .docx / .pdf / .xlsx / .pptx")
        if added:
            self.lbl_status.configure(text=f"已添加 {added} 个文件",
                                      fg=TEXT_SOFT)

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
        self._row_index = 0
        self.pb["value"] = 0
        self.lbl_status.configure(text="就绪", fg=TEXT_SOFT)
        self._refresh_counter()

    def _refresh_counter(self) -> None:
        n = len(self.files)
        if getattr(self, "file_card", None) is not None:
            self.file_card.set_title(
                f"待处理文件（{n}）" if n else "待处理文件")

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
        smart = self._current_mode() == "smart"
        if smart:
            for w in self._sens_widgets:
                w.grid()
            key = self.var_sens.get()
            lvl = SENSITIVITY_LEVELS.get(
                key, SENSITIVITY_LEVELS[SENSITIVITY_DEFAULT])
        else:
            for w in self._sens_widgets:
                w.grid_remove()

    # -- 识别引擎（高级） -------------------------------------------------

    def _build_engine_widgets(self, parent: tk.Frame) -> None:
        card = Card(parent, title="识别引擎（高级）", radius=12)
        card.pack(fill="x", padx=4, pady=8)
        frame = card.body()
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(3, weight=1)

        self._NER_DISPLAY = {
            "auto": "自动（优先 spaCy）",
            "spacy": "spaCy（领域模型）",
            "jieba": "jieba（默认）",
        }
        self._NER_DISPLAY_TO_KEY = {v: k for k, v in self._NER_DISPLAY.items()}
        _saved = self.cfg.get("ner_backend", "auto")
        self.var_ner_backend = StringVar(
            value=self._NER_DISPLAY.get(_saved, self._NER_DISPLAY["auto"]))
        tk.Label(frame, text="识别引擎：", bg=CARD_BG, fg=TEXT,
                 font=BODY_FONT).grid(row=0, column=0, sticky="w", pady=(2, 0))
        self.cmb_ner = ttk.Combobox(
            frame, textvariable=self.var_ner_backend, state="readonly",
            values=list(self._NER_DISPLAY.values()), width=28,
            font=BODY_FONT)
        self.cmb_ner.grid(row=0, column=1, sticky="w", padx=(0, 8), pady=(2, 0))
        self.cmb_ner.bind("<<ComboboxSelected>>",
                          lambda _e: self._on_ner_change())

        tk.Label(frame, text="spaCy 模型：", bg=CARD_BG, fg=TEXT,
                 font=BODY_FONT).grid(row=1, column=0, sticky="w", pady=(4, 0))
        pathrow = tk.Frame(frame, bg=CARD_BG)
        pathrow.grid(row=1, column=1, columnspan=3, sticky="ew", pady=(4, 0))
        pathrow.columnconfigure(0, weight=1)
        self.var_spacy_model = StringVar(
            value=self.cfg.get("spacy_model", ""))
        self.ent_spacy = ttk.Entry(pathrow, textvariable=self.var_spacy_model,
                                   font=BODY_FONT)
        self.ent_spacy.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.btn_spacy = ModernButton(pathrow, label="浏览…", variant="neutral",
                                      width=70, command=self.pick_spacy_model)
        self.btn_spacy.grid(row=0, column=1)

        tk.Label(frame, text="置信度下限：", bg=CARD_BG, fg=TEXT,
                 font=BODY_FONT).grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.var_min_conf = StringVar(
            value=str(self.cfg.get("min_confidence", 0.0)))
        sb = ttk.Spinbox(
            frame, from_=0.0, to=1.0, increment=0.05,
            textvariable=self.var_min_conf, width=8, font=BODY_FONT)
        sb.grid(row=2, column=1, sticky="w", padx=(0, 8), pady=(4, 0))
        tk.Label(frame,
                 text="全局兜底下限（0=按类型自动：公司0.75/人名0.60/地点0.80/项目0.75）",
                 fg=TEXT_SOFT, bg=CARD_BG, font=BODY_FONT).grid(
            row=2, column=2, columnspan=2, sticky="w", pady=(4, 0))

        self.lbl_ner_status = tk.Label(
            frame, text="", fg=TEXT_SOFT, bg=CARD_BG, font=BODY_FONT,
            justify="left")
        self.lbl_ner_status.grid(row=3, column=0, columnspan=4, sticky="w",
                                 pady=(4, 0))
        self._ner_status_tooltip = Tooltip(self.lbl_ner_status, "")
        self._refresh_ner_status()

    def _on_ner_change(self) -> None:
        self._refresh_ner_status()

    def _ner_backend_key(self) -> str:
        return self._NER_DISPLAY_TO_KEY.get(
            self.var_ner_backend.get(), "auto")

    def _min_conf_value(self) -> float:
        try:
            v = float(self.var_min_conf.get().strip())
        except (ValueError, AttributeError):
            return 0.0
        return min(max(v, 0.0), 1.0)

    def _refresh_ner_status(self) -> None:
        try:
            st = ner_status()
        except Exception:
            self.lbl_ner_status.configure(text="引擎状态：未知")
            self._ner_status_tooltip.set_text("无法获取识别引擎状态")
            return
        active = st.get("active", "jieba")
        model = st.get("model") or "（自动发现）"
        if active == "spacy":
            short = f"当前识别引擎：spaCy（模型 {model}）"
            detail = f"当前使用 spaCy 模型：{model}"
            self.lbl_ner_status.configure(text=short)
            self._ner_status_tooltip.set_text(detail)
        elif st.get("model_missing", False):
            expected = self.var_spacy_model.get().strip() or "zh_core_web_md"
            short = "当前识别引擎：jieba（内置）⚠️ 未安装 spaCy 模型，已自动回退"
            detail = (
                "当前识别引擎：jieba（内置）。\n"
                "⚠️ 未安装 spaCy 模型，已自动回退；\n"
                f"可运行  python -m spacy download {expected}  "
                f"或把模型目录放到程序根目录的 models/ 下；"
                f"当前已自动回退到内置 jieba。"
            )
            self.lbl_ner_status.configure(text=short)
            self._ner_status_tooltip.set_text(detail)
        else:
            reason = st.get("reason") or ""
            txt = "当前识别引擎：jieba（内置）"
            if reason:
                txt += f"  · spaCy：{reason}"
            self.lbl_ner_status.configure(text=txt)
            self._ner_status_tooltip.set_text(txt)

    def pick_spacy_model(self) -> None:
        p = filedialog.askdirectory(
            title="选择 spaCy 模型目录（含 config.cfg / meta.json）",
            initialdir=self._initial_dir(),
        )
        if not p:
            return
        self.var_spacy_model.set(p)
        self.ent_spacy.delete(0, "end")
        self.ent_spacy.insert(0, p)
        self._refresh_ner_status()

    # -- 用户词库 ---------------------------------------------------------

    def _build_lexicon_widgets(self, parent: tk.Frame) -> None:
        card = Card(parent, title="用户词库（strict 模式主要依赖此词典）",
                    radius=12)
        card.pack(fill="x", padx=4, pady=8)
        frame = card.body()
        frame.columnconfigure(2, weight=1)

        self.var_lex_cat = StringVar(value=LEX_CATEGORIES[0][0])
        tk.Label(frame, text="类别：", bg=CARD_BG, fg=TEXT,
                 font=BODY_FONT).grid(row=0, column=0, sticky="w")
        self.cmb_lex_cat = ttk.Combobox(
            frame, textvariable=self.var_lex_cat, state="readonly",
            values=[c[0] for c in LEX_CATEGORIES], width=12, font=BODY_FONT)
        self.cmb_lex_cat.grid(row=0, column=1, sticky="w", padx=(0, 8))

        self.PLACEHOLDER_LEX = "输入词，逗号分隔，回车添加"
        self.ent_lex = ttk.Entry(frame, width=26, font=BODY_FONT)
        self.ent_lex.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        self.ent_lex.insert(0, self.PLACEHOLDER_LEX)
        self.ent_lex.configure(foreground="#999999")
        self.ent_lex.bind("<FocusIn>", self._lex_entry_focus_in)
        self.ent_lex.bind("<FocusOut>", self._lex_entry_focus_out)
        self.ent_lex.bind("<Return>", lambda e: self._add_lex_word())

        ModernButton(frame, label="添加", variant="primary", width=70,
                     command=self._add_lex_word).grid(
            row=0, column=3, padx=(0, 6))
        ModernButton(frame, label="导入TXT", variant="neutral", width=80,
                     command=self._import_lex_txt).grid(
            row=0, column=4, padx=(0, 6))
        ModernButton(frame, label="清空", variant="neutral", width=70,
                     command=self._clear_lex).grid(row=0, column=5)

        list_row = tk.Frame(frame, bg=CARD_BG)
        list_row.grid(row=1, column=0, columnspan=6, sticky="ew", pady=(6, 0))
        list_row.columnconfigure(0, weight=1)

        self.lst_lex = Listbox(list_row, height=3, selectmode="extended",
                               exportselection=False, bg="#FFFFFF", fg=TEXT,
                               relief="flat", highlightbackground=BORDER,
                               highlightthickness=1)
        self.lst_lex.grid(row=0, column=0, sticky="nsew")
        sb_lex = ttk.Scrollbar(list_row, orient="vertical",
                               command=self.lst_lex.yview)
        sb_lex.grid(row=0, column=1, sticky="ns")
        self.lst_lex.configure(yscrollcommand=sb_lex.set)
        self.lst_lex.bind("<Delete>", lambda e: self._remove_lex_selected())

        self.lbl_lex_count = tk.Label(list_row, text="已收录 0 个词",
                                      fg=TEXT_SOFT, bg=CARD_BG, font=BODY_FONT)
        self.lbl_lex_count.grid(row=1, column=0, sticky="w", pady=(2, 0))
        ModernButton(list_row, label="删除所选", variant="neutral", width=90,
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
            self._notify("error", "读取失败", f"无法读取文件：\n{exc}")
            return
        parsed = self._parse_lex_txt(data)
        if not parsed:
            self._notify("info", "提示", "文件中没有可导入的词。")
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
        self._notify(
            "info", "导入完成",
            f"已导入 {total_added} 个新词：\n" + "；".join(summary),
        )

    def _parse_lex_txt(self, text: str) -> dict[str, list[str]]:
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
                    prefix.lower() if prefix.lower() in label_to_key.values()
                    else None)
                if cat is None:
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
        if not self._confirm("确认清空", "确定要清空全部用户词库吗？"):
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
                                    initialdir=self.var_outdir.get()
                                    or self._initial_dir())
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
            self._notify("warn", "无法开始", err)
            return

        mode = self._current_mode()
        save_mapping = bool(self.var_mapping.get())
        custom = self.var_outmode.get() == "custom"
        outdir = Path(self.var_outdir.get().strip()) if custom else None
        sensitivity = self.var_sens.get() or SENSITIVITY_DEFAULT

        ner_backend = self._ner_backend_key()
        spacy_model = self.var_spacy_model.get().strip()
        min_conf = self._min_conf_value()

        self.cfg.update({
            "mode": mode,
            "sensitivity": sensitivity,
            "output_mode": self.var_outmode.get(),
            "custom_output": self.var_outdir.get().strip(),
            "save_mapping": save_mapping,
            "ner_backend": ner_backend,
            "spacy_model": spacy_model,
            "min_confidence": min_conf,
        })
        settings.save(self.cfg)

        set_ner_backend(ner_backend, spacy_model)
        set_min_confidence(min_conf)

        self.results.clear()
        for f in self.files:
            self.tree.item(str(f), values=(f.name, human_size(f),
                                          "排队中", ""), tags=())

        self._set_running(True)
        self.pb.configure(maximum=max(len(self.files), 1), value=0)
        self.btn_run.configure(label="⏳ 处理中…", state="disabled")

        if self.tool is None:
            self._notify("error", "缺少脱敏核心", INSTALL_HELP)
            return

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
                    min_confidence=min_conf,
                    ner_backend=ner_backend,
                    spacy_model=spacy_model,
                    on_progress=lambda i, n, p:
                        self.msg_q.put(("progress", (i, n, p))),
                    on_result=lambda r: self.msg_q.put(("result", r)),
                )
                self.msg_q.put(("done", outdir or (files[0].parent
                                                   if files else None)))
            except Exception:
                self.msg_q.put(("crash", traceback.format_exc()))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def cancel_run(self) -> None:
        if self.engine and self.running:
            self.engine.cancel()
            self.lbl_status.configure(text="正在取消，请等待当前文件结束…",
                                      fg=TEXT_SOFT)
            self.btn_cancel.configure(state="disabled")

    def _set_running(self, running: bool) -> None:
        self.running = running
        state = "disabled" if running else "normal"
        for w in (self.btn_run, self.btn_detect, self.btn_whitelist,
                  self.btn_mark_miss, self.cmb_mode):
            w.configure(state="disabled" if running else
                        ("readonly" if w is self.cmb_mode else "normal"))
        self.btn_cancel.configure(state="normal" if running else "disabled")
        if not running:
            self.btn_run.configure(
                label="✨ 开始脱敏",
                state="normal" if self.engine else "disabled")
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
            self.engine = MaskEngine(payload,
                                     user_lexicon=self.cfg.get("user_lexicon")
                                     or {})
            ver = f"  {payload.version}" if payload.version else ""
            self.lbl_engine.configure(text=f"✓ 已就绪{ver}", fg="#2E9E5B")
            self.btn_detect.configure(state="normal")
            self.btn_run.configure(state="normal")
            self.lbl_status.configure(text="就绪，请添加文件", fg=TEXT_SOFT)

        elif kind == "engine_fail":
            self.tool = None
            self.engine = None
            self.lbl_engine.configure(text="✕ 未检测到 mask-tool",
                                      fg="#D9534F")
            self.btn_detect.configure(state="normal")
            self.btn_run.configure(state="disabled")
            self.lbl_status.configure(text="请先安装脱敏核心", fg=TEXT_SOFT)
            self._notify("error", "缺少脱敏核心", INSTALL_HELP)

        elif kind == "progress":
            i, n, p = payload
            self.pb.configure(value=i - 1)
            self.lbl_status.configure(text=f"正在处理 {i}/{n}：{p.name}",
                                      fg=TEXT)
            if self.tree.exists(str(p)):
                self.tree.item(str(p),
                               values=(p.name, human_size(p), "处理中", "…"),
                               tags=("busy",))
                self.tree.see(str(p))

        elif kind == "result":
            r: FileResult = payload
            iid = str(r.source)
            tag = "ok" if r.ok and not r.report_only else (
                "warn" if r.ok else "err")
            if self.tree.exists(iid):
                self.tree.item(iid, values=(r.source.name, human_size(r.source),
                                            r.status_text, r.message),
                               tags=(tag,))
            self.results[iid] = r
            self.pb.configure(value=self.pb["value"] + 1)
            if r.ok and r.output:
                self.last_output_dir = r.output.parent

        elif kind == "done":
            self._finish(payload)

        elif kind == "crash":
            self._set_running(False)
            self.btn_run.configure(label="✨ 开始脱敏", state="normal"
                                   if self.engine else "disabled")
            self.lbl_status.configure(text="处理异常中止", fg="#D9534F")
            self._notify(
                "error", "程序内部错误",
                "处理过程中出现未预期的问题，本次任务已中止。\n"
                "文档未被修改，可重试或减少一次处理的文件数量。",
            )

    def _finish(self, outdir: Path | None) -> None:
        self._set_running(False)
        self.btn_run.configure(label="✨ 开始脱敏")
        self.pb.configure(value=self.pb["maximum"])

        total = len(self.results)
        ok = sum(1 for r in self.results.values() if r.ok)
        fail = total - ok
        report_only = sum(1 for r in self.results.values()
                          if r.ok and r.report_only)
        cancelled = total < len(self.files)

        if outdir and Path(outdir).is_dir():
            self.last_output_dir = Path(outdir)
        self.btn_open.configure(state="normal" if self.last_output_dir
                                else "disabled")

        head = "处理已取消" if cancelled else "处理完成"
        self.lbl_status.configure(
            text=f"{'✅' if not fail else '❌'} {head}：成功 {ok} 个，"
                 f"失败 {fail} 个（共 {total} 个）",
            fg="#2E9E5B" if not fail else "#D9534F",
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
            self._notify("warn", head, body)
        else:
            self._notify("info", head, body)

    # -- 圆角弹窗（我知道了 / 确认） --------------------------------------

    def _notify(self, kind: str, title: str, message: str) -> None:
        d = tk.Toplevel(self.root)
        d.transient(self.root)
        d.overrideredirect(True)
        d.grab_set()
        d.configure(bg=BG_BLUE)
        card = Card(d, radius=14, padx=22, pady=18)
        card.pack(padx=20, pady=20)
        b = card.body()
        icon = {"error": "❌", "warn": "⚠️", "info": "✅",
                "ok": "✅"}.get(kind, "ℹ️")
        tk.Label(b, text=f"{icon} {title}", font=SUB_FONT, bg=CARD_BG,
                 fg=TEXT).pack(anchor="w", pady=(0, 8))
        tk.Label(b, text=message, font=BODY_FONT, bg=CARD_BG, fg=TEXT,
                 justify="left", wraplength=420).pack(anchor="w")
        ModernButton(b, label="我知道了", variant="primary",
                     command=d.destroy).pack(anchor="e", pady=(12, 0))
        self._center_dialog(d)
        d.wait_window(d)

    def _confirm(self, title: str, message: str) -> bool:
        res = {"ok": False}
        d = tk.Toplevel(self.root)
        d.transient(self.root)
        d.overrideredirect(True)
        d.grab_set()
        d.configure(bg=BG_BLUE)
        card = Card(d, radius=14, padx=22, pady=18)
        card.pack(padx=20, pady=20)
        b = card.body()
        tk.Label(b, text=f"⚠️ {title}", font=SUB_FONT, bg=CARD_BG,
                 fg=TEXT).pack(anchor="w", pady=(0, 8))
        tk.Label(b, text=message, font=BODY_FONT, bg=CARD_BG, fg=TEXT,
                 justify="left", wraplength=420).pack(anchor="w")
        row = tk.Frame(b, bg=CARD_BG)
        row.pack(anchor="e", pady=(12, 0))
        ModernButton(row, label="取消", variant="neutral", width=90,
                     command=lambda: (res.update(ok=False), d.destroy())
                     ).pack(side="left", padx=(0, 8))
        ModernButton(row, label="确定", variant="coral", width=90,
                     command=lambda: (res.update(ok=True), d.destroy())
                     ).pack(side="left")
        self._center_dialog(d)
        d.wait_window(d)
        return res["ok"]

    def _center_dialog(self, d) -> None:
        d.update_idletasks()
        w = d.winfo_width()
        h = d.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        d.geometry(f"+{max(x, 0)}+{max(y, 0)}")

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
            self._notify("info", "提示", "暂无可打开的输出目录。")

    def on_close(self) -> None:
        if self.running:
            if not self._confirm("确认退出", "任务正在进行中，确定要退出吗？"):
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
