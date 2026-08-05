# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller .spec
=================

直接双击运行 build.bat 即可调用 PyInstaller；这里保留 .spec 以便二次定制。
也支持在 macOS/Linux 上直接运行：
    python -m PyInstaller --noconfirm --clean masktool_gui.spec
"""

import sys
from pathlib import Path

ROOT = Path(SPECPATH)
ASSETS = ROOT / "assets"
ICON = ASSETS / "app.ico"
ICON_MAC = ASSETS / "app.icns"  # macOS 建议提供 .icns 图标
CONFIG_DIR = ASSETS / "mask_tool_config"
IS_MAC = sys.platform == "darwin"

block_cipher = None

# 把 mask-tool 的核心模块整体打包进 exe（排除 web/streamlit 以免膨胀）。
# 这样冻结后可通过 engine.py 的进程内调用直接使用，无需外部安装。
try:
    from PyInstaller.utils.hooks import collect_submodules
    _mt_hidden = collect_submodules(
        "mask_tool",
        filter=lambda name: not name.startswith("mask_tool.web"),
    )
except Exception:
    _mt_hidden = ["mask_tool", "mask_tool.core.pipeline", "mask_tool.models.config"]

# ---- spaCy + 中文模型（内嵌到程序，无需用户另行安装）----
# 用细粒度 hook 工具收集，避免 collect_all 在本机 PyInstaller 版本里把
# 数据文件 (src, dest) 元组混进 hiddenimports 导致 Analysis 报错。
#   · collect_submodules  → 模块名（字符串）放进 hiddenimports
#   · collect_data_files + copy_metadata → (src, dest) 元组放进 datas
#   · collect_dynamic_libs → 二进制 .so 放进 binaries
# 这样 spaCy 及其全部子模块 / 数据 / C 扩展（thinc、blis、preshed 等）连同
# zh_core_web_md 模型包（vocab/vectors/tokenizer/config）与其中文分词依赖
# spacy_pkuseg / pkuseg 的数据文件，整体卷进 exe，
# 冻结后由 ner_backend 的 discover 自动发现并使用。
def _safe_collect(fn, name):
    try:
        return list(fn(name))
    except Exception:
        return []


# 需要整体内嵌的包（spaCy 本体 + 中文模型 + 中文分词数据依赖）
_SPACY_PKGS = ("spacy", "zh_core_web_md", "spacy_pkuseg", "pkuseg")
try:
    from PyInstaller.utils.hooks import (
        collect_submodules, collect_data_files,
        collect_dynamic_libs, copy_metadata,
    )
    _spacy_hidden: list = []
    _spacy_datas: list = []
    _spacy_binaries: list = []
    for _p in _SPACY_PKGS:
        _spacy_hidden += _safe_collect(collect_submodules, _p)
        _spacy_datas += (_safe_collect(collect_data_files, _p)
                         + _safe_collect(copy_metadata, _p))
        _spacy_binaries += _safe_collect(collect_dynamic_libs, _p)
except Exception:
    _spacy_hidden, _spacy_datas, _spacy_binaries = [], [], []

# 资源文件：图标 + mask-tool 的词库配置
_datas = list(_spacy_datas)
_binaries = list(_spacy_binaries)
if ICON.is_file():
    _datas.append((str(ASSETS), "assets"))
if CONFIG_DIR.is_dir():
    _datas.append((str(CONFIG_DIR), "mask_tool_config"))

a = Analysis(
    [str(ROOT / "main.py")],
    pathex=[str(ROOT)],
    binaries=_binaries,
    datas=_datas,
    hiddenimports=[
        "tkinter", "tkinter.ttk", "tkinter.messagebox", "tkinter.filedialog",
        # PyMuPDF 在 pdf 适配器中延迟导入，静态分析可能漏掉，显式声明
        "fitz", "pymupdf",
        # 内嵌的 spaCy 后端与中文模型（spacy / zh_core_web_md / spacy_pkuseg 等
        # 的子模块已由 _spacy_hidden 收集，这里再显式兜底一次）
        "zh_core_web_md",
    ] + _spacy_hidden + _mt_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 单文件工具不需要这些模块，能显著缩小体积
        "matplotlib", "sklearn", "pytest",
        "django", "flask", "IPython", "jupyter", "setuptools", "pip",
        "unittest", "pydoc", "doctest",
        # mask-tool 的可选/重型依赖，按需排除（注意：PIL 必须保留，
        # 因为 pptx / openpyxl 在导入时就依赖它，排除会导致所有适配器加载失败）
        "streamlit", "streamlit_aggrid", "hanlp", "torch",
        "mask_tool.web", "mask_tool.web.app",
        # 注意：numpy / scipy 不可排除 —— spaCy / thinc 依赖它们；
        # 内嵌 spaCy 后即随其一起打包进 exe。
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="本地文档脱敏工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    # macOS .app 需要参数模拟，否则双击启动时 sys.argv 会异常
    argv_emulation=IS_MAC,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ICON_MAC if IS_MAC and ICON_MAC.is_file() else ICON)
    if (ICON_MAC.is_file() if IS_MAC else ICON.is_file())
    else None,
    version=str(ASSETS / "version.txt") if (ASSETS / "version.txt").is_file() else None,
)

# macOS 上额外生成 .app bundle，用户双击即可运行
if IS_MAC:
    app = BUNDLE(
        exe,
        name="本地文档脱敏工具.app",
        icon=str(ICON_MAC) if ICON_MAC.is_file() else None,
        bundle_identifier="com.yanguohui.masktool-gui",
    )
