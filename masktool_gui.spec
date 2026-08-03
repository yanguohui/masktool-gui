# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller .spec
=================

直接双击运行 build.bat 即可调用 PyInstaller；这里保留 .spec 以便二次定制。
"""

from pathlib import Path

ROOT = Path(SPECPATH)
ASSETS = ROOT / "assets"
ICON = ASSETS / "app.ico"
CONFIG_DIR = ASSETS / "mask_tool_config"

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

# 资源文件：图标 + mask-tool 的词库配置
_datas = []
if ICON.is_file():
    _datas.append((str(ASSETS), "assets"))
if CONFIG_DIR.is_dir():
    _datas.append((str(CONFIG_DIR), "mask_tool_config"))

a = Analysis(
    [str(ROOT / "main.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=_datas,
    hiddenimports=[
        "tkinter", "tkinter.ttk", "tkinter.messagebox", "tkinter.filedialog",
        # PyMuPDF 在 pdf 适配器中延迟导入，静态分析可能漏掉，显式声明
        "fitz", "pymupdf",
    ] + _mt_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 单文件工具不需要这些模块，能显著缩小体积
        "matplotlib", "numpy", "scipy", "sklearn", "pytest",
        "django", "flask", "IPython", "jupyter", "setuptools", "pip",
        "unittest", "pydoc", "doctest",
        # mask-tool 的可选/重型依赖，按需排除（注意：PIL 必须保留，
        # 因为 pptx / openpyxl 在导入时就依赖它，排除会导致所有适配器加载失败）
        "streamlit", "streamlit_aggrid", "hanlp", "torch",
        "mask_tool.web", "mask_tool.web.app",
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
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ICON) if ICON.is_file() else None,
    version=str(ASSETS / "version.txt") if (ASSETS / "version.txt").is_file() else None,
)
