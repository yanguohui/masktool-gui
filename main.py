"""程序入口。PyInstaller 以此文件为打包起点。"""

from __future__ import annotations

import sys


def _fatal(msg: str) -> None:
    """在 GUI 尚未起来时也要给出人话提示，而不是闪退。"""
    try:
        import tkinter
        from tkinter import messagebox
        r = tkinter.Tk()
        r.withdraw()
        messagebox.showerror("启动失败", msg)
        r.destroy()
    except Exception:
        print(msg, file=sys.stderr)
    sys.exit(1)


def main() -> None:
    # 隐藏的自检模式：供构建后验证 / CI 使用，不影响正常 GUI 使用
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        _selftest()
        return

    if sys.version_info < (3, 10):
        _fatal("本程序需要 Python 3.10 或更高版本。")
    try:
        from app.ui import run
    except ImportError as exc:
        _fatal(f"程序组件缺失，无法启动。\n\n技术信息：{exc}")
        return
    run()


def _selftest() -> None:
    """命令行自检：``脱敏工具 --selftest <输入文件> [输出目录]``。"""
    import json
    import tempfile
    from pathlib import Path

    if len(sys.argv) < 3:
        print(json.dumps({"ok": False, "error": "用法: --selftest <输入文件> [输出目录]"},
                         ensure_ascii=False))
        sys.exit(2)
    src = Path(sys.argv[2])
    out_dir = Path(sys.argv[3]) if len(sys.argv) > 3 else Path(tempfile.mkdtemp())
    try:
        import app.engine as E
        eng = E.MaskEngine(E.locate_mask_tool())
        r = eng.process_one(src, out_dir, mode="smart", save_mapping=True)
        print(json.dumps({
            "frozen": E.IS_FROZEN,
            "tool": E.locate_mask_tool().display,
            "ok": r.ok,
            "output": str(r.output) if r.output else None,
            "masked_count": r.masked_count,
            "message": r.message,
        }, ensure_ascii=False))
        sys.exit(0 if r.ok else 1)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
