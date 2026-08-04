"""用户偏好持久化（纯本地，仅存界面选项，不含任何文档内容）。"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

APP_NAME = "MaskToolGUI"

DEFAULTS: dict[str, Any] = {
    "mode": "smart",
    "sensitivity": "medium",      # smart/aggressive 检测灵敏度：high/medium/low/minimal
    "output_mode": "source",      # source = 原文件目录 / custom = 指定目录
    "custom_output": "",
    "save_mapping": True,
    "suffix_tag": "_脱敏",
    "mask_tool_path": "",         # 手动指定的 mask-tool 可执行文件
    "last_dir": "",
    "user_lexicon": {},           # 用户自定义敏感词词库：{类别: [词...]}
    # —— 识别质量控制 ——
    # 全局置信度下限：低于此值的识别结果只在报告里提示，不替换正文。
    # 这是凌驾于「检测灵敏度」之上的硬闸门，对词库/正则/NER 全部生效。
    "min_confidence": 0.8,
    # NER 后端：auto（有 spaCy 模型就用，否则 jieba）/ spacy / jieba
    "ner_backend": "auto",
    # spaCy 模型路径或包名；留空则自动发现（优先 程序根目录/models/）
    "spacy_model": "",
}

#: 数值型配置的合法区间，读取时做钳位
_NUMERIC_RANGE: dict[str, tuple[float, float]] = {
    "min_confidence": (0.0, 1.0),
}

#: 枚举型配置的合法取值，非法时回退默认
_ENUM_VALUES: dict[str, tuple[str, ...]] = {
    "mode": ("strict", "smart", "aggressive"),
    "sensitivity": ("high", "medium", "low", "minimal"),
    "output_mode": ("source", "custom"),
    "ner_backend": ("auto", "spacy", "jieba"),
}


def config_dir() -> Path:
    """返回配置目录（跟随各平台惯例）。"""
    if os.name == "nt":
        base = os.environ.get("APPDATA") or str(Path.home())
        return Path(base) / APP_NAME
    return Path.home() / ".config" / APP_NAME.lower()


def config_file() -> Path:
    return config_dir() / "settings.json"


def load() -> dict[str, Any]:
    """读取配置；任何异常都退回默认值，绝不因配置损坏而启动失败。"""
    data = dict(DEFAULTS)
    try:
        f = config_file()
        if f.is_file():
            loaded = json.loads(f.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                for k in DEFAULTS:
                    if k not in loaded:
                        continue
                    val = loaded[k]
                    if k in _NUMERIC_RANGE:
                        # 数值项：接受 int/float，超界钳位，非数值忽略
                        if isinstance(val, bool) or not isinstance(val, (int, float)):
                            continue
                        lo, hi = _NUMERIC_RANGE[k]
                        data[k] = min(max(float(val), lo), hi)
                    elif isinstance(val, type(DEFAULTS[k])):
                        data[k] = val
    except (OSError, ValueError, TypeError):
        pass

    # 枚举项：非法取值一律回退默认，避免下游拿到无效模式
    for key, allowed in _ENUM_VALUES.items():
        if data.get(key) not in allowed:
            data[key] = DEFAULTS[key]

    # 用户词库：清洗为 {str: [str,...]}，剔除空值与非法值
    raw_lex = data.get("user_lexicon")
    if isinstance(raw_lex, dict):
        clean: dict[str, list[str]] = {}
        for cat, words in raw_lex.items():
            if isinstance(cat, str) and isinstance(words, list):
                vals = [w for w in words if isinstance(w, str) and w.strip()]
                if vals:
                    clean[cat] = vals
        data["user_lexicon"] = clean
    else:
        data["user_lexicon"] = {}
    return data


def save(data: dict[str, Any]) -> None:
    """写入配置；失败静默忽略（不打扰用户）。"""
    try:
        d = config_dir()
        d.mkdir(parents=True, exist_ok=True)
        payload = {k: data.get(k, DEFAULTS[k]) for k in DEFAULTS}
        config_file().write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except (OSError, ValueError, TypeError):
        pass
