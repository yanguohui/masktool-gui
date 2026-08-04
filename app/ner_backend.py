"""
可插拔 NER 后端
===============

mask-tool 自带的 NER 是基于 jieba 词性标注的轻量实现，优点是零依赖、
体积小，缺点有两个：

1. **未登录词能力弱** —— 通信工程文档里的专有名词（网元、承载网、
   智慧运维管控平台…）不在 jieba 词典里，词性标注常给出 ``n`` 而非
   ``nt``/``nz``，实体直接漏掉；
2. **置信度天花板低** —— 其 ``_calc_confidence`` 上限约 0.75，在本项目
   0.8 的全局置信度下限之下，NER 结果几乎无法单独触发替换。

本模块提供一个 **spaCy 后端**，并遵循三条原则：

- **可选依赖**：未安装 spaCy 或未提供模型时自动回退 jieba，主流程不受影响；
- **同一接口**：实现 mask-tool 的 ``BaseNER`` 协议，Pipeline 无感知；
- **构词法复核**：spaCy 负责「提出候选」，本项目既有的构词法校验函数负责
  「确认或否决」。两者一致时才给出高置信度，从而既拿到深度模型的召回，
  又不把 ``二次开发``、``卖方`` 这类业务术语放进来。

模型来源（按优先级）
--------------------
1. 配置项 / 环境变量 ``MASKTOOL_SPACY_MODEL`` 指定的路径或包名；
2. 程序根目录下的 ``models/`` 目录中的 spaCy 模型目录
   （含 ``config.cfg`` 与 ``meta.json``），名字含 ``telecom`` / ``通信`` /
   ``comm`` 的优先 —— 这是放置**领域微调模型**的约定位置；
3. 已安装的通用中文模型 ``zh_core_web_lg/md/sm``。

领域微调模型可用 ``tools/train_spacy_ner.py`` 训练，详见 README。
"""

from __future__ import annotations

import importlib
import os
from pathlib import Path
from typing import Any, Iterable

__all__ = [
    "SPACY_LABEL_MAP",
    "discover_spacy_model",
    "spacy_available",
    "create_ner_engine",
    "backend_status",
]


# --------------------------------------------------------------------------
# spaCy 实体标签 → mask-tool 实体类别
# --------------------------------------------------------------------------
# 取值参考 OntoNotes 5 标注体系（zh_core_web_* 系列即基于此训练）。
# 刻意**不映射** DATE / TIME / CARDINAL / ORDINAL / PERCENT / QUANTITY：
# 这些在技术规范书里绝大多数是「3 个月质保」「第 2 标段」这类工程描述，
# 映射过来只会制造误报；真正需要的日期/金额由确定性正则负责。
SPACY_LABEL_MAP: dict[str, str] = {
    "PERSON": "person",
    "PER": "person",          # 部分社区模型用 PER
    "ORG": "company",
    "COMPANY": "company",     # 微调模型可自定义此标签
    "GOV": "government",
    "GPE": "location",
    "LOC": "location",
    "FAC": "location",
    "ADDRESS": "location",
    "MONEY": "amount",
    "AMOUNT": "amount",
    "PROJECT": "project",     # 微调模型可自定义此标签
    "WORK_OF_ART": "project",
    "EVENT": "project",
}

#: 各标签的基线置信度（未经构词法复核时）
_BASE_CONF: dict[str, float] = {
    "person": 0.70,
    "company": 0.72,
    "government": 0.72,
    "location": 0.66,
    "amount": 0.80,
    "project": 0.68,
    "custom": 0.60,
}

#: 构词法复核通过时的加成 / 否决时的扣减
_AGREE_BONUS = 0.18
_REJECT_PENALTY = 0.35

#: 领域微调模型的命名约定（放在 models/ 下，命中即优先加载）
_DOMAIN_HINTS = ("telecom", "通信", "comm", "domain", "finetune", "ft")

#: 兜底的通用中文模型（按效果从好到差）
_GENERIC_MODELS = ("zh_core_web_trf", "zh_core_web_lg",
                   "zh_core_web_md", "zh_core_web_sm")

ENV_MODEL = "MASKTOOL_SPACY_MODEL"


# --------------------------------------------------------------------------
# 可用性探测
# --------------------------------------------------------------------------


def _import_spacy():
    """动态导入 spaCy。

    刻意使用 ``importlib`` 而非 ``import spacy``：PyInstaller 的静态分析
    看不到动态导入，因此默认打包出的 exe 不会把 spaCy（及其 torch/thinc
    等数百 MB 依赖）卷进去。需要 spaCy 版 exe 时在 spec 里显式声明即可。
    """
    try:
        return importlib.import_module("spacy")
    except Exception:
        return None


def spacy_available() -> bool:
    return _import_spacy() is not None


def _is_model_dir(p: Path) -> bool:
    """判断一个目录是否是 spaCy 模型目录。"""
    try:
        return p.is_dir() and (p / "config.cfg").is_file() and (p / "meta.json").is_file()
    except OSError:
        return False


def _iter_model_dirs(root: Path) -> Iterable[Path]:
    """在 models/ 下寻找模型目录（支持一层 pip 包式嵌套）。"""
    try:
        entries = sorted(root.iterdir())
    except OSError:
        return
    for d in entries:
        if _is_model_dir(d):
            yield d
            continue
        # pip 安装的模型包形如 zh_core_web_sm/zh_core_web_sm-3.7.0/
        if d.is_dir():
            try:
                for sub in sorted(d.iterdir()):
                    if _is_model_dir(sub):
                        yield sub
                        break
            except OSError:
                continue


def discover_spacy_model(explicit: str | None = None,
                         root: Path | None = None) -> str | None:
    """按优先级找出要加载的 spaCy 模型（路径或包名）；找不到返回 None。"""
    # ① 显式配置
    for cand in (explicit, os.environ.get(ENV_MODEL)):
        if cand and str(cand).strip():
            return str(cand).strip()

    # ② 程序根目录 models/ 下的模型目录（领域微调模型的约定位置）
    if root is not None:
        models_dir = root / "models"
        found = list(_iter_model_dirs(models_dir))
        if found:
            for d in found:
                if any(h in d.name.lower() for h in _DOMAIN_HINTS):
                    return str(d)
            return str(found[0])

    # ③ 已安装的通用中文模型
    spacy = _import_spacy()
    if spacy is not None:
        for name in _GENERIC_MODELS:
            try:
                importlib.import_module(name)
                return name
            except Exception:
                continue
    return None


# --------------------------------------------------------------------------
# spaCy NER 引擎
# --------------------------------------------------------------------------


class SpacyNER:
    """spaCy 实现的 NER 引擎（鸭子类型兼容 mask-tool 的 ``BaseNER``）。

    与 jieba 后端的关键差异：**每条候选都要经过构词法复核**。
    spaCy 说「这是机构名」只是投票之一，最终置信度由
    「模型基线 ± 构词法结论」共同决定，因此：

    - 真机构（苏州赛微电子科技股份有限公司）→ 0.72 + 0.18 = 0.90，过闸；
    - 业务术语（卖方 / 二次开发）→ 先被白名单拦掉，即使没拦住，
      构词法否决后降到 0.37，也进不了 0.8 的门。
    """

    MIN_LENGTH = 2

    def __init__(self, model: str | None = None, root: Path | None = None) -> None:
        self._whitelist: set[str] = set()
        self._nlp: Any = None
        self._model_name: str = ""
        self._error: str = ""
        self._load(model, root)

    # -- 加载 -------------------------------------------------------------

    def _load(self, model: str | None, root: Path | None) -> None:
        spacy = _import_spacy()
        if spacy is None:
            self._error = "未安装 spaCy（pip install spacy）"
            return
        name = discover_spacy_model(model, root)
        if not name:
            self._error = "未找到可用的 spaCy 模型"
            return
        try:
            # 只保留 NER 需要的组件，速度快很多
            self._nlp = spacy.load(name, exclude=["lemmatizer", "textcat"])
            self._model_name = name
        except Exception as exc:  # 模型损坏 / 版本不匹配
            self._error = f"spaCy 模型加载失败：{exc}"
            self._nlp = None

    # -- BaseNER 协议 -----------------------------------------------------

    def set_whitelist(self, whitelist: set[str]) -> None:
        self._whitelist = set(whitelist or ())

    def is_available(self) -> bool:
        return self._nlp is not None

    @property
    def model_name(self) -> str:
        return self._model_name

    @property
    def error(self) -> str:
        return self._error

    def recognize(self, text: str, file_path: str = "") -> list:
        if self._nlp is None or not (text or "").strip():
            return []

        from mask_tool.models.detection import (  # 延迟导入，避免循环依赖
            DetectionResult, DetectionType, Location,
        )

        try:
            doc = self._nlp(text)
        except Exception:
            return []

        type_by_name = {t.value: t for t in DetectionType}
        results: list = []
        seen: set[str] = set()

        for ent in getattr(doc, "ents", ()):
            word = (ent.text or "").strip()
            if len(word) < self.MIN_LENGTH or word in seen:
                continue
            if word in self._whitelist or self._is_whitelisted(word):
                continue

            cat = SPACY_LABEL_MAP.get(ent.label_)
            if cat is None:
                continue
            det_type = type_by_name.get(cat)
            if det_type is None:
                continue

            conf = self._score(word, cat)

            start = max(0, ent.start_char - 50)
            end = min(len(text), ent.end_char + 50)
            results.append(DetectionResult(
                text=word,
                text_type=det_type,
                source="ner",
                confidence=round(conf, 3),
                location=Location(file=file_path),
                context=text[start:end],
            ))
            seen.add(word)

        return results

    # -- 打分 -------------------------------------------------------------

    @staticmethod
    def _is_whitelisted(word: str) -> bool:
        """复用引擎的 whitelist.txt（含「本 / 该」前缀归一化）。"""
        try:
            from app.engine import is_whitelisted
        except Exception:
            return False
        try:
            return bool(is_whitelisted(word))
        except Exception:
            return False

    def _score(self, word: str, cat: str) -> float:
        """模型基线 + 构词法复核。"""
        conf = _BASE_CONF.get(cat, 0.60)

        # 长实体更可信（专名通常较长）
        if len(word) >= 8:
            conf += 0.06
        elif len(word) >= 5:
            conf += 0.03

        verdict = self._structural_verdict(word, cat)
        if verdict is True:
            conf += _AGREE_BONUS
        elif verdict is False:
            conf -= _REJECT_PENALTY

        return max(0.0, min(conf, 0.99))

    @staticmethod
    def _structural_verdict(word: str, cat: str) -> bool | None:
        """用本项目的构词法校验函数复核；无对应校验器时返回 None。"""
        try:
            from app import engine as E
        except Exception:
            return None
        fn = {
            "company": getattr(E, "_looks_like_org", None),
            "government": getattr(E, "_looks_like_org", None),
            "person": getattr(E, "_looks_like_person", None),
            "location": getattr(E, "_looks_like_geo", None),
            "project": getattr(E, "_looks_like_project", None),
        }.get(cat)
        if fn is None:
            return None
        try:
            return bool(fn(word))
        except Exception:
            return None


# --------------------------------------------------------------------------
# 工厂
# --------------------------------------------------------------------------


def create_ner_engine(backend: str = "auto",
                      model: str | None = None,
                      root: Path | None = None):
    """按配置创建 NER 引擎。

    Args:
        backend: ``auto``（有 spaCy 模型就用，否则 jieba）/ ``spacy`` / ``jieba``
        model:   spaCy 模型路径或包名；``None`` 则自动发现
        root:    程序根目录（用于定位 ``models/``）

    Returns:
        NER 引擎实例；``None`` 表示「让上游按原逻辑用 jieba」。
    """
    b = (backend or "auto").strip().lower()
    if b == "jieba":
        return None
    if b not in ("auto", "spacy"):
        b = "auto"

    if b == "auto" and discover_spacy_model(model, root) is None:
        return None

    eng = SpacyNER(model, root)
    if eng.is_available():
        return eng
    # spacy 明确指定但加载失败 → 回退 jieba，不让整个流程挂掉
    return None


def backend_status(backend: str = "auto",
                   model: str | None = None,
                   root: Path | None = None) -> dict:
    """返回后端状态，供 UI / 自检展示（不加载模型，尽量轻量）。"""
    has_spacy = spacy_available()
    found = discover_spacy_model(model, root) if has_spacy else None
    b = (backend or "auto").strip().lower()
    if b == "jieba":
        active, reason = "jieba", "配置指定使用 jieba"
    elif not has_spacy:
        active, reason = "jieba", "未安装 spaCy，已回退 jieba"
    elif not found:
        active, reason = "jieba", "未找到 spaCy 模型，已回退 jieba"
    else:
        active, reason = "spacy", f"使用 spaCy 模型：{found}"
    return {
        "configured": b,
        "active": active,
        "spacy_installed": has_spacy,
        "model": found or "",
        "reason": reason,
    }
