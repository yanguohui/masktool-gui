"""
脱敏引擎封装层
==============

职责：定位 mask-tool 命令行工具、以子进程方式调用它、把产物搬运到用户指定位置。

设计要点（都是踩过的坑）：
1. mask-tool 的 ``--output`` 是**目录**而不是文件路径，且产物命名固定为
   ``{stem}_masked.docx`` / ``{stem}_masked_report.json``，还会额外写出
   ``mapping.json`` 与 ``report.json``。因此这里统一先输出到临时目录，
   再按用户要求重命名搬运，避免污染用户目录。
2. mask-tool 对单个文件的异常是 **内部 catch 掉** 的，进程退出码仍然是 0。
   所以不能只看 returncode，必须以"预期产物是否存在"作为成功判据。
3. Windows 下 PyInstaller 的 windowed 程序调用子进程会弹黑色控制台窗口，
   必须用 CREATE_NO_WINDOW + STARTUPINFO 抑制。
4. mask-tool 输出大量中文与 rich 富文本，Windows 控制台默认 GBK 会炸编码，
   必须强制 PYTHONUTF8 / PYTHONIOENCODING 并以 utf-8+replace 解码。
5. **打包模式（frozen）下，mask-tool 被内嵌进 exe**，此时无法再用
   ``python -m mask_tool`` 子进程方式（PyInstaller 冻结后 sys.executable
   不是解释器）。改为**进程内调用** ``Pipeline.process_file``，既自包含又
   避免了子进程编码 / 黑窗 / 找不到命令等一系列问题。
"""

from __future__ import annotations

import atexit
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Sequence

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------

IS_WINDOWS = os.name == "nt"

#: PyInstaller 冻结后该属性为真，程序以单文件 exe 形式运行
IS_FROZEN = bool(getattr(sys, "frozen", False))

#: mask-tool 支持的全部格式
SUPPORTED_EXTS: tuple[str, ...] = (".docx", ".pdf", ".xlsx", ".pptx")

#: 这些格式能真正产出"脱敏后的同类型文件"
REWRITABLE_EXTS: tuple[str, ...] = (".docx", ".xlsx", ".pptx")

#: PDF 在 mask-tool 当前版本只输出检测报告（JSON），不回写 PDF 本体
REPORT_ONLY_EXTS: tuple[str, ...] = (".pdf",)

MODES: tuple[str, ...] = ("strict", "smart", "aggressive")


# --------------------------------------------------------------------------
# 上游缺陷修复：mask-tool 的 DocxAdapter 在“段落内多实体替换”时，
# 位置映射算法会丢弃替换区间内的原始下标，导致多个人名/地名同段时
# 出现 `[PE[PERSON_xxx]ON]` 之类的嵌套乱码。这里在打包后进程内调用前
# 打一个补丁，用“按原字符下标重建段落”的正确实现替换之。
# --------------------------------------------------------------------------

def _patched_docx_replace(self, paragraph, replace_map: dict) -> None:
    """正确、无串扰的段落级替换（修复上游 DocxAdapter 的实现）。"""
    if not paragraph.runs or not replace_map:
        return
    runs = paragraph.runs
    full_text = "".join(r.text for r in runs)
    n = len(full_text)
    if n == 0:
        return

    # 1) 计算非重叠的替换区间（左到右，遇到重叠跳过）
    occupied = [False] * n
    replacements = []
    for original, token in replace_map.items():
        if not original:
            continue
        start = 0
        while True:
            idx = full_text.find(original, start)
            if idx == -1:
                break
            end = idx + len(original)
            if not any(occupied[idx:end]):
                replacements.append((idx, end, token))
                for i in range(idx, end):
                    occupied[i] = True
            start = end
    if not replacements:
        return
    replacements.sort(key=lambda x: x[0])

    # 2) 重建新文本，并为每个新字符记录“归属的原始下标”
    #    插入的 token 字符归属到其替换起点（必落在某 run 内），从而能正确回填。
    new_chars: list[str] = []
    owners: list[int] = []
    pos = 0
    for (s, e, tok) in replacements:
        while pos < s:
            new_chars.append(full_text[pos]); owners.append(pos); pos += 1
        for c in tok:
            new_chars.append(c); owners.append(s)
        pos = e
    while pos < n:
        new_chars.append(full_text[pos]); owners.append(pos); pos += 1

    # 3) 计算各 run 在原文本中的区间，并把新字符按归属下标回填
    run_spans: list[tuple[int, int, int]] = []
    p = 0
    for i, r in enumerate(runs):
        run_spans.append((i, p, p + len(r.text)))
        p += len(r.text)

    buckets: dict[int, list[str]] = {i: [] for i in range(len(runs))}
    for ch, owner in zip(new_chars, owners):
        for (ri, rs, re) in run_spans:
            if rs <= owner < re:
                buckets[ri].append(ch)
                break
    for i, r in enumerate(runs):
        r.text = "".join(buckets[i])


def _patch_mask_tool_adapters() -> None:
    """对 mask-tool 的 docx 适配器打替换补丁（幂等）。"""
    try:
        from mask_tool.adapters.docx_adapter import DocxAdapter
        if not getattr(DocxAdapter, "_patched_by_engine", False):
            DocxAdapter._replace_all_in_paragraph = _patched_docx_replace  # type: ignore[assignment]
            DocxAdapter._patched_by_engine = True  # type: ignore[attr-defined]
    except Exception:
        pass


def _bundled_config_dir() -> Path | None:
    """返回 mask-tool 的 config/ 资源目录。

    - 冻结模式：位于 PyInstaller 解压目录 ``_MEIPASS/mask_tool_config``。
    - 开发模式：回退到项目自身的 ``assets/mask_tool_config``，确保两端行为一致。
    """
    if IS_FROZEN:
        base = Path(getattr(sys, "_MEIPASS", "."))
        cand = base / "mask_tool_config"
        if cand.is_dir():
            return cand
    dev = Path(__file__).resolve().parent.parent / "assets" / "mask_tool_config"
    if dev.is_dir():
        return dev
    return None

MODE_LABELS: dict[str, str] = {
    "strict": "strict（严格）— 仅使用用户词库，误伤最少",
    "smart": "smart（智能）— 词库 + 正则 + NER，推荐",
    "aggressive": "aggressive（激进）— 高召回，宁可错杀",
}

# --------------------------------------------------------------------------
# smart 模式的「检测灵敏度」（仅 smart 模式可见/生效）
# --------------------------------------------------------------------------
# 设计要点：
#   · 词库匹配（lexicon）、正则（手机/身份证/邮箱/金额/IP/MAC/日期/中文金额/
#     项目工程名）属于高置信信息，无论灵敏度如何都照常脱敏（见 _patch_policy）。
#   · jieba NER 的置信度上限仅 0.85，且「人名/机构/地名」的置信度分布差异很大：
#       机构名（company）几乎都在 0.83 左右，可靠；
#       人名（person）多在 0.60~0.70，偏低的真名与偏高的误报并存；
#       地名（location）误报极多（"大海""新建"等都被误标），不可信；
#       其他专名（custom，nz）鱼龙混杂。
#     因此「灵敏度」按实体类型分别设阈值，而非一刀切：
#       - company 阈值最低（最可靠），person 次之，custom 较严，location 最严。
#   · 同时用一份「常见误报白名单」把 jieba 常误标为实体的非敏感词
#     （严格遵守/中华民族/系统安全/新建…）排除在脱敏之外，
#     这样即便调低 person/custom 阈值，也不会重新误伤这些普通词。
#   · 项目名、客户名若未被 NER 命中，最稳妥的方式仍是加入「用户词库」。

# 每个档位给出「按实体类型」的 NER 阈值（confidence >= 阈值才脱敏）。
# aggressive 模式等价于 high 档（尽可能高召回）。
SENSITIVITY_LEVELS: dict[str, dict] = {
    "high": {
        "label": "高（激进召回）",
        "thr": {"company": 0.55, "person": 0.55, "custom": 0.55, "location": 0.55},
        "desc": (
            "几乎不漏掉任何疑似信息：人名 / 机构 / 地名等智能识别项全部参与脱敏。"
            "代价是误伤最多，普通词语也常被误判为敏感。仅在宁可错杀、事后人工复核的场景使用。"
        ),
    },
    "medium": {
        "label": "中（推荐）",
        "thr": {"company": 0.60, "person": 0.64, "custom": 0.68, "location": 0.78},
        "desc": (
            "平衡之选：手机、身份证、邮箱、金额、IP、MAC、日期、中文金额、项目/工程名等"
            "高置信信息照常脱敏；机构名、人名（含客户/联系人姓名）基本都能命中；"
            "仅压制低置信的地名误报。常见非敏感词（严格遵守/中华民族/系统安全/新建…）已加入白名单不再误伤。"
        ),
    },
    "low": {
        "label": "低（保守精准）",
        "thr": {"company": 0.75, "person": 0.72, "custom": 0.78, "location": 0.85},
        "desc": (
            "只脱敏置信度较高的智能识别项（多为真实机构名、长专有名词），几乎不误伤普通词语；"
            "短人名、部分客户名可能漏掉，建议配合「用户词库」或 strict 模式。"
        ),
    },
    "minimal": {
        "label": "极低（审慎）",
        "thr": {"company": 0.85, "person": 0.82, "custom": 0.88, "location": 0.90},
        "desc": (
            "仅处理最高置信的少量项，误伤最少；适合对误报零容忍、宁可漏掉的场合。"
            "智能识别基本只保留用户词库匹配项。"
        ),
    },
}
SENSITIVITY_KEYS: tuple[str, ...] = ("high", "medium", "low", "minimal")
SENSITIVITY_DEFAULT = "medium"

#: aggressive 模式使用与 high 相同的「尽可能高召回」阈值
_AGGRESSIVE_THR: dict[str, float] = dict(SENSITIVITY_LEVELS["high"]["thr"])

#: 正则类（高精度）始终脱敏的下限，与灵敏度无关
_REGEX_AUTO = 0.70
_REGEX_SUGGEST = 0.55

#: 当前生效的 smart 灵敏度阈值（由 UI 在每次处理前设置）
_SMART_THR: dict[str, float] = dict(SENSITIVITY_LEVELS[SENSITIVITY_DEFAULT]["thr"])


def set_smart_sensitivity(key: str) -> None:
    """切换当前生效的 smart 检测灵敏度（影响后续 smart 处理）。"""
    global _SMART_THR
    lvl = SENSITIVITY_LEVELS.get(key, SENSITIVITY_LEVELS[SENSITIVITY_DEFAULT])
    _SMART_THR = dict(lvl["thr"])


def _ner_threshold(text_type) -> float:
    """把 DetectionType 映射到当前灵敏度下的阈值。"""
    name = getattr(text_type, "value", str(text_type)).lower()
    if name in ("company", "government"):
        return _SMART_THR["company"]
    if name == "person":
        return _SMART_THR["person"]
    if name == "location":
        return _SMART_THR["location"]
    return _SMART_THR["custom"]


def _patch_detector_rules() -> None:
    """给 mask-tool 的检测器补充高置信正则（幂等）。

    解决两类原版漏检：
      1) 阿拉伯数字金额必须带「万/亿」才被识别（"1,200,000元" 漏检）；
         新增「纯数字+元」规则。
      2) 中文大写金额（陆拾万元/壹佰万元）与「X项目/X工程」这类项目名
         原版完全不识别；新增对应正则，作为高置信项始终脱敏。
    """
    try:
        from mask_tool.core.detector import Detector
        from mask_tool.models.detection import DetectionType

        if getattr(Detector, "_patched_rules_by_engine", False):
            return

        _orig = Detector._build_regex_rules  # type: ignore[assignment]

        def _build(self):  # noqa: ANN001
            rules = _orig(self)
            extra = [
                # 阿拉伯数字金额（含「元」但不带万/亿，如 1,200,000元 / 500元）
                (re.compile(r"[\d,]+\.?\d*元"), DetectionType.AMOUNT, 0.80),
                # 中文大写金额：陆拾万元 / 壹佰万元 / 伍仟元 / 十亿
                # 注意：前导数字不含「万/亿」，且整体须以 元/圆/万/亿 收尾，
                # 避免把「一个」「十」之类的普通词误判为金额。
                (re.compile(r"[零〇一二三四五六七八九十百千壹贰叁肆伍陆柒捌玖拾佰仟]+(?:[万亿]?元|[万亿]?圆|万|亿)"),
                 DetectionType.AMOUNT, 0.85),
                # 项目 / 工程名（以「项目」「工程」结尾的专有名词短语）
                (re.compile(r"[一-鿿]{2,}(?:项目|工程)"), DetectionType.PROJECT, 0.85),
            ]
            return rules + extra

        Detector._build_regex_rules = _build  # type: ignore[assignment]
        Detector._patched_rules_by_engine = True  # type: ignore[attr-defined]
    except Exception:
        pass


def _patch_policy() -> None:
    """对 mask-tool 的 PolicyEngine 打补丁（幂等）。

    仅对 smart / aggressive 模式生效：
      · 词库与正则（含本项目补充的高置信正则）始终脱敏；
      · NER 项按「实体类型 × 用户灵敏度」阈值裁剪，机构名/人名更宽松、
        地名更严格，从而在尽量少误伤的前提下覆盖真实的客户名/项目名/金额。
    """
    try:
        from mask_tool.core.policy import PolicyEngine
        from mask_tool.models.detection import DetectionStatus

        if getattr(PolicyEngine, "_patched_by_engine", False):
            return

        _orig_decide = PolicyEngine._decide  # type: ignore[assignment]

        def _decide(self, result):  # noqa: ANN001
            mode = self.config.mode
            if mode in ("smart", "aggressive"):
                conf = result.confidence
                src = getattr(result, "source", "") or ""
                if src == "dictionary":
                    # 用户词库 / 基准词库：永远脱敏
                    return DetectionStatus.AUTO_MASK if conf >= 0.95 \
                        else DetectionStatus.HINT_ONLY
                if src == "regex":
                    # 高置信正则（手机/身份证/邮箱/金额/IP/MAC/日期/中文金额/项目名）始终脱敏
                    if conf >= _REGEX_AUTO:
                        return DetectionStatus.AUTO_MASK
                    if conf >= _REGEX_SUGGEST:
                        return DetectionStatus.SUGGEST_MASK
                    return DetectionStatus.HINT_ONLY
                # NER（及未知来源）：按实体类型 × 灵敏度裁剪
                thr = _AGGRESSIVE_THR if mode == "aggressive" else _SMART_THR
                name = getattr(result.text_type, "value", str(result.text_type)).lower()
                if name in ("company", "government"):
                    t = thr["company"]
                elif name == "person":
                    t = thr["person"]
                elif name == "location":
                    t = thr["location"]
                else:
                    t = thr["custom"]
                return DetectionStatus.AUTO_MASK if conf >= t else DetectionStatus.HINT_ONLY
            # strict / focused 等模式沿用上游原始逻辑
            return _orig_decide(self, result)

        PolicyEngine._decide = _decide  # type: ignore[assignment]
        PolicyEngine._patched_by_engine = True  # type: ignore[attr-defined]
    except Exception:
        pass


#: 单个文件处理超时（秒）
PER_FILE_TIMEOUT = 600

_CREATE_NO_WINDOW = 0x08000000


# --------------------------------------------------------------------------
# 异常
# --------------------------------------------------------------------------


class MaskToolNotFound(Exception):
    """未能定位到 mask-tool 命令行工具。"""


class MaskToolCallError(Exception):
    """调用 mask-tool 过程中出现无法恢复的错误。"""


# --------------------------------------------------------------------------
# 数据结构
# --------------------------------------------------------------------------


@dataclass
class FileResult:
    """单个文件的处理结果。"""

    source: Path
    ok: bool
    output: Path | None = None
    mapping: Path | None = None
    masked_count: int = 0
    message: str = ""
    report_only: bool = False

    @property
    def status_text(self) -> str:
        if not self.ok:
            return "失败"
        return "仅报告" if self.report_only else "成功"


@dataclass
class ToolInfo:
    """已定位到的 mask-tool 调用方式。"""

    argv: list[str]
    origin: str
    version: str = ""

    @property
    def display(self) -> str:
        base = " ".join(self.argv)
        return f"{base}  （{self.origin}）"


@dataclass
class BatchOutcome:
    """一次批量处理的整体结果。"""

    results: list[FileResult] = field(default_factory=list)
    cancelled: bool = False

    @property
    def success_count(self) -> int:
        return sum(1 for r in self.results if r.ok)

    @property
    def failure_count(self) -> int:
        return sum(1 for r in self.results if not r.ok)


# --------------------------------------------------------------------------
# 子进程工具
# --------------------------------------------------------------------------


def _subprocess_kwargs() -> dict:
    """构造跨平台的 subprocess 参数，Windows 下隐藏控制台窗口。"""
    env = os.environ.copy()
    # 强制 UTF-8，规避 Windows GBK 控制台把中文输出打成乱码/抛异常
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    # 关闭 rich 的颜色转义，便于解析文本
    env["NO_COLOR"] = "1"
    env["TERM"] = "dumb"

    kwargs: dict = {
        "env": env,
        "stdout": subprocess.PIPE,
        "stderr": subprocess.STDOUT,
        "encoding": "utf-8",
        "errors": "replace",
    }

    if IS_WINDOWS:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        kwargs["startupinfo"] = startupinfo
        kwargs["creationflags"] = _CREATE_NO_WINDOW

    return kwargs


def _run(argv: Sequence[str], timeout: int = 60, cwd: Path | None = None):
    """执行命令并返回 CompletedProcess，不抛 CalledProcessError。"""
    return subprocess.run(
        list(argv),
        timeout=timeout,
        cwd=str(cwd) if cwd else None,
        check=False,
        **_subprocess_kwargs(),
    )


# --------------------------------------------------------------------------
# mask-tool 定位
# --------------------------------------------------------------------------


def _candidate_script_dirs() -> list[Path]:
    """收集可能存放 mask-tool 可执行文件的目录。"""
    dirs: list[Path] = []
    bin_name = "Scripts" if IS_WINDOWS else "bin"

    # 当前解释器所在环境
    dirs.append(Path(sys.prefix) / bin_name)
    if hasattr(sys, "base_prefix"):
        dirs.append(Path(sys.base_prefix) / bin_name)

    if IS_WINDOWS:
        local = os.environ.get("LOCALAPPDATA")
        if local:
            base = Path(local) / "Programs" / "Python"
            if base.is_dir():
                for child in base.iterdir():
                    if child.is_dir():
                        dirs.append(child / "Scripts")
            # pip install --user 的落点
            roaming = os.environ.get("APPDATA")
            if roaming:
                py_dir = Path(roaming) / "Python"
                if py_dir.is_dir():
                    for child in py_dir.iterdir():
                        if child.is_dir():
                            dirs.append(child / "Scripts")
    else:
        dirs.append(Path.home() / ".local" / "bin")
        dirs.append(Path("/usr/local/bin"))
        dirs.append(Path("/opt/homebrew/bin"))

    # 去重且保持顺序
    seen: set[str] = set()
    uniq: list[Path] = []
    for d in dirs:
        key = str(d)
        if key not in seen:
            seen.add(key)
            uniq.append(d)
    return uniq


def _python_candidates() -> list[list[str]]:
    """可用于 ``-m mask_tool`` 的解释器候选。

    注意：PyInstaller 打包后 sys.executable 指向 exe 自身，绝不能拿来当解释器用。
    """
    cands: list[list[str]] = []
    frozen = getattr(sys, "frozen", False)

    if not frozen:
        cands.append([sys.executable])

    for name in ("python", "python3"):
        found = shutil.which(name)
        if found:
            cands.append([found])

    if IS_WINDOWS:
        py_launcher = shutil.which("py")
        if py_launcher:
            cands.append([py_launcher, "-3"])

    return cands


def _verify(argv: list[str]) -> str | None:
    """验证候选命令确实是可用的 mask-tool，返回版本号文本。"""
    try:
        proc = _run([*argv, "--help"], timeout=45)
    except (OSError, subprocess.SubprocessError):
        return None

    out = (proc.stdout or "")
    if proc.returncode == 0 and "mask" in out.lower():
        return _probe_version(argv)
    return None


def _probe_version(argv: list[str]) -> str:
    """尽力取出版本号；失败则返回空串（不影响主流程）。"""
    try:
        proc = _run([*argv, "version"], timeout=45)
        text = (proc.stdout or "").strip()
        if proc.returncode == 0 and text:
            for line in text.splitlines():
                line = line.strip()
                if line:
                    return line[:80]
    except (OSError, subprocess.SubprocessError):
        pass
    return ""


def locate_mask_tool(manual_path: str | None = None) -> ToolInfo:
    """多级回退定位 mask-tool。

    顺序：用户手动指定 → PATH → 常见 Scripts 目录 → ``python -m mask_tool``。

    Raises:
        MaskToolNotFound: 全部候选都不可用。
    """
    # 0) 冻结（打包）模式：mask-tool 已内嵌，直接进程内调用验证可用性
    if IS_FROZEN:
        try:
            import importlib
            mt = importlib.import_module("mask_tool")
            ver = str(getattr(mt, "__version__", ""))
            return ToolInfo(["(内嵌进程调用)"], "已内嵌于程序", ver)
        except Exception:
            # 极端情况下未打包成功，继续走下面的回退（基本都会失败）
            pass

    tried: list[str] = []

    # 1) 用户手动指定
    if manual_path:
        p = Path(manual_path.strip('"').strip())
        if p.is_file():
            ver = _verify([str(p)])
            if ver is not None:
                return ToolInfo([str(p)], "手动指定", ver)
        tried.append(f"手动指定路径无效：{manual_path}")

    # 2) PATH
    for name in ("mask-tool", "mask_tool"):
        found = shutil.which(name)
        if found:
            ver = _verify([found])
            if ver is not None:
                return ToolInfo([found], "系统 PATH", ver)
            tried.append(f"PATH 中的 {found} 无法正常执行")

    # 3) 常见脚本目录
    exe_names = ("mask-tool.exe", "mask-tool") if IS_WINDOWS else ("mask-tool",)
    for d in _candidate_script_dirs():
        if not d.is_dir():
            continue
        for exe in exe_names:
            cand = d / exe
            if cand.is_file():
                ver = _verify([str(cand)])
                if ver is not None:
                    return ToolInfo([str(cand)], f"发现于 {d}", ver)

    # 4) python -m mask_tool
    for py in _python_candidates():
        argv = [*py, "-m", "mask_tool"]
        ver = _verify(argv)
        if ver is not None:
            return ToolInfo(argv, "python -m mask_tool", ver)

    detail = "\n".join(f"  · {t}" for t in tried)
    raise MaskToolNotFound(detail)


# --------------------------------------------------------------------------
# 产物解析
# --------------------------------------------------------------------------


def _expected_artifact(tmp_dir: Path, src: Path) -> Path | None:
    """按 mask-tool 的命名规则找出本次产物。"""
    stem, ext = src.stem, src.suffix.lower()

    if ext in REPORT_ONLY_EXTS:
        cand = tmp_dir / f"{stem}_masked_report.json"
        return cand if cand.is_file() else None

    cand = tmp_dir / f"{stem}_masked{ext}"
    if cand.is_file():
        return cand

    # 兜底：目录里除 mapping/report 之外的任意新文件
    for f in sorted(tmp_dir.iterdir()):
        if f.is_file() and f.name not in ("mapping.json", "report.json"):
            return f
    return None


def _read_masked_count(tmp_dir: Path, artifact: Path, report_only: bool) -> int:
    """读取本次脱敏/检出的条目数量，用于结果展示。"""
    try:
        if report_only:
            data = json.loads(artifact.read_text(encoding="utf-8"))
            return int(data.get("total_detections", 0))
        mapping = tmp_dir / "mapping.json"
        if mapping.is_file():
            data = json.loads(mapping.read_text(encoding="utf-8"))
            meta = data.get("metadata", {})
            if "total_mappings" in meta:
                return int(meta["total_mappings"])
            return len(data.get("tokens", {}))
    except (OSError, ValueError, TypeError):
        pass
    return 0


def _unique_path(target: Path) -> Path:
    """避免覆盖已有文件，自动追加 (2)(3)…"""
    if not target.exists():
        return target
    stem, suffix, parent = target.stem, target.suffix, target.parent
    for i in range(2, 1000):
        cand = parent / f"{stem} ({i}){suffix}"
        if not cand.exists():
            return cand
    raise MaskToolCallError("输出目录中同名文件过多，无法生成唯一文件名。")


def build_output_name(src: Path, suffix_tag: str = "_脱敏") -> str:
    """生成用户可见的输出文件名。"""
    ext = src.suffix.lower()
    if ext in REPORT_ONLY_EXTS:
        return f"{src.stem}{suffix_tag}_检测报告.json"
    return f"{src.stem}{suffix_tag}{src.suffix}"


# --------------------------------------------------------------------------
# 引擎
# --------------------------------------------------------------------------


class MaskEngine:
    """对 mask-tool 的高层封装。"""

    def __init__(self, tool: ToolInfo, user_lexicon: dict | None = None):
        self.tool = tool
        self._current: subprocess.Popen | None = None
        self._cancelled = False
        # 进程内调用后端：(Pipeline 类, MaskConfig 类, DetectionStatus)；不可用为 None
        self._inproc = self._init_inproc()
        # 用户词库：合并基准词库后写出临时 YAML，返回路径；空则 None
        self._user_lexicon_file = self._build_user_lexicon(user_lexicon or {})

    # -- 取消控制 ---------------------------------------------------------

    def cancel(self) -> None:
        """请求中止当前批处理。"""
        self._cancelled = True
        proc = self._current
        if proc and proc.poll() is None:
            try:
                proc.terminate()
            except OSError:
                pass

    def reset(self) -> None:
        self._cancelled = False

    # -- 进程内后端探测 ---------------------------------------------------

    @staticmethod
    def _init_inproc():
        """尝试导入 mask-tool 的 Pipeline，以便冻结后进程内调用。"""
        try:
            from mask_tool.core.pipeline import Pipeline
            from mask_tool.models.config import MaskConfig
            from mask_tool.models.detection import DetectionStatus
            # 修复上游 docx 多实体替换串扰缺陷
            _patch_mask_tool_adapters()
            # 补充高置信正则（中文金额 / 项目工程名 / 纯数字金额）
            _patch_detector_rules()
            # 注入「检测灵敏度」可控的 smart/aggressive 策略
            _patch_policy()
            return (Pipeline, MaskConfig, DetectionStatus)
        except Exception:
            return None

    @staticmethod
    def _build_user_lexicon(user_lexicon: dict) -> str | None:
        """把打包基准词库与用户词库合并，写出临时 YAML，返回路径。

        用户词库为空时返回 ``None``，引擎改用打包基准词库
        （``assets/mask_tool_config/sample_lexicon.yaml``）。
        词库匹配置信度为 0.95，strict/smart/aggressive 三档都会自动替换，
        因此用户词库对所有模式都有效，尤其 strict 几乎完全依赖它。
        """
        if not user_lexicon:
            return None

        # 基准词库（打包进 exe 的示例词库）
        base: dict = {}
        cfg_dir = _bundled_config_dir()
        if cfg_dir is not None:
            sample = cfg_dir / "sample_lexicon.yaml"
            if sample.is_file():
                try:
                    import yaml
                    loaded = yaml.safe_load(sample.read_text(encoding="utf-8"))
                    if isinstance(loaded, dict):
                        base = loaded
                except Exception:
                    base = {}

        merged: dict = {k: list(v) for k, v in base.items()}
        for cat, words in user_lexicon.items():
            if not isinstance(words, list):
                continue
            cur = list(merged.get(cat, []))
            for w in words:
                if isinstance(w, str) and w.strip() and w not in cur:
                    cur.append(w)
            if cur:
                merged[cat] = cur

        if not merged:
            return None

        tmp_dir = tempfile.mkdtemp(prefix="masktool_lex_")
        out = Path(tmp_dir) / "user_lexicon.yaml"
        try:
            import yaml
            out.write_text(
                yaml.dump(merged, allow_unicode=True, default_flow_style=False),
                encoding="utf-8",
            )
            atexit.register(shutil.rmtree, tmp_dir, ignore_errors=True)
            return str(out)
        except Exception:
            return None

    # -- 核心 -------------------------------------------------------------

    def process_one(
        self,
        src: Path,
        out_dir: Path,
        mode: str = "smart",
        save_mapping: bool = True,
        suffix_tag: str = "_脱敏",
        sensitivity: str = SENSITIVITY_DEFAULT,
    ) -> FileResult:
        """处理单个文件。任何异常都被转换为 FileResult，不向上抛。"""
        ext = src.suffix.lower()

        if not src.is_file():
            return FileResult(src, False, message="文件不存在或已被移动")
        if ext not in SUPPORTED_EXTS:
            return FileResult(src, False, message=f"不支持的格式 {ext or '（无扩展名）'}")
        if mode not in MODES:
            mode = "smart"

        # 优先使用进程内调用（冻结打包后唯一可靠路径）
        if self._inproc is not None:
            return self._process_one_inproc(
                src, out_dir, mode, save_mapping, suffix_tag, sensitivity
            )

        report_only = ext in REPORT_ONLY_EXTS

        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            return FileResult(src, False, message=f"无法创建输出目录：{exc.strerror or exc}")

        with tempfile.TemporaryDirectory(prefix="masktool_") as tmp:
            tmp_dir = Path(tmp)
            argv = [
                *self.tool.argv,
                "mask",
                str(src),
                "--output", str(tmp_dir),
                "--mode", mode,
                "--no-learn",
            ]

            try:
                proc = subprocess.Popen(argv, **_subprocess_kwargs())
                self._current = proc
                try:
                    stdout, _ = proc.communicate(timeout=PER_FILE_TIMEOUT)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.communicate()
                    return FileResult(src, False, message="处理超时（超过 10 分钟），已中止")
                finally:
                    self._current = None
            except FileNotFoundError:
                return FileResult(src, False, message="mask-tool 命令已失效，请重新检测")
            except OSError as exc:
                return FileResult(src, False, message=f"启动脱敏进程失败：{exc.strerror or exc}")

            if self._cancelled:
                return FileResult(src, False, message="已取消")

            artifact = _expected_artifact(tmp_dir, src)
            if artifact is None:
                return FileResult(src, False, message=_extract_error(stdout))

            masked_count = _read_masked_count(tmp_dir, artifact, report_only)

            # 搬运产物
            try:
                target = _unique_path(out_dir / build_output_name(src, suffix_tag))
                shutil.copy2(artifact, target)
            except (OSError, MaskToolCallError) as exc:
                return FileResult(src, False, message=f"保存结果失败：{exc}")

            # 搬运映射表（可逆脱敏还原时需要）
            mapping_target: Path | None = None
            src_mapping = tmp_dir / "mapping.json"
            if save_mapping and not report_only and src_mapping.is_file():
                try:
                    mapping_target = _unique_path(
                        out_dir / f"{src.stem}{suffix_tag}_映射表.json"
                    )
                    shutil.copy2(src_mapping, mapping_target)
                except OSError:
                    mapping_target = None

            msg = (
                f"仅生成检测报告（当前 mask-tool 不回写 PDF），检出 {masked_count} 项"
                if report_only
                else f"已脱敏 {masked_count} 项"
            )
            return FileResult(
                source=src,
                ok=True,
                output=target,
                mapping=mapping_target,
                masked_count=masked_count,
                message=msg,
                report_only=report_only,
            )

    def _process_one_inproc(
        self,
        src: Path,
        out_dir: Path,
        mode: str,
        save_mapping: bool,
        suffix_tag: str,
        sensitivity: str = SENSITIVITY_DEFAULT,
    ) -> FileResult:
        """进程内调用 mask-tool 的 Pipeline（冻结打包后的主路径）。"""
        ext = src.suffix.lower()
        report_only = ext in REPORT_ONLY_EXTS

        # 应用用户选择的 smart/aggressive 检测灵敏度（仅对 NER 生效）
        set_smart_sensitivity(sensitivity)

        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            return FileResult(src, False, message=f"无法创建输出目录：{exc.strerror or exc}")

        Pipeline, MaskConfig, _DetectionStatus = self._inproc
        cfg_dir = _bundled_config_dir()
        cfg = None
        if cfg_dir is not None:
            default_yaml = cfg_dir / "default.yaml"
            if default_yaml.is_file():
                # 用官方默认配置（已开启 jieba NER 与合理阈值），仅覆盖运行模式
                try:
                    cfg = MaskConfig.from_yaml(default_yaml)
                    cfg.mode = mode
                except Exception:
                    cfg = None
        if cfg is None:
            cfg = MaskConfig(mode=mode)

        # 把白名单指向上面打包进 exe 的资源，保证冻结后也能加载
        if cfg_dir is not None:
            wlt = cfg_dir / "whitelist.yaml"
            if wlt.is_file():
                cfg.whitelist_path = str(wlt)

        # 用户词库优先（已合并基准词库）；否则用打包基准词库
        if self._user_lexicon_file and Path(self._user_lexicon_file).is_file():
            cfg.lexicon_path = self._user_lexicon_file
        elif cfg_dir is not None:
            lex = cfg_dir / "sample_lexicon.yaml"
            if lex.is_file():
                cfg.lexicon_path = str(lex)

        with tempfile.TemporaryDirectory(prefix="masktool_") as tmp:
            tmp_dir = Path(tmp)
            try:
                pipeline = Pipeline(cfg)
                result_path = pipeline.process_file(src, tmp_dir)
            except Exception as exc:  # 任何内部异常都转为人话
                return FileResult(src, False, message=_humanize(str(exc))[:200])

            if result_path is None:
                return FileResult(src, False, message="mask-tool 不支持该文件格式")
            artifact = Path(result_path)
            if not artifact.is_file():
                return FileResult(src, False, message="脱敏未产生输出文件")

            # 写出映射表与报告，便于结果统计与后续还原
            try:
                pipeline.save_mapping(tmp_dir / "mapping.json")
                pipeline.save_report(tmp_dir / "report.json")
            except Exception:
                pass

            if report_only:
                try:
                    data = json.loads(artifact.read_text(encoding="utf-8"))
                    masked_count = int(data.get("total_detections", 0))
                except Exception:
                    masked_count = 0
            else:
                try:
                    masked_count = len(pipeline.masker.get_mappings())
                except Exception:
                    masked_count = 0

            try:
                target = _unique_path(out_dir / build_output_name(src, suffix_tag))
                shutil.copy2(artifact, target)
            except (OSError, MaskToolCallError) as exc:
                return FileResult(src, False, message=f"保存结果失败：{exc}")

            mapping_target: Path | None = None
            src_mapping = tmp_dir / "mapping.json"
            if save_mapping and not report_only and src_mapping.is_file():
                try:
                    mapping_target = _unique_path(
                        out_dir / f"{src.stem}{suffix_tag}_映射表.json"
                    )
                    shutil.copy2(src_mapping, mapping_target)
                except OSError:
                    mapping_target = None

            msg = (
                f"仅生成检测报告（当前 mask-tool 不回写 PDF），检出 {masked_count} 项"
                if report_only
                else f"已脱敏 {masked_count} 项"
            )
            return FileResult(
                source=src, ok=True, output=target, mapping=mapping_target,
                masked_count=masked_count, message=msg, report_only=report_only,
            )

    def process_batch(
        self,
        files: Iterable[Path],
        out_dir_for: Callable[[Path], Path],
        mode: str = "smart",
        save_mapping: bool = True,
        suffix_tag: str = "_脱敏",
        sensitivity: str = SENSITIVITY_DEFAULT,
        on_progress: Callable[[int, int, Path], None] | None = None,
        on_result: Callable[[FileResult], None] | None = None,
    ) -> BatchOutcome:
        """批量处理。``out_dir_for`` 用于为每个源文件决定输出目录。"""
        self.reset()
        items = list(files)
        outcome = BatchOutcome()

        for idx, src in enumerate(items, start=1):
            if self._cancelled:
                outcome.cancelled = True
                break
            if on_progress:
                on_progress(idx, len(items), src)

            res = self.process_one(
                src, out_dir_for(src), mode=mode,
                save_mapping=save_mapping, suffix_tag=suffix_tag,
                sensitivity=sensitivity,
            )
            outcome.results.append(res)
            if on_result:
                on_result(res)

        return outcome


def _extract_error(stdout: str | None) -> str:
    """从 mask-tool 的输出里提炼出人话错误信息。"""
    if not stdout:
        return "脱敏未产生输出文件，原因未知"

    text = stdout.strip()

    # mask-tool 逐文件失败时打印 "✗ <错误>"
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("✗") or "✗" in line:
            detail = line.split("✗", 1)[-1].strip()
            if detail:
                return _humanize(detail)

    lowered = text.lower()
    if "pymupdf" in lowered or "fitz" in lowered:
        return "缺少 PDF 解析库 PyMuPDF，请执行 pip install pymupdf"
    if "permission" in lowered or "denied" in lowered:
        return "文件被占用或无写入权限，请关闭文档后重试"
    if "traceback" in lowered:
        last = [l.strip() for l in text.splitlines() if l.strip()][-1]
        return _humanize(last)

    tail = [l.strip() for l in text.splitlines() if l.strip()]
    return _humanize(tail[-1]) if tail else "脱敏未产生输出文件"


def _humanize(raw: str) -> str:
    """把常见的英文异常翻译成中文提示。"""
    low = raw.lower()
    table = [
        ("package not found", "文档格式无法识别，可能不是有效的 Office 文件"),
        ("badzipfile", "文件已损坏或不是有效的 .docx/.xlsx/.pptx"),
        ("permissionerror", "文件被占用或无写入权限，请关闭文档后重试"),
        ("no such file", "找不到该文件，可能已被移动或删除"),
        ("memoryerror", "文件过大导致内存不足"),
        ("cannot open broken document", "PDF 文件已损坏，无法打开"),
    ]
    for key, msg in table:
        if key in low:
            return msg
    return raw[:200]
