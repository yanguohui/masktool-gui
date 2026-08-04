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
from collections.abc import Iterable
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
#     其中「项目/工程名」正则已收紧：必须是边界起始、含实体指示词、且不以动词/
#     指代词开头的专有名词短语，避免把「根据项目」「卖方应在本项目」整句吞掉。
#   · jieba NER 的置信度上限仅 0.85，且「人名/机构/地名」的置信度分布差异很大：
#       机构名（company）几乎都在 0.83 左右，可靠；
#       人名（person）多在 0.60~0.70，偏低的真名与偏高的误报并存；
#       地名（location）误报极多（"大海""新建"等都被误标），不可信；
#       其他专名（custom，nz）鱼龙混杂。
#     因此「灵敏度」按实体类型分别设阈值，而非一刀切：
#       - company 阈值最低（最可靠），person 次之，custom 较严，location 最严。
#   · 在策略层额外加了一道「NER 误报抑制」（_ner_should_suppress，与灵敏度无关、
#     对 smart/aggressive 生效）：精确白名单（网络安全/买卖双方/根据/日志…）+ 类型
#     校验（人名需 2~3 字且首字为姓氏；地名需以省/市/区/县/路…后缀结尾）+ 单字实体
#     强制不脱敏。这样即便调到「高（激进召回）」，根据/买卖双方/网络安全等无关词
#     也不会被脱敏。
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


# --------------------------------------------------------------------------
# 置信度下限：按实体类型分别设置（低于各自下限的一律「只提示、不替换」）
# --------------------------------------------------------------------------
# 这是一道**凌驾于灵敏度档位之上**的硬闸门，对 dictionary / regex / ner
# 全部来源生效。业务含义：宁可漏掉一个把握不大的猜测，也不要动到正文用词。
#
# 与「检测灵敏度」的分工：
#   · 置信度下限 = 绝对底线，回答「这条识别本身可不可信」；
#   · 检测灵敏度 = 底线之上的取舍，回答「这类实体要不要那么激进」。
#
# 按实体类型分别设下限：机构名（company）/ 项目名（project）构词稳定、可靠，
# 可放低门槛以提召回；地名（location）误报极多、须更严；人名（person）置信度
# 分布分散且上限低，给最低门槛保召回（误报由构词法复核闸门二次拦截）。
MIN_CONFIDENCE_BY_TYPE: dict[str, float] = {
    "company": 0.75,
    "person": 0.60,
    "location": 0.80,
    "project": 0.75,
}
#: 未列入上表的类型（government / subject / custom / amount …）兜底下限
MIN_CONFIDENCE_DEFAULT_FLOOR = 0.80

#: 全局「兜底下限」（由 UI 设置）。0 表示不额外抬高，完全按各类型下限；
#: 若被设为 >0，则作为所有类型下限的统一抬高值（取 max）。
MIN_CONFIDENCE_DEFAULT = 0.0

#: 当前生效的全局兜底下限（由 UI / 配置在每次处理前设置）
_MIN_CONFIDENCE: float = MIN_CONFIDENCE_DEFAULT


def _min_conf_for_type(name: str) -> float:
    """返回某实体类型的置信度下限。

    ``name`` 为 DetectionType 的小写字符串（company/person/location/
    project/government/subject/custom/amount…）。全局兜底下限 ``_MIN_CONFIDENCE``
    若 >0，则作为统一抬高值参与 max。
    """
    base = MIN_CONFIDENCE_BY_TYPE.get(name, MIN_CONFIDENCE_DEFAULT_FLOOR)
    if _MIN_CONFIDENCE > 0:
        return max(base, _MIN_CONFIDENCE)
    return base


def set_min_confidence(value: float | None) -> None:
    """设置全局「兜底下限」（0 表示按各类型下限，不额外抬高）。"""
    global _MIN_CONFIDENCE
    try:
        v = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        v = MIN_CONFIDENCE_DEFAULT
    _MIN_CONFIDENCE = min(max(v, 0.0), 1.0)


def get_min_confidence() -> float:
    return _MIN_CONFIDENCE


def _type_name_of(result) -> str:
    """提取 DetectionResult 的实体类型名（小写），便于查表。"""
    tt = getattr(result, "text_type", None)
    if tt is None:
        return ""
    return str(getattr(tt, "value", tt)).lower()


# --------------------------------------------------------------------------
# 程序根目录 & whitelist.txt（用户可直接编辑的「绝不脱敏」词表）
# --------------------------------------------------------------------------


def app_root() -> Path:
    """返回「程序根目录」——用户能直接看到并编辑配置文件的那个目录。

    - 冻结模式：exe 所在目录（用户双击的那个文件旁边）；
    - 开发模式：项目根目录（``app/`` 的上一级）。
    """
    if IS_FROZEN:
        try:
            return Path(sys.executable).resolve().parent
        except OSError:
            return Path.cwd()
    return Path(__file__).resolve().parent.parent


WHITELIST_FILENAME = "whitelist.txt"

#: 首次运行时自动生成的模板内容
_WHITELIST_TEMPLATE = """\
# ==========================================================================
# 白名单 —— 写在这里的词「绝不脱敏」
# ==========================================================================
# 用法：
#   · 一行一个词；也可用中文逗号「，」或英文逗号「,」在一行里写多个
#   · 以 # 开头的行是注释，会被忽略
#   · 保存后无需重启程序，下一次点「开始脱敏」即自动重新读取
#
# 匹配规则（重要）：
#   按「整条识别结果」精确匹配，不做包含匹配。
#   例：白名单写了「工程」，则单独被识别出的「工程」不会脱敏，
#       而「上海市建设工程监理咨询有限公司」仍会正常脱敏（它是一个完整机构名）。
#   程序会自动忽略「本 / 该 / 此 / 各 / 上述」等前缀和结尾的「的」，
#   因此写「工程」即可同时保护「本工程」「该工程」。
# --------------------------------------------------------------------------

# ---- 合同 / 商务通用术语 ----
二次开发
许可
许可证
授权
卖方
买方
买卖双方
甲方
乙方
丙方
供方
需方
承包方
发包方
中标人
投标人
招标人

# ---- 工程 / 技术通用术语 ----
工程
项目
标段
系统
平台
网络
设备
软件
硬件
接口
验收
调试
运维
质保
技术文档
技术规范
施工
安装
集成

# ---- 常见非敏感套话（易被 NER 误判为人名 / 地名 / 机构）----
根据
按照
本项目
本工程
本合同
通用条款
专用条款
网络安全
信息安全
严格遵守
中华民族
伟大复兴
新建
改造
扩容
"""


def whitelist_path() -> Path:
    """whitelist.txt 的完整路径（程序根目录下）。"""
    return app_root() / WHITELIST_FILENAME


def ensure_whitelist_file() -> Path:
    """确保 whitelist.txt 存在；不存在则写入模板。返回其路径。

    任何写入失败都静默忽略（例如程序装在只读目录），此时白名单为空，
    不影响主流程。
    """
    p = whitelist_path()
    try:
        if not p.is_file():
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(_WHITELIST_TEMPLATE, encoding="utf-8")
    except OSError:
        pass
    return p


#: (mtime, size, 词集) —— 按文件指纹缓存，用户改完文件立即生效
_WL_CACHE: tuple[float, int, frozenset[str]] | None = None


def _parse_whitelist_text(text: str) -> set[str]:
    """解析白名单文本：# 注释、空行忽略，支持逗号分隔多词。"""
    words: set[str] = set()
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        # 行内注释：「工程   # 通用术语」
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        for part in re.split(r"[,，、;；\t]+", line):
            w = part.strip()
            if w:
                words.add(w)
    return words


def load_whitelist(force: bool = False) -> frozenset[str]:
    """读取 whitelist.txt（带文件指纹缓存）。

    文件缺失时自动生成模板；读取失败返回空集合，绝不影响主流程。
    """
    global _WL_CACHE
    p = ensure_whitelist_file()
    try:
        st = p.stat()
        fp = (st.st_mtime, st.st_size)
    except OSError:
        return frozenset()

    if not force and _WL_CACHE is not None:
        if (_WL_CACHE[0], _WL_CACHE[1]) == fp:
            return _WL_CACHE[2]

    try:
        words = _parse_whitelist_text(p.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError):
        try:
            words = _parse_whitelist_text(p.read_text(encoding="gbk"))
        except (OSError, UnicodeDecodeError, LookupError):
            words = set()

    wl = frozenset(words)
    _WL_CACHE = (fp[0], fp[1], wl)
    return wl


#: 归一化时剥离的限定前缀（汉语指示词闭集，非实体清单）
_WL_DET_PREFIX = ("本项", "该项", "上述", "本次", "本", "该", "此", "各", "其")


def _wl_normalize(text: str) -> set[str]:
    """生成一条识别结果的等价写法集合，用于与白名单比对。

    「本工程」「该工程」「工程的」都能匹配到白名单里的「工程」。
    """
    s = (text or "").strip()
    if not s:
        return set()
    forms = {s, s.lower()}
    # 去掉结尾的「的」
    core = s[:-1] if len(s) > 1 and s.endswith("的") else s
    forms.add(core)
    # 去掉指示性前缀
    for pre in _WL_DET_PREFIX:
        if core.startswith(pre) and len(core) > len(pre) + 1:
            forms.add(core[len(pre):])
            break
    return {f for f in forms if f}


def is_whitelisted(text: str, whitelist: frozenset[str] | set[str] | None = None) -> bool:
    """判断一条识别结果是否命中白名单（整条精确匹配 + 限定词归一化）。"""
    wl = load_whitelist() if whitelist is None else whitelist
    if not wl:
        return False
    return bool(_wl_normalize(text) & set(wl))


def save_whitelist(words: Iterable[str]) -> bool:
    """将白名单词集写回 ``whitelist.txt``（每行一词）。返回是否写入成功。

    GUI 的「白名单…」编辑器调用此函数。成功后会强制刷新文件指纹缓存，
    下一次脱敏立即生效。写入失败（如程序装在只读目录）返回 ``False``，
    不影响主流程。
    """
    p = whitelist_path()
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        ordered = sorted(
            {w.strip() for w in words if w and w.strip()},
            key=lambda s: (len(s), s),
        )
        header = (
            "# 白名单 —— 写在这里的词「绝不脱敏」\n"
            "# 一行一个词；保存后无需重启，下一次脱敏自动生效。\n"
            "# 程序自动忽略「本/该/此」等前缀与结尾的「的」，故「工程」即可保护「本工程」。\n"
            "# -------------------------------------------------------------------------\n\n"
        )
        body = "\n".join(ordered)
        p.write_text(header + body + ("\n" if body else ""), encoding="utf-8")
    except OSError:
        return False
    load_whitelist(force=True)
    return True


# --------------------------------------------------------------------------
# 自主识别层（构词法 + 上下文驱动，不依赖实体词库 / 映射表）
# --------------------------------------------------------------------------
# 设计原则：不去「记住」哪些公司、项目、人名是敏感的——那需要无穷无尽的词库，
# 换一份文档就失效。改为利用中文专名的「构词法」与「上下文位置」自主判断：
#
#   1) 组织机构名 = [专名 / 字号] + [行业属性] + [组织形式后缀]
#      「组织形式后缀」（有限公司 / 集团 / 研究院 / 管理局 / 人民政府 …）是
#      汉语机构命名法的闭集，属语言规则而非实体清单。只要以其收尾且前缀合法，
#      无论这家单位此前是否见过都能识别 —— 「中电信数智科技有限公司」这类
#      全新机构无需任何词库即可命中，且能拿到完整名称（不会被 NER 截断）。
#
#   2) 地址 = 行政区划链（省 / 市 / 区 / 县）+ 街路巷号 / 园区 / 大厦
#
#   3) 人名 = 姓氏（汉语姓氏闭集）+ 1~2 字名；或由「联系人 / 负责人 / 签字」
#      等结构化标签、职务称谓等上下文触发。
#
#   4) 项目名 = [限定修饰] + [项目类中心词]（项目 / 工程 / 标段 / 子系统 …），
#      中心词分强弱：强中心词（项目 / 工程 / 标段）可直接认定；弱中心词
#      （系统 / 平台 / 网络）过于泛用，须有足够长的专名修饰或处在标题行 /
#      结构化字段中才认定，避免把「本系统」「开放平台」当成项目名。
#
#   5) 反向判据（取代「误报词黑名单」）：一个候选若在构词上不具备上述任何
#      专名特征 —— 没有组织后缀、没有地名后缀、不是姓名结构、不含项目中心词
#      —— 即判定为普通词组，一律不脱敏。于是「网络安全 / 买卖双方 / 根据 /
#      日志 / 深度 / 加密 / 外联 / 中华民族 / 严格遵守」等无需逐个登记，
#      也会被自动排除；反之新出现的普通词同样自动排除，不需要再维护词表。
#
# 下面这些常量都是「语素级构词规则」（后缀 / 前缀 / 虚词 / 姓氏），
# 不是「实体映射库」——它们不记录任何一家具体公司、项目或人名。

#: 组织形式「强后缀」：出现即可判定为机构。
#: 顺序要求：长后缀排在短后缀之前，保证正则交替时取到最长形式。
_ORG_SUFFIX_STRONG: tuple[str, ...] = (
    "股份有限责任公司", "股份有限公司", "有限责任公司", "集团有限公司",
    "有限公司", "总公司", "分公司", "子公司", "公司",
    "集团股份", "集团", "控股", "实业",
    "人民政府", "人民法院", "人民检察院", "人民医院", "人民银行",
    "管理委员会", "工作委员会", "委员会", "监督管理局", "管理局",
    "科学院", "研究院", "设计院", "规划院", "医学院", "职业学院", "学院",
    "研究所", "设计所", "事务所", "研究中心", "检测中心", "疾控中心",
    "大学", "中学", "小学", "医院", "卫生院",
    "银行", "支行", "分行", "信用社", "证券", "保险",
    "协会", "学会", "商会", "基金会", "合作社", "联合会", "促进会",
    "事业部", "办事处", "指挥部", "管理处", "监理站",
    "党支部", "党委", "工会",
)

#: 组织形式「弱后缀」：过于泛用（数据中心 / 生产车间…），
#: 需要更长的专名前缀佐证才认定为机构。
_ORG_SUFFIX_WEAK: tuple[str, ...] = (
    "中心", "基地", "工厂", "电厂", "矿业", "农场", "林场",
    "局", "厅", "处", "科", "室", "站", "所", "厂", "队", "园区",
)

#: 组织形式「中等后缀」：单独看过于泛用，需 >=4 字专名前缀佐证。
#: 覆盖政府部门（…监督局 /…管理厅 /…执法总队）与各类中心 / 基地。
_ORG_SUFFIX_MEDIUM: tuple[str, ...] = (
    "办公室", "总队", "支队", "大队", "中心", "基地",
    "局", "厅", "署", "园区",
)

#: 行政区划 / 地址构词后缀
_GEO_SUFFIX: tuple[str, ...] = (
    "特别行政区", "自治区", "自治州", "自治县",
    "省", "市", "区", "县", "旗", "州",
    "镇", "乡", "村", "街道", "社区",
    "路", "街", "大道", "巷", "号", "栋", "幢", "楼", "室",
    "软件园", "科技园", "工业园", "开发区", "高新区", "园区", "新区",
    "大厦", "广场", "小区", "公寓",
)

#: 机构性构词前缀：以国名 / 行政层级起头的组合几乎必为机构
#: （中国联通 / 国家电网 / 中央结算…），用于识别无标准后缀的机构简称。
_ORG_HEAD: tuple[str, ...] = ("中国", "中华", "国家", "中央", "全国", "国务院")

#: 抽象名词构词后缀——以此收尾的中文词是抽象概念，不可能是机构 / 地名 / 人名。
#: 这是「语素级」规则而非实体黑名单：任何以「…性 / …度 / …化 / …族」收尾的词
#: （安全性 / 深度 / 云化 / 中华民族 …）都会被自动排除，无需逐词登记。
_ABSTRACT_TAIL: tuple[str, ...] = (
    "思想", "精神", "文明", "传统", "方式", "方法", "过程", "状态", "水平",
    "能力", "要求", "标准", "规范", "原则", "措施", "手段", "内容", "功能",
    "性能", "指标", "范围", "条件", "环境", "基础", "核心", "特点", "优势",
    "问题", "情况", "结果", "目标", "任务", "职责", "义务", "权利", "责任",
    "性", "度", "化", "族", "观", "史", "论", "说", "率", "量",
    "力", "感", "式", "型", "法", "制", "义", "派", "风", "情",
)

#: 功能语素 / 虚词 / 动词起始——以此开头的片段是句子成分而非专名。
#: 同样是构词规则（汉语虚词与高频谓词是闭集），不记录任何实体。
_FUNC_HEAD: tuple[str, ...] = (
    "根据", "按照", "依照", "遵照", "依据", "通过", "对于", "关于", "针对",
    "鉴于", "为了", "由于", "如果", "虽然", "但是", "并且", "同时", "以及",
    "或者", "以便", "从而", "因此", "所以", "凡是", "任何", "所有", "其他",
    "提出", "进行", "完成", "实现", "采用", "使用", "要求", "需要", "负责",
    "成立", "提交", "监督", "保证", "确保", "涉及", "包括", "参加", "开展",
    "具备", "满足", "支持", "提供", "承担", "遵守", "执行", "落实", "加强",
    "本", "该", "此", "其", "各", "这", "那", "上述", "下列", "有关", "相关",
    "卖方", "买方", "双方", "各方", "甲方", "乙方", "丙方", "我方", "对方",
)

#: 项目类「强中心词」：以此收尾即可认定为项目 / 工程名
_PROJECT_CORE_STRONG: tuple[str, ...] = (
    "项目", "工程", "标段", "标包", "课题", "专项", "工程建设",
)

#: 项目类「弱中心词」：过于泛用，需长修饰或结构化上下文佐证
_PROJECT_CORE_WEAK: tuple[str, ...] = (
    "子系统", "系统", "平台", "网络", "基地", "中心",
)

#: 项目 / 工程名内部须出现的实体指示语素（区分真专名与「重要项目」这类泛指）
#: 注：已移除过于泛用的「服务 / 软件 / 硬件」（几乎任何业务短语都含这些字眼，
#: 保留只会放大误报），仅保留具备专名区分度的语素。
_PROJECT_INDICATOR: tuple[str, ...] = (
    "系统", "平台", "建设", "网络", "扩容", "改造", "升级",
    "中心", "体系", "能力", "研发", "基地", "方案", "应用", "采购",
    "集成", "运维", "安防", "监控", "感知", "数据", "信息", "智能", "数字",
    "网", "院", "厂", "站", "所", "线", "库", "实验室", "子系统",
)

#: 汉语单姓（闭集构词成分，非人名库——它不记录任何一个具体姓名）
_SURNAMES: frozenset[str] = frozenset(
    "王李张刘陈杨黄赵周吴徐孙朱马胡郭林何高梁郑罗宋谢唐韩冯于董"
    "萧程曹袁邓许傅沈曾彭吕苏卢蒋蔡贾丁魏薛叶阎余潘杜戴夏钟汪田任"
    "姜范方石姚谭廖邹熊金陆郝孔白崔康毛邱秦江史顾侯邵孟龙万段雷钱"
    "汤尹黎易常武乔贺赖龚文庞樊兰殷颜鲁韦毕聂庄卓项祝霍骆包诸左"
    "石戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁昌马苗凤花方俞任袁柳"
    "酆鲍史唐费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮卞齐康伍余元"
    "卜顾孟平黄和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞熊纪"
    "舒屈项祝董梁杜阮蓝闵席季麻强贾路娄危江童颜郭梅盛林刁钟徐邱骆高"
    "夏蔡田樊胡凌霍虞万支柯昝管卢莫经房裘缪干解应宗丁宣贲邓郁单杭洪"
    "包诸左石崔吉钮龚程嵇邢滑裴陆荣翁荀羊於惠甄曲家封芮羿储靳汲邴糜"
    "松井段富巫乌焦巴弓牧隗山谷车侯宓蓬全郗班仰秋仲伊宫宁仇栾暴甘钭"
    "厉戎祖武符刘景詹束龙叶幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓"
    "蔺屠蒙池乔阴鬱胥能苍双闻莘党翟谭贡劳逄姬申扶堵冉宰郦雍卻璩桑桂"
    "濮牛寿通边扈燕冀郏浦尚农温别庄晏柴瞿阎充慕连茹习宦艾鱼容向古易"
    "慎戈廖庾终暨居衡步都耿满弘匡国文寇广禄阙东欧殳沃利蔚越夔隆师巩"
    "厍聂晁勾敖融冷訾辛阚那简饶空曾毋沙乜养鞠须丰巢关蒯相查后荆红游"
    "竺权逯盖益桓公"
)

#: 汉语复姓
_SURNAMES_DOUBLE: frozenset[str] = frozenset([
    "欧阳", "太史", "端木", "上官", "司马", "东方", "独孤", "南宫", "万俟",
    "闻人", "夏侯", "诸葛", "尉迟", "公羊", "赫连", "澹台", "皇甫", "宗政",
    "濮阳", "公冶", "太叔", "申屠", "公孙", "慕容", "仲孙", "钟离", "长孙",
    "宇文", "司徒", "鲜于", "司空", "闾丘", "子车", "亓官", "司寇", "巫马",
    "公西", "颛孙", "壤驷", "公良", "漆雕", "乐正", "宰父", "谷梁", "拓跋",
    "夹谷", "轩辕", "令狐", "段干", "百里", "呼延", "东郭", "南门", "羊舌",
    "微生", "公户", "公玉", "公仪", "梁丘", "公仲", "公上", "公门", "公山",
])

#: 人口高频姓氏先验（约前 120 单姓 + 常见复姓）。
#: 用于 NER 人名的「误报抑制」闸门：真实人名绝大多数以高频姓氏起头，
#: 罕见姓氏的 NER 命中更可能是普通词（「明白」以「明」起头、「应予以」以
#: 「应」起头）。这是统计先验而非实体库——目的是在高 precision 下抑制误报；
#: 罕见姓氏的真实人名若出现于「联系人 / 负责人：」等标签后，仍由
#: _looks_like_person_labeled（完整姓氏集）兜底命中，不会漏检。
_SURNAMES_COMMON: frozenset[str] = frozenset(
    "王李张刘陈杨黄赵周吴徐孙朱马胡郭林何高梁郑罗宋谢唐韩冯于董"
    "萧程曹袁邓许傅沈曾彭吕苏卢蒋蔡贾丁魏薛叶阎余潘杜戴夏钟汪田任"
    "姜范方石姚谭廖邹熊金陆郝孔白崔康毛邱秦江史顾侯邵孟龙万段雷钱"
    "汤尹黎易常武乔贺赖龚文庞樊兰殷颜鲁韦毕聂庄卓项祝霍骆包诸左"
    "石戚喻柏水窦章云葛奚郎昌苗凤花俞柳"
).union(frozenset([
    "欧阳", "司马", "上官", "诸葛", "东方", "独孤", "南宫", "慕容", "仲孙",
    "皇甫", "夏侯", "尉迟", "公孙", "令狐", "宇文", "司徒", "长孙", "万俟",
    "闻人", "端木",
]))


def _starts_with_function_word(text: str) -> bool:
    """片段是否以虚词 / 高频谓词开头（说明它是句子成分，不是专名）。"""
    return text.startswith(_FUNC_HEAD)


def _ends_with_abstract(text: str) -> bool:
    """片段是否以抽象名词语素收尾（说明它是概念而非专名）。"""
    return text.endswith(_ABSTRACT_TAIL)


def _ends_with_function_word(text: str) -> bool:
    """片段是否以虚词 / 谓词收尾（如「国家有关」「本项目相关」）。"""
    return text.endswith(_FUNC_HEAD)


def _looks_like_org(text: str) -> bool:
    """按中文机构命名法判断是否为组织机构名（无需机构词库）。"""
    n = len(text)
    if n < 3 or _starts_with_function_word(text):
        return False
    for suf in _ORG_SUFFIX_STRONG:
        if text.endswith(suf) and n - len(suf) >= 2:
            return True
    for suf in _ORG_SUFFIX_WEAK:
        # 弱后缀要求更长的专名前缀，避免「数据中心」「生产车间」误判
        if text.endswith(suf) and n - len(suf) >= 4:
            return True
    # 「中国 / 国家 / 中央 …」+ 专名 的机构简称（中国联通、国家电网），
    # 但排除「中华民族」「中国传统」这类以抽象语素收尾的普通词组。
    if (text.startswith(_ORG_HEAD) and 4 <= n <= 12
            and not _ends_with_abstract(text)
            and not _ends_with_function_word(text)):
        return True
    return False


def _looks_like_geo(text: str) -> bool:
    """按行政区划 / 地址构词法判断是否为地点名。"""
    n = len(text)
    if n < 3 or _starts_with_function_word(text) or _ends_with_abstract(text):
        return False
    for suf in _GEO_SUFFIX:
        if text.endswith(suf) and n - len(suf) >= 2:
            return True
    return False


def _looks_like_person(text: str) -> bool:
    """按「高频姓氏 + 名」构词法判断 NER 候选是否为真实人名（误报抑制闸门）。

    此处用高频姓氏先验（_SURNAMES_COMMON）而非完整百家姓：罕见姓氏的 NER
    命中绝大多数是普通词（明白 / 应予以…），据此过滤可保精度；标签语境下
    的罕见姓氏人名由 _looks_like_person_labeled（完整姓氏集）兜底。
    """
    n = len(text)
    if not (2 <= n <= 4):
        return False
    if not all("\u4e00" <= c <= "\u9fff" for c in text):
        return False
    if _ends_with_abstract(text) or _starts_with_function_word(text):
        return False
    if n >= 3 and text[:2] in _SURNAMES_DOUBLE and text[:2] in _SURNAMES_COMMON:
        return True
    return text[0] in _SURNAMES_COMMON


def _looks_like_person_labeled(text: str) -> bool:
    """结构化字段（联系人 / 负责人：）里的值是否为人名。

    标签本身已提供「这是人」的语境，因此不强制要求命中姓氏闭集，
    只要求它不是机构 / 抽象概念 / 句子成分。
    """
    n = len(text)
    if not (2 <= n <= 4):
        return False
    if not all("\u4e00" <= c <= "\u9fff" for c in text):
        return False
    if _ends_with_abstract(text) or _starts_with_function_word(text):
        return False
    if _ends_with_function_word(text):
        return False
    if text.endswith(_ORG_SUFFIX_STRONG) or text.endswith(_ORG_SUFFIX_WEAK):
        return False
    if text.endswith(_GEO_SUFFIX):
        return False
    if n >= 3 and text[:2] in _SURNAMES_DOUBLE:
        return True
    if text[0] in _SURNAMES:
        return True
    # 兜底：分词器整体判定为人名
    toks = _pos_tokens(text)
    return len(toks) == 1 and toks[0][1] == "nr"


def _looks_like_project(text: str) -> bool:
    """按「修饰 + 项目类中心词」构词法判断是否为项目 / 工程名。

    长度门槛 ≥6 字符：过短的片段（如「本工程」「该平台」）多为泛指或
    句子成分，不具备专名区分度，直接排除以收紧误报。
    """
    n = len(text)
    if n < 6 or _starts_with_function_word(text):
        return False
    # 「重要项目」「该工程」这类泛指不含实体指示语素，据此与真项目名区分
    has_ind = any(ind in text for ind in _PROJECT_INDICATOR)
    for core in _PROJECT_CORE_STRONG:
        if text.endswith(core) and n - len(core) >= 2 and has_ind:
            return True
    for core in _PROJECT_CORE_WEAK:
        # 弱中心词需要 >=4 字专名修饰：「网络威胁感知子系统」✓ 「开放平台」✗
        if text.endswith(core) and n - len(core) >= 4:
            return True
    return False


def _ner_should_suppress(text: str, text_type_name: str) -> bool:
    """结构准入制：只有构词上具备专名特征的候选才允许脱敏。

    返回 True 表示该候选是普通词组 / 句子成分，应当原样保留。

    与旧版「误报词黑名单」的区别：这里不查任何词表，而是检查候选是否
    满足对应实体类型的构词规则，因此对没见过的普通词同样有效。
    """
    t = (text or "").strip()
    if len(t) <= 1:
        return True
    # 句子成分（以虚词 / 谓词开头）与抽象概念（以…性 /…度 /…化 收尾）一律排除
    if _starts_with_function_word(t) or _ends_with_abstract(t):
        return True

    if text_type_name in ("company", "government"):
        return not _looks_like_org(t)
    if text_type_name == "person":
        return not _looks_like_person(t)
    if text_type_name == "location":
        # 地名后缀，或本身就是带地名前缀的机构（北京市公安局）
        return not (_looks_like_geo(t) or _looks_like_org(t))
    if text_type_name == "project":
        return not _looks_like_project(t)
    # custom 及其它未知专名：须具备任一专名特征
    return not (_looks_like_org(t) or _looks_like_geo(t) or _looks_like_project(t))


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



# ══════════════════════════════════════════════════════════════════
# 词性定界：用汉语语法（虚词 / 谓词的分布规律）确定专名的左边界。
# 这一层不含任何实体词，纯粹依赖分词器的词性标注。
# ══════════════════════════════════════════════════════════════════

#: 「硬断词」——绝不可能出现在专名内部的词类，专名起点必在其右侧。
#:   p 介词(由/向/与/对)  r 代词(本/该/其)  d 副词(须/应/均)
#:   u 助词(的/地/得)     y 语气词  e 叹词  o 拟声词  w 标点
_POS_HARD: tuple[str, ...] = ("p", "r", "d", "u", "y", "e", "o", "w")

#: 「不可起头」——可以出现在专名内部，但不能作为专名的第一个词。
#:   c 连词(并/和/与)——「发展和改革委员会」内部合法，但不能起头
#:   v 动词——「行政执法监督局」内部合法，但「邀请中国…协会」不合法
#:   m/q 数量词、t 时间词、f 方位词
_POS_NO_START: tuple[str, ...] = ("c", "m", "q", "t", "f")


def _pos_tokens(text: str) -> list:
    """分词 + 词性标注（带缓存）。分词器不可用时退化为整体一个名词。"""
    cached = _POS_CACHE.get(text)
    if cached is not None:
        return cached
    try:
        import jieba.posseg as pseg
        toks = [(w.word, w.flag) for w in pseg.cut(text)]
    except Exception:
        toks = [(text, "n")]
    if len(_POS_CACHE) > 20000:
        _POS_CACHE.clear()
    _POS_CACHE[text] = toks
    return toks


_POS_CACHE: dict = {}


def _pos_is_hard(flag: str) -> bool:
    """是否为硬断词（专名内部不可能出现）。"""
    return bool(flag) and flag[0] in _POS_HARD


def _pos_is_soft(word: str, flag: str) -> bool:
    """是否为软断词：多字纯动词，多半是句子谓语而非专名内部语素。

    单字动词（中铁「建」、「新」建）常是专名构词成分，故不计入。
    动名词 vn（建设 / 咨询 / 监理 / 运营）在机构名中极常见，也不计入。
    """
    return len(word) >= 2 and flag == "v"


def _pos_can_start(flag: str) -> bool:
    """该词类能否作为专名的起始词。"""
    if not flag:
        return False
    if flag[0] in _POS_HARD:
        return False
    if flag[0] in _POS_NO_START:
        return False
    if flag[0] == "v" and not flag.startswith("vn"):
        return False
    return True


def _pos_seg_from(toks: list, i: int) -> str:
    """从第 i 个词起拼接为片段；若起始词不能起头则右移。"""
    n = len(toks)
    while i < n and not _pos_can_start(toks[i][1]):
        i += 1
    if i >= n:
        return ""
    return "".join(w for w, _ in toks[i:])


#: 置于专名前、明确表示其后的「机构 / 地名」是本句宾语（而非专名一部分）
#: 的谓语动词。这些是语法功能词，不是实体名，符合「构词法 + 语法」而非
#: 「实体库」的识别思路。
_PRED_BEFORE_ORG: tuple[str, ...] = (
    "委托", "交由", "交给", "发包", "承包", "中标", "提供", "承担",
    "通过", "经", "由", "向", "对", "就", "关于",
)


def _pos_org_head_shift(toks: list, k: int, cand: str, validate) -> str:
    """机构名惯以行政区划 / 国名起头。

    若起始词只是个普通名词，而其后出现地名（ns/nt）或行政层级词
    （中国 / 国家 / 全国…），且两者衔接词是明确的谓语动词
    （委托 / 提供 / 由 / 向…），说明前面的词其实是句子成分而非机构名的
    一部分（「检测工作委托」国家…检验中心），据此把起点右移到该地名 /
    层级词处。扫描整段 token，避免只看了紧邻的两三个词而错过深层机构名。
    """
    n = len(toks)
    # 找到第一个可作为机构名起头的地名 / 层级词
    idx = None
    for j in range(k, n):
        w, f = toks[j]
        if f in ("ns", "nt") or w in _ORG_HEAD:
            idx = j
            break
    if idx is None or idx == k:
        return cand
    prev = toks[idx - 1][0]
    if prev not in _PRED_BEFORE_ORG:
        return cand
    better = "".join(x for x, _ in toks[idx:])
    if better and validate(better):
        return better
    return cand


def _pos_left_trim(text: str, validate, org_head: bool = False) -> str:
    """剥离左侧句子成分，返回构词上成立的专名片段（失败返回空串）。

    策略：
      ① 硬断词（介词 / 代词 / 副词 / 助词）绝不出现在专名内部，
         据此确定搜索下界 lo；
      ② 候选起点按「最靠右的软断词之后」→「次靠右」→ … →「lo」
         逐级回退，取第一个构词成立的片段。
         这样「运营数据接入|昆山市大数据管理局」能在动词后正确断开，
         而「全线通信传输网络扩容项目」里的「扩容」因断开后不成立
         （只剩「项目」），会自动退回完整名称。
    """
    toks = _pos_tokens(text)
    n = len(toks)
    if n == 0:
        return ""

    lo = 0
    for i in range(n):
        if _pos_is_hard(toks[i][1]):
            lo = i + 1
    if lo >= n:
        return ""

    softs = [i + 1 for i in range(lo, n) if _pos_is_soft(toks[i][0], toks[i][1])]
    for k in list(reversed(softs)) + [lo]:
        cand = _pos_seg_from(toks, k)
        if cand and validate(cand):
            if org_head:
                cand = _pos_org_head_shift(toks, k, cand, validate)
            return cand
    return ""


class _PseudoMatch:
    """轻量匹配对象，只实现 mask-tool 用到的 group / start / end。"""

    __slots__ = ("_t", "_s", "_e")

    def __init__(self, text: str, start: int, end: int) -> None:
        self._t, self._s, self._e = text, start, end

    def group(self, *_a):  # noqa: ANN002
        return self._t

    def start(self, *_a):  # noqa: ANN002
        return self._s

    def end(self, *_a):  # noqa: ANN002
        return self._e


class _SmartPattern:
    """正则包装器：贪婪匹配定右边界，再用词性定左边界。

    mask-tool 只调用 ``pattern.finditer(text)`` 以及 match 的
    group/start/end，因此这里实现这三个接口即可无侵入接入。
    """

    __slots__ = ("_re", "_validate", "_trim", "_split", "_org_head")

    def __init__(self, pattern, validate, trim: bool = True,
                 split: bool = False, org_head: bool = False) -> None:
        self._re = pattern
        self._validate = validate
        self._trim = trim
        self._split = split
        self._org_head = org_head

    def _emit(self, raw: str, base: int, depth: int = 0):
        """对一次原始匹配做左边界修剪，必要时递归回收被丢弃的左半段。"""
        if not raw or depth > 3:
            return
        if not self._trim:
            if self._validate(raw):
                yield _PseudoMatch(raw, base, base + len(raw))
            return

        seg = _pos_left_trim(raw, self._validate, org_head=self._org_head)
        if not seg:
            return
        off = len(raw) - len(seg)
        # 被剥掉的左半段里可能还藏着另一个并列专名（A公司与B公司）
        if self._split and off > 0:
            head = raw[:off]
            for m in self._re.finditer(head):
                yield from self._emit(m.group(), base + m.start(), depth + 1)
        yield _PseudoMatch(seg, base + off, base + off + len(seg))

    def finditer(self, text: str):
        for m in self._re.finditer(text):
            yield from self._emit(m.group(), m.start())


def _patch_detector_rules() -> None:
    """给 mask-tool 的检测器补充「自主识别」正则（幂等）。

    这些正则完全基于中文构词法与文档结构，不依赖任何实体词库 / 映射表：
      · 组织机构名 —— 以组织形式后缀（有限公司 / 集团 / 研究院 / 管理局…）
        收尾的专名。可识别任意此前没见过的单位，且能拿到完整名称，
        不受 jieba 分词切分影响（NER 常把「中国移动通信集团有限公司」
        截断成「中国移动通信集团」）。
      · 详细地址 —— 行政区划链 + 街路号 / 园区 / 大厦。
      · 项目 / 工程名 —— 修饰语 + 项目类中心词；以及标题行中以
        「子系统 / 系统 / 平台」收尾的项目名。
      · 结构化字段 —— 泛化的「XX名称：」「XX单位：」「XX人：」标签模式。
        对标签「尾字」做定宽 lookbehind，因此「项目名称 / 采购人 /
        使用单位 / 最终用户 / 供应商 / 甲方 / 项目地点」等任意写法都能
        命中，无需逐个登记标签。
      · 金额 —— 阿拉伯数字与中文大写两种写法（原版仅识别带万/亿的）。
    """
    try:
        from mask_tool.core.detector import Detector
        from mask_tool.models.detection import DetectionType

        if getattr(Detector, "_patched_rules_by_engine", False):
            return

        _orig = Detector._build_regex_rules  # type: ignore[assignment]

        _org_strong = "|".join(_ORG_SUFFIX_STRONG)
        _org_medium = "|".join(_ORG_SUFFIX_MEDIUM)
        _org_any = "|".join(_ORG_SUFFIX_STRONG + _ORG_SUFFIX_MEDIUM)

        # ① 组织机构名。贪婪前缀确保拿到完整名称（不会像 NER 那样把
        #    「中国移动通信集团有限公司」截断成「中国移动通信集团」），
        #    左边界随后交给 _SmartPattern 按词性剥离。
        _org_re = _SmartPattern(
            re.compile(rf"[一-鿿]{{2,24}}(?:{_org_strong})", re.UNICODE),
            _looks_like_org, split=True, org_head=True,
        )
        # ②「弱后缀」机构（…监督局 /…管理厅 /…检验中心）：需更长专名前缀
        _org_re2 = _SmartPattern(
            re.compile(rf"[一-鿿]{{4,24}}(?:{_org_medium})", re.UNICODE),
            _looks_like_org, split=True, org_head=True,
        )

        # ③ 详细地址：行政区划链 + 街路号 / 园区 / 大厦（贪婪取最长）
        _addr_re = _SmartPattern(
            re.compile(
                r"(?:[一-鿿]{2,6}(?:省|自治区|特别行政区))?"
                r"[一-鿿]{2,6}(?:市|自治州|地区)"
                r"(?:[一-鿿]{2,6}(?:区|县|市|旗))?"
                r"[一-鿿]{0,16}"
                r"(?:软件园|科技园|工业园|开发区|高新区|园区|新区|大厦|广场|大道|路|街|巷)"
                r"(?:[0-9０-９]{1,4}号)?",
                re.UNICODE,
            ),
            _looks_like_geo,
        )

        # ④ 项目 / 工程名（强中心词）。校验函数要求内部含实体指示语素，
        #    因此「重要项目」「该工程」这类泛指不会命中。
        _proj_strong = "|".join(_PROJECT_CORE_STRONG)
        _project_re = _SmartPattern(
            re.compile(rf"[一-鿿]{{2,24}}?(?:{_proj_strong})", re.UNICODE),
            _looks_like_project,
        )

        # ⑤ 标题行项目名（弱中心词）：独立成行、无章节编号的短行，
        #    以「子系统 / 系统 / 平台」收尾。技术规范书封面与大标题里的
        #    项目名多是这种形态；正文中的「入侵检测系统」不独立成行，
        #    因此不会被命中。该规则本身已由行结构锚定，无需词性修剪。
        _proj_weak = "|".join(_PROJECT_CORE_WEAK)
        _func = "|".join(_FUNC_HEAD)
        _title_re = _SmartPattern(
            re.compile(
                r"(?<=\n)[ \t]*"
                r"(?![0-9０-９一二三四五六七八九十]{1,3}[\.、．　 ])"
                rf"(?!{_func})"
                r"[一-鿿A-Za-z0-9（）()·\-]{4,32}?"
                rf"(?:{_proj_weak})"
                r"[ \t]*(?=\n)",
                re.UNICODE,
            ),
            lambda s: True, trim=False,
        )

        # ⑥ 结构化字段值。标签本身即强语境，值按「遇标点即止」取整。
        _val = r"[一-鿿A-Za-z0-9·（）()\.\-]{2,40}"
        _val_org = rf"[一-鿿A-Za-z0-9·（）()\.\-]{{2,30}}?(?:{_org_any})"
        _val_person = r"[一-鿿]{2,4}(?![一-鿿])"

        def _lab(tails, val_re):
            """构造「标签尾字 + 冒号」→ 值 的定宽 lookbehind 正则。

            对标签「尾字」而非整个标签做匹配，因此「项目名称 / 采购人 /
            使用单位 / 最终用户 / 供应商 / 甲方 / 项目地点」等任意写法
            都能命中，无需逐个登记标签。
            """
            alts = "|".join(f"(?<={t}[:：])" for t in tails)
            return re.compile(f"(?:{alts})[ \t　]*{val_re}", re.UNICODE)

        # —— 置信度校准 ——
        # 上游给「确定性模式」的分数偏低（邮箱 0.75 / 日期 0.70），在
        # 0.8 的全局下限下会被连带误滤。这里按规则的**真实精度**重新标定：
        # 正则匹配到的邮箱、手机号、身份证几乎不可能是别的东西。
        # 反之，银行卡规则 \d{16,19} 会命中任意长数字串（订单号、编号），
        # 精度确实低，保持 0.65 —— 在 0.8 下限下自动退化为「仅提示」，
        # 这正是本次改造想要的效果。
        _recal: dict[str, float] = {
            r"[\w.-]+@[\w.-]+\.\w+": 0.95,        # 邮箱
            r"1[3-9]\d{9}": 0.95,                 # 手机号
            r"\d{17}[\dXx]": 0.92,                # 身份证号
            r"\d{4}年\d{1,2}月\d{1,2}日": 0.85,   # 中文日期
        }

        def _build(self):  # noqa: ANN001
            rules = [
                (rx, dt, _recal.get(getattr(rx, "pattern", ""), cf))
                for rx, dt, cf in _orig(self)
            ]
            extra = [
                # 金额：阿拉伯数字（含「元」但不带万 / 亿，如 1,200,000元）
                (re.compile(r"[\d,]+\.?\d*元"), DetectionType.AMOUNT, 0.80),
                # 金额：中文大写（陆拾万元 / 壹佰万元 / 伍仟元）。
                # 前导数字不含万 / 亿，且整体须以 元 / 圆 / 万 / 亿 收尾，
                # 避免把「一个」「十」之类普通词误判为金额。
                (re.compile(
                    r"[零〇一二三四五六七八九十百千壹贰叁肆伍陆柒捌玖拾佰仟]+"
                    r"(?:[万亿]?元|[万亿]?圆|万|亿)"),
                 DetectionType.AMOUNT, 0.85),
                # 组织机构名（自主识别，可命中任意陌生单位）
                (_org_re, DetectionType.COMPANY, 0.92),
                (_org_re2, DetectionType.COMPANY, 0.90),
                # 详细地址
                (_addr_re, DetectionType.LOCATION, 0.88),
                # 项目 / 工程名
                (_project_re, DetectionType.PROJECT, 0.85),
                # 标题行项目名：由「独立成行 + 无章节号 + 非虚词开头」三重
                # 结构锚定，精度接近强中心词规则，标定 0.82（高于 0.8 下限）
                (_title_re, DetectionType.PROJECT, 0.82),
                # 结构化编号：格式确定性高，直接标定高置信（均高于各类型下限）
                # 项目编号：AB12-2024-001
                (re.compile(r"[A-Z]{2,4}-\d{4}-\d{3,}"), DetectionType.PROJECT, 0.90),
                # 合同编号：HT2024-0001 / HT_2024_0001
                (re.compile(r"HT[-_]?\d{4}[-_]?\d{4,}"), DetectionType.CUSTOM, 0.90),
                # 订单号：ORD1234567
                (re.compile(r"ORD[-_]?\d{6,}"), DetectionType.CUSTOM, 0.85),
            ]

            # —— 结构化字段：泛化「标签 → 值」模式（不枚举具体标签）——
            extra += [
                # 「…名称：」值带组织后缀 → 机构；否则 → 项目名
                (_lab(["名称"], _val_org), DetectionType.COMPANY, 0.92),
                (_lab(["名称"], _val), DetectionType.PROJECT, 0.90),
                # 「…单位 / 公司 / 商 / 机构 / 客户 / 用户 / 业主 /
                #   部门 / 厂商 / 方：」→ 机构
                (_lab(["单位", "公司", "商", "机构", "客户", "用户", "业主",
                       "部门", "厂商", "甲方", "乙方", "丙方"],
                      _val), DetectionType.COMPANY, 0.90),
                # 「…地点 / 地址：」→ 地点
                (_lab(["地点", "地址"], _val), DetectionType.LOCATION, 0.90),
                # 「…人：」→ 人名（构词校验，排除机构 / 地名 / 抽象概念）
                (_SmartPattern(_lab(["人"], _val_person),
                               _looks_like_person_labeled, trim=False),
                 DetectionType.PERSON, 0.88),
            ]
            return rules + extra

        Detector._build_regex_rules = _build  # type: ignore[assignment]
        Detector._patched_rules_by_engine = True  # type: ignore[attr-defined]
    except Exception:
        pass


def _patch_detector_dedup() -> None:
    """检测结果子串去重（幂等，置信度感知）。

    自主识别的多条规则常对同一实体给出长短不一的片段
    （「中铁建电气化局」「工程有限公司」「中铁建电气化局集团第三工程
    有限公司」）。这里只保留最可信的那条，避免映射表被碎片刷屏。

    **置信度感知**（关键修复）：长短片段重叠时，保留**置信度更高**的
    一方；置信度相等才保留更长（完整名称优先）。

    为什么必须如此？以 spaCy 后端为例：它可能把
    「西南交通大学轨道交通运载系统全国重点实验室」整体识别为一个低置信
    实体，而结构化规则同时精确命中其中的「西南交通大学」（0.92）。若仍
    按「仅保留最长」的旧逻辑，低置信长片段会先把精确的短片段吞掉，随后
    长片段又因低于置信度下限被策略层丢弃 —— 最终「西南交通大学」凭空
    消失。改为「高置信优先」后，精确短片段得以保留。

    用户词库（``source == "dictionary"``）命中条目永远保留、且永远胜出，
    不受长短与置信度影响。
    """
    try:
        from mask_tool.core.detector import Detector

        if getattr(Detector, "_patched_dedup_by_engine", False):
            return

        _orig_detect = Detector.detect  # type: ignore[assignment]

        def _conf_of(r):
            try:
                c = getattr(r, "confidence", None)
                return float(c) if c is not None else 0.0
            except (TypeError, ValueError):
                return 0.0

        def detect(self, text, file_path=""):  # noqa: ANN001
            results = _orig_detect(self, text, file_path)
            if not results:
                return results

            # 先做「策略层预筛」：剔除注定会被置信度下限 / 白名单丢弃的
            # 低质量候选，再对幸存者做子串去重。否则低质长片段会先把精确
            # 短片段吞掉，随后自己又被下限丢弃，导致精确短片段凭空消失。
            # 典型场景：spaCy 把「西南交通大学轨道交通运载系统全国重点
            # 实验室」整体识别为低置信实体，吞掉其中精确的「西南交通大学」。
            survivors = [
                r for r in results
                if getattr(r, "source", "") == "dictionary"
                or (_conf_of(r) >= _min_conf_for_type(_type_name_of(r))
                     and not is_whitelisted(r.text))
            ]

            kept_texts: list[str] = []
            drop: set[str] = set()
            for t in sorted({r.text for r in survivors}, key=len, reverse=True):
                if any(t in o for o in kept_texts):
                    drop.add(t)
                else:
                    kept_texts.append(t)
            if not drop:
                return survivors
            out = [r for r in survivors
                   if r.text not in drop
                   or getattr(r, "source", "") == "dictionary"]
            return out

        Detector.detect = detect  # type: ignore[assignment]
        Detector._patched_dedup_by_engine = True  # type: ignore[attr-defined]
    except Exception:
        pass


#: NER 后端配置（由 UI / 配置在每次处理前设置）
#: 默认后端切换为 spaCy，默认模型优先 zh_core_web_md（未安装时回退到内置
#: jieba，仍开箱即用；详见 discover_spacy_model 的兜底）。
_NER_BACKEND: str = "spacy"
_NER_MODEL: str = "zh_core_web_md"


def set_ner_backend(backend: str | None, model: str | None = "") -> None:
    """设置 NER 后端。

    Args:
        backend: ``auto`` / ``spacy`` / ``jieba``
        model:   spaCy 模型路径或包名；留空则自动发现
    """
    global _NER_BACKEND, _NER_MODEL
    b = (backend or "auto").strip().lower()
    _NER_BACKEND = b if b in ("auto", "spacy", "jieba") else "auto"
    _NER_MODEL = (model or "").strip()


def _load_ner_backend_module():
    """导入可插拔后端模块（兼容包内 / 脚本两种引用方式）。"""
    try:
        from app import ner_backend as _nb
        return _nb
    except Exception:
        try:
            import ner_backend as _nb  # type: ignore[no-redef]
            return _nb
        except Exception:
            return None


def ner_status() -> dict:
    """返回当前 NER 后端状态（供 UI 与 --selftest 展示）。"""
    nb = _load_ner_backend_module()
    if nb is None:
        return {"configured": _NER_BACKEND, "active": "jieba",
                "spacy_installed": False, "model": "",
                "reason": "未找到 ner_backend 模块，使用 jieba"}
    try:
        return nb.backend_status(_NER_BACKEND, _NER_MODEL or None, app_root())
    except Exception as exc:
        return {"configured": _NER_BACKEND, "active": "jieba",
                "spacy_installed": False, "model": "",
                "reason": f"后端探测失败：{exc}"}


def _patch_ner_backend() -> None:
    """把 mask-tool 写死的 jieba NER 换成可插拔后端（幂等）。

    上游 ``Pipeline.__init__`` 里是
    ``from mask_tool.core.ner.jieba_ner import JiebaNER; ner_engine = JiebaNER()``，
    属于**调用时导入**，因此替换模块属性即可生效，无需改动上游源码。

    任何一步失败都退回原生 jieba —— 换引擎是增强项，不能成为故障点。
    """
    try:
        from mask_tool.core.ner import jieba_ner as _jn

        if getattr(_jn, "_patched_backend_by_engine", False):
            return

        _OrigJieba = _jn.JiebaNER

        def _factory(*args, **kwargs):
            nb = _load_ner_backend_module()
            if nb is not None:
                try:
                    eng = nb.create_ner_engine(
                        _NER_BACKEND, _NER_MODEL or None, app_root()
                    )
                    if eng is not None:
                        return eng
                except Exception:
                    pass
            return _OrigJieba(*args, **kwargs)

        _jn.JiebaNER = _factory  # type: ignore[assignment]
        _jn._patched_backend_by_engine = True  # type: ignore[attr-defined]
    except Exception:
        pass


def _patch_policy() -> None:
    """对 mask-tool 的 PolicyEngine 打补丁（幂等）。

    对**所有模式**生效的两道硬闸门（顺序即优先级）：
      ① 白名单：``whitelist.txt`` 里的词绝不替换，无视来源与置信度；
      ② 置信度下限：``confidence < MIN_CONFIDENCE``（默认 0.8）只提示不替换。

    其后仅对 smart / aggressive 模式生效：
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
            # ① 白名单优先级最高：用户明确声明「绝不脱敏」的业务术语
            #    （二次开发 / 许可 / 卖方 / 买方 / 工程 …），
            #    无论来自词库、正则还是 NER，一律保留原文。
            if is_whitelisted(result.text):
                return DetectionStatus.HINT_ONLY

            # ② 置信度下限（按实体类型）：把握不大的识别只在报告里提示，不动正文
            _name = _type_name_of(result)
            if float(getattr(result, "confidence", 0.0) or 0.0) < _min_conf_for_type(_name):
                return DetectionStatus.HINT_ONLY

            conf = result.confidence
            src = getattr(result, "source", "") or ""
            mode = self.config.mode

            # strict（严格）：仅「用户词库 / 基准词库」（dictionary 来源）参与脱敏；
            # 正则与 NER 命中一律只提示、不替换，最大限度避免自动误伤。
            if mode == "strict":
                if src == "dictionary":
                    return DetectionStatus.AUTO_MASK if conf >= 0.95 \
                        else DetectionStatus.HINT_ONLY
                return DetectionStatus.HINT_ONLY

            if mode in ("smart", "aggressive"):
                if src == "dictionary":
                    # 用户词库 / 基准词库：永远脱敏
                    return DetectionStatus.AUTO_MASK if conf >= 0.95 \
                        else DetectionStatus.HINT_ONLY
                if src == "regex":
                    # 高置信正则（手机/身份证/邮箱/金额/IP/MAC/日期/中文金额/
                    # 项目名/项目编号/合同编号/订单号）始终脱敏
                    if conf >= _REGEX_AUTO:
                        return DetectionStatus.AUTO_MASK
                    if conf >= _REGEX_SUGGEST:
                        return DetectionStatus.SUGGEST_MASK
                    return DetectionStatus.HINT_ONLY
                # NER（及未知来源）：先强制过滤明显误报，再按灵敏度裁剪
                name = _name or getattr(result.text_type, "value", str(result.text_type)).lower()
                if _ner_should_suppress(result.text, name):
                    return DetectionStatus.HINT_ONLY
                thr = _AGGRESSIVE_THR if mode == "aggressive" else _SMART_THR
                if name in ("company", "government"):
                    t = thr["company"]
                elif name == "person":
                    t = thr["person"]
                elif name == "location":
                    t = thr["location"]
                else:
                    t = thr["custom"]
                verdict = DetectionStatus.AUTO_MASK if conf >= t else DetectionStatus.HINT_ONLY
                return verdict
            # focused 等其它模式沿用上游原始逻辑
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

    # 最终回退：mask-tool 已在当前进程内可导入（如开发环境从装好 mask_tool
    # 的 venv 直接跑 python main.py）。此时走进程内调用，无需外部命令。
    try:
        import importlib
        importlib.import_module("mask_tool")
        return ToolInfo(["(内嵌进程调用)"], "已安装 mask_tool（进程内调用）", "")
    except Exception:
        pass

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
        # 白名单：内置 whitelist.yaml + 用户 whitelist.txt 合并后的临时 YAML。
        # 懒构建 + 实例级缓存，避免批量处理时每个文件都建一个临时目录。
        self._whitelist_file: str | None = None
        self._whitelist_built = False

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

    # -- 白名单 -----------------------------------------------------------

    def refresh_whitelist(self) -> str | None:
        """重新读取 whitelist.txt 并重建合并白名单，下一次处理即生效。"""
        self._whitelist_file = self._build_merged_whitelist()
        self._whitelist_built = True
        return self._whitelist_file

    def _get_whitelist_file(self) -> str | None:
        """取合并后的白名单文件路径（首次调用时构建）。"""
        if not self._whitelist_built:
            self.refresh_whitelist()
        return self._whitelist_file

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
            _patch_detector_dedup()
            # 可插拔 NER 后端（spaCy / jieba），未配置时行为不变
            _patch_ner_backend()
            # 注入「白名单 + 置信度下限 + 检测灵敏度」策略
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

    @staticmethod
    def _build_merged_whitelist() -> str | None:
        """把内置 whitelist.yaml 与用户的 whitelist.txt 合并，写出临时 YAML。

        为什么要合并到 mask-tool 自己的白名单里，而不是只在策略层拦截？
        因为 mask-tool 在**更早的检测阶段**就会用它过滤
        （``Detector`` 的正则/词库匹配、``JiebaNER.recognize``），
        白名单词根本不会进入候选列表 —— 既更省事也更彻底。
        策略层的 :func:`is_whitelisted` 是第二道保险，负责兜住那些
        经过左边界修剪、去重后才成形的片段。
        """
        words: set[str] = set()

        cfg_dir = _bundled_config_dir()
        if cfg_dir is not None:
            wlt = cfg_dir / "whitelist.yaml"
            if wlt.is_file():
                try:
                    import yaml
                    loaded = yaml.safe_load(wlt.read_text(encoding="utf-8"))
                    if isinstance(loaded, dict):
                        base = loaded.get("whitelist")
                        if isinstance(base, list):
                            words |= {w for w in base if isinstance(w, str) and w.strip()}
                    elif isinstance(loaded, list):
                        words |= {w for w in loaded if isinstance(w, str) and w.strip()}
                except Exception:
                    pass

        # 用户可编辑的 whitelist.txt（程序根目录）
        words |= set(load_whitelist(force=True))

        if not words:
            return None

        tmp_dir = tempfile.mkdtemp(prefix="masktool_wl_")
        out = Path(tmp_dir) / "whitelist.yaml"
        try:
            import yaml
            out.write_text(
                yaml.dump({"whitelist": sorted(words)},
                          allow_unicode=True, default_flow_style=False),
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
        min_confidence: float | None = None,
        ner_backend: str | None = None,
        spacy_model: str | None = None,
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
                src, out_dir, mode, save_mapping, suffix_tag,
                sensitivity, min_confidence, ner_backend, spacy_model,
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
        min_confidence: float | None = None,
        ner_backend: str | None = None,
        spacy_model: str | None = None,
    ) -> FileResult:
        """进程内调用 mask-tool 的 Pipeline（冻结打包后的主路径）。"""
        ext = src.suffix.lower()
        report_only = ext in REPORT_ONLY_EXTS

        # 应用用户选择的 smart/aggressive 检测灵敏度（仅对 NER 生效）
        set_smart_sensitivity(sensitivity)
        # 应用全局置信度下限（对所有来源生效；None 表示沿用当前值）
        if min_confidence is not None:
            set_min_confidence(min_confidence)
        # 应用 NER 后端选择（auto/spacy/jieba；None 表示沿用当前值）。
        # _patch_ner_backend 注入的工厂在每次实例化时读取这两个全局量，
        # 因此此处设置会在本文件处理时立即生效。
        if ner_backend is not None:
            set_ner_backend(ner_backend, spacy_model)

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

        # 白名单：内置 whitelist.yaml + 用户 whitelist.txt 合并后的临时文件。
        # 合并失败（极端情况）时回退到打包资源，保证冻结后也能加载。
        merged_wl = self._get_whitelist_file()
        if merged_wl:
            cfg.whitelist_path = merged_wl
        elif cfg_dir is not None:
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
        min_confidence: float | None = None,
        ner_backend: str | None = None,
        spacy_model: str | None = None,
        on_progress: Callable[[int, int, Path], None] | None = None,
        on_result: Callable[[FileResult], None] | None = None,
    ) -> BatchOutcome:
        """批量处理。``out_dir_for`` 用于为每个源文件决定输出目录。

        白名单与置信度下限在**整批开始前统一装载一次**，因此批量处理中的
        每个文件都适用同一套规则，不会出现「第一个文件生效、后面失效」。
        """
        self.reset()
        items = list(files)
        outcome = BatchOutcome()

        # —— 整批统一装载规则 ——
        # 重新读取 whitelist.txt（用户可能刚改完就点了开始），并把
        # 合并后的白名单重新注入引擎；置信度下限与 NER 后端同理。
        if min_confidence is not None:
            set_min_confidence(min_confidence)
        if ner_backend is not None:
            set_ner_backend(ner_backend, spacy_model)
        load_whitelist(force=True)
        self.refresh_whitelist()

        for idx, src in enumerate(items, start=1):
            if self._cancelled:
                outcome.cancelled = True
                break
            if on_progress:
                on_progress(idx, len(items), src)

            res = self.process_one(
                src, out_dir_for(src), mode=mode,
                save_mapping=save_mapping, suffix_tag=suffix_tag,
                sensitivity=sensitivity, min_confidence=min_confidence,
                ner_backend=ner_backend, spacy_model=spacy_model,
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
