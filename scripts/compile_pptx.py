#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt-visual-art-director · 确定性 PPTX 编译器

职责：消费符合 plan.schema.json 的 slide plan，编译为原生可编辑 PPTX，并在写文件前执行
机器可判定的契约检查（结构、几何、引用、图表数据与口径声明、配色数量、素材登记），以及
主张—数据数值筛查（E008/E016）与文本装载估算（E010）。

设计边界：
- 不决定叙事、构图、风格或表达取舍；忠实执行已确定的设计意图。
- 契约失败在创建 PPTX 前返回，并给出位置、条件、实际结果与级别。
- 估算与筛查只报风险，不冒充真实字体测量或事实核验；不静默改写设计。

用法：
    python scripts/compile_pptx.py plan.json out.pptx [--assets .] [--strict]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from datetime import date
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

try:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION, XL_TICK_MARK
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt
except ImportError as exc:  # pragma: no cover - 依赖缺失时给出清晰提示
    print(f"[FATAL] 缺少依赖 python-pptx：{exc}\n        请先安装：pip install python-pptx", file=sys.stderr)
    sys.exit(2)

try:
    from jsonschema import Draft202012Validator
except ImportError as exc:  # pragma: no cover
    print(f"[FATAL] 缺少依赖 jsonschema：{exc}\n        请先安装：pip install jsonschema", file=sys.stderr)
    sys.exit(2)


def _load_authority(filename: str) -> Dict[str, Any]:
    """读取随包分发的唯一机器权威源；缺失或损坏时不得静默回退。"""
    path = Path(__file__).resolve().with_name(filename)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:  # pragma: no cover - 安装损坏
        raise RuntimeError(f"机器权威源不可读：{path}: {exc}") from exc
    if not isinstance(data, dict):
        raise RuntimeError(f"机器权威源顶层必须是对象：{path}")
    return data


PLAN_SCHEMA = _load_authority("plan.schema.json")
POLICY = _load_authority("policy.json")
Draft202012Validator.check_schema(PLAN_SCHEMA)
PLAN_VALIDATOR = Draft202012Validator(PLAN_SCHEMA)


# ---------------------------------------------------------------------------
# 错误模型
# ---------------------------------------------------------------------------

class ErrorLevel(Enum):
    HARD = "HARD"      # 阻断执行路径
    SOFT = "SOFT"      # 复核提示；strict 下须处置
    INFO = "INFO"      # 确定性默认值披露；不冒充风险或通过证据


@dataclass
class CompileError:
    code: str
    level: ErrorLevel
    location: str        # 页面 id / 元素 id / 字段
    condition: str       # 触发条件
    actual: str          # 实际结果
    message: str         # 人类可读说明

    def __str__(self) -> str:
        return (
            f"[{self.code}] {self.level.value} @ {self.location} | "
            f"条件：{self.condition} | 实际：{self.actual} | {self.message}"
        )


def unresolved_reviews(plan: Dict[str, Any], errors: List[CompileError]) -> List[CompileError]:
    """只允许对 policy 指定的经验提示留具名复核；硬错与语义/合规不可豁免。"""
    records = plan.get("design_reviews", [])
    return [error for error in errors if error.level == ErrorLevel.SOFT and not (
        error.code in POLICY["diagnostics"]["reviewable_codes"] and any(
            record["code"] == error.code and record["location"] == error.location
            for record in records
        )
    )]


# ---------------------------------------------------------------------------
# 机器契约与策略视图
# Schema 管字段/枚举；policy.json 管默认值/阈值。这里只派生只读视图，不另存副本。
# ---------------------------------------------------------------------------

_CANVAS = POLICY["canvas"]
BOUND_TOL_IN = _CANVAS["bound_tolerance_in"]

DEFAULT_FONTS = POLICY["defaults"]["fonts"]
MIN_PT_SETS = POLICY["typography"]["min_pt_sets"]
MIN_PT = {etype: size for sizes in MIN_PT_SETS.values() for etype, size in sizes.items()}
# 展示角色（大题/巨数）由 policy 的字号角色分档唯一定义：它同时决定缺省字重、字族角色与行距档。
DISPLAY_TEXT_TYPES = frozenset(MIN_PT_SETS["display"])

VALID_ASPECT_RATIOS = set(PLAN_SCHEMA["properties"]["aspect_ratio"]["enum"])
TEXT_ELEMENT_TYPES = set(PLAN_SCHEMA["$defs"]["textElementType"]["enum"])


def _check_policy_consistency() -> None:
    """policy 自身的一致性：图表辅助文字（图例/坐标轴/直标/表头/表体）不得低于
    information 角色的字号下限——两个值都在 policy 里，关系在此校验，不另立第二套规则。
    """
    floor = min(POLICY["typography"]["min_pt_sets"]["information"].values())
    chart_pts = {key: float(value) for key, value in POLICY["chart"]["font_pt"].items()}
    too_small = {key: value for key, value in chart_pts.items() if value < float(floor)}
    if too_small:
        raise RuntimeError(
            f"policy.json 不一致：图表辅助文字 {too_small} 低于 information 角色下限 {floor}pt；"
            "修正 policy 的取值，不要在代码里另设一套字号规则")
    for role, sizes in POLICY["typography"]["default_pt"].items():
        if role in MIN_PT and float(sizes) < float(MIN_PT[role]):
            raise RuntimeError(f"policy.json 不一致：{role} 缺省字号 {sizes} 低于其角色下限 {MIN_PT[role]}pt")


_check_policy_consistency()


def min_pt_for(etype: str) -> float:
    """字号下限只按字体角色分档（展示/阅读/信息），不为每个元素名各发明一套规则。"""
    return MIN_PT[etype]


def text_style(el: Dict[str, Any], fonts: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
    """元素级文本式样的唯一解析口径：编译写入、交付核对、装载估算与对比度计算共用。

    缺省值只来自 policy；任何一处自行推导默认值都会让\"写进去的\"与\"核对的\"漂移。
    """
    etype = el["type"]
    display = etype in DISPLAY_TEXT_TYPES
    role = "heading" if display else "body"
    fonts = fonts or {}
    size = el.get("font_size")
    return {
        "size": float(POLICY["typography"]["default_pt"][etype] if size is None else size),
        "bold": bool(el["bold"]) if "bold" in el else display,
        "font": el.get("font") or fonts.get(role) or DEFAULT_FONTS[role],
        "color": str(el.get("color") or POLICY["defaults"]["text_color"]).upper(),
        "line_spacing": float(el.get("line_spacing") or POLICY["typography"]["line_spacing"]["display" if display else "body"]),
        "letter_spacing_pt": el.get("letter_spacing_pt"),
        "valign": el.get("valign") or POLICY["defaults"]["element"]["valign"],
    }


def run_styles(el: Dict[str, Any], fonts: Optional[Dict[str, str]] = None) -> List[Dict[str, Any]]:
    """逐片段的实际式样（无 runs 时即元素本身）：未声明的属性继承元素级解析结果。"""
    base = text_style(el, fonts)
    specs = el.get("runs") or [{"text": el.get("content") or ""}]
    out = []
    for spec in specs:
        style = dict(base)
        style["text"] = str(spec.get("text", ""))
        for key, field in (("size", "font_size"), ("font", "font"), ("bold", "bold"), ("letter_spacing_pt", "letter_spacing_pt")):
            if spec.get(field) is not None:
                style[key] = float(spec[field]) if key == "size" else spec[field]
        if spec.get("color"):
            style["color"] = str(spec["color"]).upper()
        out.append(style)
    return out


def _slide_dimensions(aspect_ratio: Optional[str]) -> Tuple[float, float]:
    """只接受 Schema 已声明的画幅；缺省值显式来自 policy。"""
    selected = aspect_ratio or _CANVAS["default_aspect_ratio"]
    if selected not in VALID_ASPECT_RATIOS:
        raise ValueError(f"未定义的画幅：{selected!r}")
    return tuple(_CANVAS["dimensions_in"][selected])


# ---------------------------------------------------------------------------
# 数值解析与图表口径推导
# ---------------------------------------------------------------------------

_NUM_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9#])(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?")


def _extract_numbers(text: str) -> List[float]:
    """提取文本中的定量数值（兼容千分位逗号，忽略 Q1/H2/#HEX 等标识符前缀数字）。"""
    if not isinstance(text, str):
        return []
    out: List[float] = []
    for m in _NUM_TOKEN_RE.finditer(text):
        try:
            out.append(float(m.group(0).replace(",", "")))
        except ValueError:
            continue
    return out


# ---------------------------------------------------------------------------
# 主张—数据数值筛查（E008 矛盾 / E016 未能核对）
#
# 设计边界：这是**数值一致性筛查**，不是事实核验器。它只回答\"页面可见的定量表述能否在
# 本页图表/表格数据中找到方向、量纲、数量级、期间一致的对应来源\"。因此：
#   · 量纲（base）、数量级（scale）、存储值（value）、显示值（主张按其精度四舍五入）分开建模；
#   · 只有经确认的换算才参与比较：同量纲按数量级换算；量纲未知、无确认换算关系一律不判一致；
#   · 原始值与派生值分开；派生值若依赖数据本身无法证明的假设（可加性、类别先后、系列互斥），
#     命中只记\"需确认假设\"，不记已验证；
#   · 与本页数据直接矛盾（方向相反、期间不符、同量纲数量级错位）记 E008；无法核对记 E016；
#   · 每一项数值独立判定，任一数字碰巧匹配不得使整句通过。
# ---------------------------------------------------------------------------

# 数量级：唯一换算表（长前缀优先）。正则全部由它派生，不另写第二份字面列表。
_SCALE_PREFIXES = (
    ("万亿", 1e12), ("千亿", 1e11), ("百亿", 1e10), ("十亿", 1e9), ("亿", 1e8),
    ("千万", 1e7), ("百万", 1e6), ("十万", 1e5), ("万", 1e4), ("千", 1e3), ("百", 1e2),
)
_SCALE_FACTORS = dict(_SCALE_PREFIXES)
_SCALE_ALT = "|".join(re.escape(prefix) for prefix, _ in _SCALE_PREFIXES)
# 已确认量纲的别名（季度/月/年等归入期间标记，不参与数值匹配）
_BASE_ALIASES = {
    "元": "元", "块": "元", "人民币": "元", "rmb": "元", "cny": "元",
    "美元": "美元", "美金": "美元", "usd": "美元",
    "港元": "港元", "港币": "港元", "欧元": "欧元", "eur": "欧元",
    "人": "人", "位": "人", "名": "人", "员工": "人", "人次": "人次", "台次": "台次", "架次": "架次",
    "件": "件", "台": "台", "套": "套", "个": "个", "条": "条", "次": "次",
    "单": "单", "户": "户", "家": "家", "辆": "辆", "份": "份", "张": "张", "例": "例",
    "吨": "吨", "平方米": "平方米", "㎡": "平方米",
    # 常见商品/物量计量：保守收录（“千克/千米”由“千”+“克/米”解析，不另设别名）。“部/项/场/间/座”等易与“部门/项目/场景”等普通词冲突，故意不收。
    "箱": "箱", "瓶": "瓶", "盒": "盒", "袋": "袋", "罐": "罐", "桶": "桶",
    "公斤": "公斤", "斤": "斤", "克": "克", "毫升": "毫升", "升": "升",
    "公里": "公里", "米": "米", "公顷": "公顷", "亩": "亩",
    "千瓦时": "千瓦时", "兆瓦时": "兆瓦时",
    "gb": "GB", "tb": "TB", "mb": "MB", "ms": "ms",
    "天": "天", "日": "天", "小时": "小时", "分钟": "分钟", "倍": "倍",
}
_BASE_UNITS = sorted(_BASE_ALIASES, key=len, reverse=True)
_UNIT_ALT = "|".join(re.escape(u) for u in _BASE_UNITS)
_PCT_UNITS = {"%", "％"}
_PCT_POINT_UNITS = {"个百分点", "百分点", "pct", "pp", "个点", "点"}

_PCT_SUFFIX_RE = re.compile(r"\s*(%|％|个百分点|百分点|pct|pp)")
_SCALED_UNIT_RE = re.compile(r"\s*((?:" + _SCALE_ALT + r")?(?:" + _UNIT_ALT + r"))(?![A-Za-z])", re.IGNORECASE)
# 未登记的拉丁单位（MWh、GW、km…）按原文身份读取：同名可比，异名无确认换算；期间记号不是单位。
_LATIN_UNIT_RE = re.compile(r"\s?((?:" + _SCALE_ALT + r")?(?![QqHh][1-4](?![A-Za-z0-9]))[A-Za-z][A-Za-z0-9/²³]{0,7})(?![A-Za-z0-9])")
_BARE_SCALE_RE = re.compile(r"\s*(" + _SCALE_ALT + r")")

_UP_CHANGE_RE = re.compile(r"(增长|上升|上涨|提升|增加|扩大|提高|攀升|翻倍|翻番|走高)")
_DOWN_CHANGE_RE = re.compile(r"(下降|下滑|下跌|降低|减少|回落|收窄|萎缩|走低)")
_UP_TO_RE = re.compile(r"(增至|升至|涨至|提高到|扩大到|增加到|上升至|增长至)")
_DOWN_TO_RE = re.compile(r"(降至|跌至|下降到|降低到|减少到|下滑至|回落至)")
_COMPARATIVE_RE = re.compile(r"(同比|环比|较上|较去年|较前|相比|对比|年比)")

# 支持 Q3 / 26Q3 / 2026Q3 三种常见写法；年份只作附加限定，不单独构成期间
_PERIOD_Q_RE = re.compile(r"(?<![A-Za-z0-9])(?:(20\d{2}|[1-9]\d)Q|Q)([1-4])(?![0-9])", re.IGNORECASE)
_PERIOD_H_RE = re.compile(r"(?<![A-Za-z0-9])(?:(20\d{2}|[1-9]\d)\s*)?H([12])", re.IGNORECASE)
_PERIOD_CN_Q_RE = re.compile(r"第?([一二三四])季度")
_PERIOD_MONTH_RE = re.compile(r"(?<!\d)(1[0-2]|[1-9])\s?月")  # 中文排版常在数字与汉字间留空格

_CN_QUARTER = {"一": "1", "二": "2", "三": "3", "四": "4"}


@dataclass(frozen=True)
class Unit:
    """单位 = 数量级 × 量纲。量纲为空表示只知数量级（如裸写的\"万\"），不知道是元、人还是件。"""

    scale: float
    base: str
    known: bool     # 量纲来自已确认别名表（或百分号/百分点）；未登记的原文单位只按身份比较
    text: str = ""


def _parse_unit(unit: Optional[str]) -> Unit:
    """'亿元' → (1e8,'元',known)、'万' → (1e4,'',-)、'MWh' → (1,'MWh',unknown)、'' → 缺失。"""
    u = (unit or "").strip().replace(" ", "")
    if not u:
        return Unit(1.0, "", False, "")
    if u in _PCT_UNITS:
        return Unit(1.0, "%", True, u)
    if u.lower() in _PCT_POINT_UNITS:
        return Unit(1.0, "pct", True, u)
    if u in _SCALE_FACTORS:
        return Unit(_SCALE_FACTORS[u], "", False, u)
    scale, rest = 1.0, u
    for prefix, factor in _SCALE_PREFIXES:
        if u.startswith(prefix) and len(u) > len(prefix):
            scale, rest = factor, u[len(prefix):]
            break
    canonical = _BASE_ALIASES.get(rest.lower()) or _BASE_ALIASES.get(rest)
    if canonical:
        return Unit(scale, canonical, True, u)
    return Unit(scale, rest, False, u)


def _unit_relation(a: Unit, b: Unit) -> str:
    """两侧单位的可比关系：same | missing | unknown_dimension | incompatible | unconfirmed。

    只有 same（同一量纲，数量级可换算）允许判定一致。一侧缺失或只知数量级，不得因
    另一侧量纲已知就当作可比；两个已知量纲不同是 incompatible；涉及未登记单位且
    原文不同是 unconfirmed——没有确认的换算关系，既不能判一致，也不能断言矛盾。
    """
    if not a.text or not b.text:
        return "missing"
    if not a.base or not b.base:
        return "unknown_dimension"
    if a.base == b.base:
        return "same"
    return "incompatible" if (a.known and b.known) else "unconfirmed"


def _same_unit(a: Unit, b: Unit) -> bool:
    return a.base == b.base and abs(a.scale - b.scale) <= 1e-12 * max(a.scale, b.scale)


def _read_unit(text: str, pos: int) -> Tuple[str, int]:
    """从 pos 处读取紧随数值的单位文本（百分号 / 已登记量纲 / 未登记拉丁单位 / 独立数量级）。"""
    for regex in (_PCT_SUFFIX_RE, _SCALED_UNIT_RE, _LATIN_UNIT_RE, _BARE_SCALE_RE):
        m = regex.match(text, pos)
        if m:
            return m.group(1), m.end()
    return "", pos


@dataclass
class Quantity:
    """一条可比较的定量表达式（主张侧或数据侧通用）。

    value 恒为幅值；符号与方向在 direction（+1 上升 / -1 下降 / 0 无方向，或原始值自身
    带符号）。change 标记这是**变化量主张**（变化动词，如「增长 / 下滑 5 亿元」）还是
    **水平主张**（「Q2 营收 5 亿元」）——两者对应不同的数据来源，幅值相同也不得互相顶替。
    """

    value: float
    kind: str                 # level | percent | pct_points
    unit: str = ""            # 规范化单位文本（percent 恒为 "%"）
    direction: int = 0        # +1 上升 / -1 下降 / 0 非变化量或无方向
    change: bool = False      # 变化量主张（对差额/变化率）而非水平主张
    origin: str = ""          # raw | total | share | diff | change | cell
    label: str = ""           # 数据来源标签（系列名 / 表格行名 / 类别）
    note: str = ""            # 派生说明，例如 "Q1→Q3"
    period: str = ""          # 期间标记
    token: str = ""           # 主张原文片段
    clause: str = ""          # 主张所在小句
    comparative: bool = False  # 同比 / 环比等比较口径
    pair_role: str = ""       # source | target（从…到…结构）
    pair_direction: int = 0   # 从…到…结构里动词表达的方向（+1/-1）；不参与端点数值匹配
    verb_direction: int = 0   # 动词方向提示（change/reach）；不改变幅值符号
    span: Tuple[int, int] = (0, 0)  # 在主张文本中的字符区间
    source: str = "claim"     # claim 或实际可见元素 title:<id>
    source_text: str = ""     # 该定量表述所属的完整可见文本
    basis: str = ""           # 统计口径（净值/含税等），来自系列或图表声明
    unit_source: str = ""     # series | chart | inferred | ""
    period_source: str = ""   # category | series | chart | ""
    basis_source: str = ""    # series | chart | ""
    series_index: Optional[int] = None
    decimals: Optional[int] = None  # 主张数字的小数位：按"显示精度"判等（5.2 亿 ≙ 5.24 亿）；数据侧为 None
    assumption: str = ""     # 派生值依赖的、数据本身无法证明的前提；命中只能记"需确认"


def _read_number(text: str, m: "re.Match[str]") -> Tuple[float, int, str, str, int]:
    """数值记号的唯一解析：返回 (值, 前置正负号, 类型 level|percent|pct_points, 规范单位, 单位结束位置)。"""
    sign, k = 0, m.start() - 1
    while k >= 0 and text[k] in "+-−":
        # 连字符紧跟数字时是日期/区间分隔符（2026-10-27、3-5 年），不是正负号：
        # 否则日期里的每个数字都会背上相反方向，污染方向性主张的判定。
        if k - 1 >= 0 and text[k - 1].isdigit():
            break
        sign = 1 if text[k] == "+" else -1
        k -= 1
    unit, unit_end = _read_unit(text, m.end())
    if unit in _PCT_UNITS:
        kind, unit = "percent", "%"
    elif unit.lower() in _PCT_POINT_UNITS:
        kind, unit = "pct_points", "pct"
    else:
        kind = "level"
    return float(m.group(0).replace(",", "")), sign, kind, unit, unit_end


def _parse_cell_quantity(text: str, default_unit: str = "") -> List[Quantity]:
    """解析表格单元格中的定量值，保留符号、百分号与单位；单元格未写单位时继承列/图表声明。"""
    out: List[Quantity] = []
    for m in _NUM_TOKEN_RE.finditer(text if isinstance(text, str) else ""):
        value, sign, kind, unit, _end = _read_number(text, m)
        unit = unit or default_unit
        if kind == "level":
            # 列/图表声明「%」或「百分点」时，单元格里的裸数字就是该口径（与系列侧同一解析），
            # 否则「-3」永远对不上主张「-3%」。
            kind, unit = _kind_from_unit(unit)
        out.append(Quantity(value=value, kind=kind, unit=unit, direction=sign, origin="cell"))
    return out


_VERB_WINDOW = 3  # 变化动词必须紧贴数字：间隔超过 3 字符视为叙述中的远处动词，不构成对它的变化量断言
_VERB_REACH_RE = re.compile(r"\s*[到至为]\s*$")  # 「增长到 / 下降为 X」是水平抵达，不是变化量


def _claim_verb_direction(prefix: str) -> Tuple[str, int]:
    """返回 (变化类型, 方向)。类型：change | reach | level。

    只认紧贴数字的动词：「增长 5 亿元」是变化量；「市场规模持续扩大，2026 年预计达到
    186 亿元」里的 186 是水平值——小句分隔已把远处动词挡在 clause 外，窗口再兜住同小句
    内的远处动词（「增长后达到 186 亿元」）。动词后紧跟 到/至/为 时按抵达（水平值）处理。
    """
    tail = prefix[-14:]
    best_pos, best_match, best = -1, None, ("level", 0)
    for regex, kind, direction in (
        (_DOWN_CHANGE_RE, "change", -1),
        (_UP_CHANGE_RE, "change", 1),
        (_DOWN_TO_RE, "reach", -1),
        (_UP_TO_RE, "reach", 1),
    ):
        found = None
        for m in regex.finditer(tail):
            found = m
        if found and found.start() > best_pos:
            best_pos, best_match, best = found.start(), found, (kind, direction)
    if best_match is None:
        return best
    if len(tail) - best_match.end() > _VERB_WINDOW:
        return ("level", 0)
    if best[0] == "change" and _VERB_REACH_RE.match(tail[best_match.end():]):
        return ("reach", best[1])
    return best


def _normalize_year(value: Optional[str]) -> Optional[int]:
    """把 25 / 2025 统一成四位年份；空值返回 None。"""
    if not value:
        return None
    number = int(value)
    return 2000 + number if number < 100 else number


def _extract_periods(text: str) -> Set[str]:
    """提取年份 / Q / H / 第X季度 / X月，统一为 Y2025、2025Q1、H2、M7。

    同一文本只有一个年份时，把该年份附着到未写年份的 Q/H/M 上：
    「2025 年 Q1」→ {Y2025, 2025Q1}，避免把 Q1 当成无年份的任意季度。
    多年度区间保留各自年份，不把两个年度错误合并。
    """
    if not isinstance(text, str):
        return set()
    years = [_normalize_year(y) for y in _YEAR_RE.findall(text)]
    unique_year = years[0] if len(set(years)) == 1 else None
    out: Set[str] = {f"Y{year}" for year in years}
    for m in _PERIOD_Q_RE.finditer(text):
        year = _normalize_year(m.group(1)) if m.group(1) else unique_year
        out.add(f"{year}Q{m.group(2)}" if year else f"Q{m.group(2)}")
    for m in _PERIOD_H_RE.finditer(text):
        year = _normalize_year(m.group(1)) if m.group(1) else unique_year
        out.add(f"{year}H{m.group(2)}" if year else f"H{m.group(2)}")
    for m in _PERIOD_CN_Q_RE.finditer(text):
        quarter = _CN_QUARTER[m.group(1)]
        out.add(f"{unique_year}Q{quarter}" if unique_year else f"Q{quarter}")
    for m in _PERIOD_MONTH_RE.finditer(text):
        out.add(f"{unique_year}M{m.group(1)}" if unique_year else f"M{m.group(1)}")
    return out


def _period_parts(token: str) -> Tuple[str, Optional[int], int]:
    """解析期间记号 → (family, year, number)。Y2025 → ("Y",2025,0)。"""
    match = re.fullmatch(r"Y((?:19|20)\d{2})", str(token).strip(), re.IGNORECASE)
    if match:
        return "Y", int(match.group(1)), 0
    match = re.fullmatch(r"(?:((?:19|20)?\d{2}))?([QHM])(\d+)", str(token).strip(), re.IGNORECASE)
    if not match:
        return "U", None, 0
    year = _normalize_year(match.group(1))
    return match.group(2).upper(), year, int(match.group(3))


def _period_stem(token: str) -> Tuple[str, Optional[str]]:
    """把 26Q3 → ("Q3","2026")；Q3 → ("Q3",None)；Y2025 → ("Y","2025")。"""
    family, year, number = _period_parts(token)
    if family == "Y":
        return "Y", (str(year) if year else None)
    if family == "U":
        return str(token), None
    return f"{family}{number}", (str(year) if year else None)


_YEAR_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})(?!\d)")


def _is_temporal_label(label: str) -> bool:
    """标签包含期间标记；仅是识别线索，不证明标签之间的时间顺序。"""
    return bool(_extract_periods(label) or _YEAR_RE.search(label))


_TEMPORAL_LABEL_RE = re.compile(
    r"(?:FY\s*)?(?:(?:(?:19|20)\d{2}\s*年?)|(?:[1-9]\d(?=Q)))?[\s/-]*"
    r"(?:Q[1-4]|H[12]|第?[一二三四]季度|(?:1[0-2]|[1-9])月)?", re.IGNORECASE)


def _best_period(tokens: Set[str]) -> str:
    """从记号集合选出最具体的一个：Q/H/M 优先于整年，带年份优先于裸期间。"""
    if not tokens:
        return ""
    return sorted(tokens, key=lambda t: (
        _period_parts(t)[0] == "Y",
        _period_parts(t)[1] is None,
        -_period_parts(t)[2],
        t,
    ))[0]


def _ordered_temporal_labels(labels: List[str]) -> bool:
    """仅当类别提供同粒度、无歧义且严格递增的期间，才用于时间方向硬判。"""
    if len(labels) < 2:
        return False
    texts = [str(label).strip() for label in labels]
    if all(re.fullmatch(r"(?:19|20)\d{2}-\d{2}(?:-\d{2})?", t) for t in texts):
        if len({len(t) for t in texts}) != 1:
            return False
        try:
            dates = [date.fromisoformat(t if len(t) == 10 else t + "-01") for t in texts]
        except ValueError:
            return False
        return all(a < b for a, b in zip(dates, dates[1:]))
    parts: List[Tuple[str, Optional[int], int]] = []
    for text in texts:
        if not _TEMPORAL_LABEL_RE.fullmatch(text):
            return False
        periods = _extract_periods(text)
        representative = _best_period(periods)
        if not representative:
            return False
        family, year, number = _period_parts(representative)
        if family == "Y":
            if year is None:
                return False
            parts.append(("Y", year, 0))
        else:
            parts.append((family, year, number))
    if len({p[0] for p in parts}) != 1 or len({p[1] is None for p in parts}) != 1:
        return False  # 混合粒度或年份有无不一致，无法比较
    keys = [((year if year is not None else 0), number) for _family, year, number in parts]
    return all(a < b for a, b in zip(keys, keys[1:]))


def _period_months(token: str) -> Optional[Set[int]]:
    """期间覆盖的月份（不含年份）；无法解析返回 None。Y 覆盖 1–12。"""
    family, _year, number = _period_parts(token)
    if family == "Y":
        return set(range(1, 13))
    if family == "M" and 1 <= number <= 12:
        return {number}
    if family == "Q" and 1 <= number <= 4:
        return set(range(3 * number - 2, 3 * number + 1))
    if family == "H" and 1 <= number <= 2:
        return set(range(6 * number - 5, 6 * number + 1))
    return None


def _period_relation(a: str, b: str) -> str:
    """返回 overlap | disjoint | unknown。

    已知年份不同必然不相交；一方是整年而另一方是未标年份的 Q/H/M 时无法判断具体年份，
    记 unknown，不能拿来证明一致，也不能判为矛盾。共享上级不算覆盖：11 月与 Q3 不相交。
    """
    fa, ya, _na = _period_parts(a)
    fb, yb, _nb = _period_parts(b)
    if fa == "U" or fb == "U":
        return "unknown"
    if ya is not None and yb is not None and ya != yb:
        return "disjoint"
    if (ya is None) != (yb is None) and "Y" in {fa, fb}:
        return "unknown"
    months_a, months_b = _period_months(a), _period_months(b)
    if months_a is None or months_b is None:
        return "unknown"
    return "overlap" if months_a & months_b else "disjoint"


def _periods_overlap(a: str, b: str) -> bool:
    """两个期间是否有确定交集（只用于判断文字谈的期间是否被图表覆盖）。"""
    return _period_relation(a, b) == "overlap"


def _periods_contradict(a: str, b: str) -> bool:
    """两条期间声明是否互相矛盾：已知年份不同，或双方都有更细粒度且全部不相交。"""
    pa, pb = _extract_periods(a), _extract_periods(b)
    if not pa or not pb:
        return False
    years_a = {_period_parts(tok)[1] for tok in pa if _period_parts(tok)[0] == "Y"}
    years_b = {_period_parts(tok)[1] for tok in pb if _period_parts(tok)[0] == "Y"}
    if years_a and years_b and not (years_a & years_b):
        return True
    fine_a = {tok for tok in pa if _period_parts(tok)[0] != "Y"}
    fine_b = {tok for tok in pb if _period_parts(tok)[0] != "Y"}
    if not fine_a or not fine_b:
        return False
    relations = [_period_relation(x, y) for x in fine_a for y in fine_b]
    return "disjoint" in relations and "overlap" not in relations and "unknown" not in relations


def _period_conflict(claim_tokens: Set[str], data_token: str) -> bool:
    """期间冲突的判据：所有主张记号都与数据点确定不相交。

    任一 overlap 说明主张覆盖该数据点；任一 unknown 说明信息不足，不能断言错期。
    """
    if not claim_tokens or not data_token:
        return False
    relations = {_period_relation(tok, data_token) for tok in claim_tokens}
    return relations == {"disjoint"}


def _meta_text(value: Any) -> str:
    return str(value or "").strip()


def _resolve_series_scope(chart: Dict[str, Any], series: Optional[Dict[str, Any]] = None) -> Dict[str, str]:
    """系列 unit/period/basis 覆盖图表级缺省；空字符串视为未声明。

    覆盖只是\"谁的声明生效\"，不是换算。两级声明互相矛盾时由 _declared_scope_conflicts
    在契约层阻断，而不是在这里悄悄让系列声明胜出。
    """
    series = series or {}
    chart_unit, series_unit = _meta_text(chart.get("unit")), _meta_text(series.get("unit"))
    chart_period, series_period = _meta_text(chart.get("period")), _meta_text(series.get("period"))
    chart_basis, series_basis = _meta_text(chart.get("basis")), _meta_text(series.get("basis"))
    return {
        "unit": series_unit or chart_unit,
        "unit_source": "series" if series_unit else ("chart" if chart_unit else ""),
        "period_decl": series_period or chart_period,
        "period_decl_source": "series" if series_period else ("chart" if chart_period else ""),
        "basis": series_basis or chart_basis,
        "basis_source": "series" if series_basis else ("chart" if chart_basis else ""),
        "series_name": _meta_text(series.get("name")),
    }


def _declared_scope_conflicts(chart: Dict[str, Any], series: Dict[str, Any]) -> List[Tuple[str, str]]:
    """系列与图表同时声明、且会让可见图题/表题与所画数据不一致的矛盾：[(字段, 说明)]。

    图题只显示图表级（或全体系列共同的）单位与期间；任何一个系列声明了不同的数量级、
    量纲或期间，观众看到的标注就与该系列数据不符——这与系列数量无关，单系列同样阻断。
    """
    out: List[Tuple[str, str]] = []
    cu, su = _meta_text(chart.get("unit")), _meta_text(series.get("unit"))
    if cu and su and not _same_unit(_parse_unit(cu), _parse_unit(su)):
        out.append(("unit", f"图表级单位「{cu}」≠ 系列「{series.get('name')}」单位「{su}」"))
    cp, sp = _meta_text(chart.get("period")), _meta_text(series.get("period"))
    if cp and sp and _periods_contradict(cp, sp):
        out.append(("period", f"图表级期间「{cp}」≠ 系列「{series.get('name')}」期间「{sp}」"))
    return out


def _effective_chart_meta(chart: Dict[str, Any]) -> Tuple[str, str]:
    """图题应显示的 (单位, 期间)：图表级声明优先；未声明时取全体系列一致的声明，否则为空。"""
    series_list = [s for s in (chart.get("data") or {}).get("series") or [] if isinstance(s, dict)]
    result = []
    for field in ("unit", "period"):
        own = _meta_text(chart.get(field))
        if own or not series_list:
            result.append(own)
            continue
        declared = [_meta_text(s.get(field)) for s in series_list]
        first = declared[0]
        if field == "unit":
            same = first and all(d and _same_unit(_parse_unit(d), _parse_unit(first)) for d in declared)
        else:
            same = first and all(d == first for d in declared)
        result.append(first if same else "")
    return result[0], result[1]


_TITLE_PAREN_RE = re.compile(r"[（(]([^（）()]+)[）)]")


def _title_unit_conflict(chart: Dict[str, Any]) -> Optional[str]:
    """图题括注里写着的单位（如\"营收（万元）\"）与数据单位矛盾时返回说明。

    只认括注中能完整解析为已登记量纲/数量级/百分号的片段，避免把普通文字当单位。
    """
    unit_text, _ = _effective_chart_meta(chart)
    if not unit_text:
        return None
    data_unit = _parse_unit(unit_text)
    for segment in _TITLE_PAREN_RE.findall(str(chart.get("title") or "")):
        for piece in re.split(r"[，,、·/\s]+", segment):
            u = _parse_unit(piece)
            if not piece or not (u.known or (u.text in _SCALE_FACTORS)):
                continue
            scale_differs = abs(u.scale - data_unit.scale) > 1e-12 * max(u.scale, data_unit.scale)
            base_differs = bool(u.base and data_unit.base and u.base != data_unit.base)
            if scale_differs or base_differs:
                return f"图题括注「{piece}」与数据单位「{unit_text}」不一致"
    return None


def _observation_period(category: Any, scope: Dict[str, str]) -> Tuple[str, str]:
    """点期间：类别上的年份/Q/H/M 优先，否则用系列/图表声明。类别名本身不是期间。

    类别只写 Q1、图表级声明 2025 时，把年份并回类别再取最具体记号，得到 2025Q1；
    这样不会把「2025 年」的主张错误核对到无年份的任意 Q1 上。
    """
    cat_text = str(category or "")
    decl = scope.get("period_decl") or ""
    tokens = _extract_periods(cat_text)
    if tokens:
        if decl and any(_period_parts(tok)[1] is None for tok in tokens):
            tokens = _extract_periods(f"{cat_text} {decl}")
        return _best_period(tokens), "category"
    if not decl:
        return "", ""
    extracted = _extract_periods(decl)
    return (_best_period(extracted) if extracted else decl), (scope.get("period_decl_source") or "chart")


def _stamp_scope(
    q: Quantity,
    scope: Dict[str, str],
    *,
    period: str = "",
    period_source: str = "",
    series_index: Optional[int] = None,
) -> Quantity:
    if not q.unit:
        q.unit = scope["unit"]
        q.unit_source = scope["unit_source"]
    elif not q.unit_source:
        q.unit_source = scope["unit_source"] or "inferred"
    q.basis = q.basis or scope["basis"]
    q.basis_source = q.basis_source or scope["basis_source"]
    if period:
        q.period = period
        q.period_source = period_source
    elif not q.period and scope.get("period_decl"):
        q.period, q.period_source = _observation_period("", scope)
    if series_index is not None:
        q.series_index = series_index
    if not q.label:
        q.label = scope.get("series_name") or ""
    return q


def _kind_from_unit(unit: str) -> Tuple[str, str]:
    """单位决定数值类型：% 是百分比，「百分点」是百分点，其余是水平值。

    系列侧与表格单元格侧共用这一解析，避免同一列数据在两条路径上变成两种类型。
    """
    base = _parse_unit(unit).base
    if base == "%":
        return "percent", "%"
    if base == "pct":
        return "pct_points", "pct"
    return "level", unit


def _parse_claim_quantities(claim: str) -> List[Quantity]:
    """把页面主张拆成独立的定量表述，并识别方向、单位、比较口径与「从 X 到 Y」配对。"""
    if not isinstance(claim, str) or not claim.strip():
        return []
    quantities: List[Quantity] = []
    for m in _NUM_TOKEN_RE.finditer(claim):
        start = m.start()
        # 每个数字只属于最近标点后的当前小句，避免前一句的同比/期间泄漏到后一句。
        clause_start = max(claim.rfind(mark, 0, start) for mark in "，,。；;、") + 1
        clause = claim[clause_start:start]
        prefix = claim[max(clause_start, start - 16):start]
        value, sign, kind, unit_norm, unit_end = _read_number(claim, m)
        verb_kind, verb_dir = _claim_verb_direction(prefix)
        change_claim = verb_kind == "change"
        direction = 0
        if kind in {"percent", "pct_points"}:
            if sign:
                direction = sign
            elif change_claim:
                direction = verb_dir
            # 「升至/降至 X%」中的 X 是水平值而非变化量：方向只作提示，不参与变化量匹配
        elif change_claim or verb_kind == "reach" or sign:
            # 水平量也能带方向：显式正负号是值的一部分；抵达动词只提示方向。
            # 变化动词（增长/下滑 X 亿元）必须带上方向，否则「下滑 5 亿元」会凭幅值
            # 命中图上 +5 亿元的差额——绝对值相同而符号相反的误判。
            direction = verb_dir or sign

        comparative = bool(_COMPARATIVE_RE.search(clause))
        number = m.group(0)
        quantities.append(Quantity(
            value=value, kind=kind, unit=unit_norm, direction=direction, change=change_claim,
            origin="claim", token=claim[m.start():unit_end], clause=clause,
            comparative=comparative, span=(m.start(), unit_end),
            decimals=len(number.partition(".")[2]),
            verb_direction=verb_dir if verb_kind in {"change", "reach"} else 0,
        ))

    # 「从 X 到/降至 Y」配对：登记源值/目标值，供序列顺序与方向一致性核对。
    # 配对双方都是水平值（起点与终点的状态），不因中途动词被当作变化量。
    def _explicit_sign(q: Quantity) -> bool:
        return bool(re.search(r"[-−+]\s*$", claim[max(0, q.span[0] - 2):q.span[0]]))

    for i in range(len(quantities) - 1):
        a, b = quantities[i], quantities[i + 1]
        between = claim[a.span[1]:b.span[0]]
        lead = claim[max(0, a.span[0] - 3):a.span[0]]
        if "从" in lead and re.search(r"(到|至|为)\s*$", between):
            a.pair_role, b.pair_role = "source", "target"
            a.change = b.change = False
            # 「从 X 降至 Y」里的 Y 是**终点水平值**，不是 -Y 的变化量。
            # 动词方向只用于后面的序列方向核对，绝不能作为幅值符号参与数值匹配，
            # 否则所有合法的下降主张都会因 -Y 找不到对应原始值而被误报为 E016。
            narrative = b.verb_direction or a.verb_direction or (1 if b.value >= a.value else -1)
            b.pair_direction = narrative
            if not _explicit_sign(a):
                a.direction = 0
            if not _explicit_sign(b):
                b.direction = 0
    return quantities


def _collect_data_quantities(chart: Dict[str, Any]) -> Tuple[List[Quantity], List[Tuple[str, List[str], List[float], str, bool]]]:
    """把一个图表/表格元素展开为数据侧定量表达式，并返回可做顺序比对的序列轨迹。

    每个系列独立解析 unit/period/basis；图表级字段只在系列未声明时作为缺省。
    派生值继承其来源系列的身份与口径，跨口径合计不会被当成可互换来源。
    """
    out: List[Quantity] = []
    traces: List[Tuple[str, List[str], List[float], str, bool]] = []
    kind, data = chart["kind"], chart["data"]
    chart_scope = _resolve_series_scope(chart)

    if kind == "table" and "headers" in data:
        rows = data["rows"]
        headers = data["headers"]
        for ri, row in enumerate(rows):
            row_label = str(row[0]) if row and not _extract_numbers(str(row[0])) else ""
            for ci, cell in enumerate(row):
                qs = _parse_cell_quantity(str(cell), chart_scope["unit"])
                col_periods = _extract_periods(str(headers[ci])) if ci < len(headers) else set()
                if col_periods:
                    period, period_source = sorted(col_periods)[0], "category"
                else:
                    period, period_source = _observation_period("", chart_scope)
                for q in qs:
                    q.label = row_label
                    q.note = f"表格第{ri + 1}行" + (f"·{row_label}" if row_label else "")
                    out.append(_stamp_scope(q, chart_scope, period=period, period_source=period_source))
        return out, traces
    if kind == "table":
        categories = [str(c) for c in data["categories"]]
        for si, series in enumerate(data["series"]):
            scope = _resolve_series_scope(chart, series)
            name = str(series["name"])
            kind_u, unit = _kind_from_unit(scope["unit"])
            for ci, value in enumerate(series["values"]):
                category = categories[ci]
                period, period_source = _observation_period(category, scope)
                q = Quantity(
                    value=float(value), kind=kind_u, unit=unit, origin="cell", label=name,
                    note=f"表格 {name}@{category or ci + 1}",
                )
                out.append(_stamp_scope(q, scope, period=period, period_source=period_source, series_index=si))
        return out, traces

    parsed = [(str(s["name"]), [float(v) for v in s["values"]], _resolve_series_scope(chart, s), si)
              for si, s in enumerate(data["series"])]
    cat_labels = [str(c) for c in data["categories"]]
    # 图型本身不证明时间顺序；类别须给出可比较且递增的日期或期间。
    # 地区、倒序日期、混合粒度及不明标签的差值只能作为待确认假设。
    ordered = _ordered_temporal_labels(cat_labels)
    order_assumption = "" if ordered else "把类别当作时间先后（标签未证明同粒度递增期间）"
    additive_assumption = "假定该指标可跨类别相加（流量而非存量、比率或均值）"
    parts_assumption = "假定各系列是同一整体的互斥组成（可相加、份额有共同分母）"

    def _period_at(i: int, scope: Dict[str, str]) -> Tuple[str, str]:
        category = cat_labels[i] if i < len(cat_labels) else ""
        return _observation_period(category, scope)

    def _change_pair(name, scope, si, i, j, label):
        p0, _s0 = _period_at(i, scope)
        p1, s1 = _period_at(j, scope)
        kind_u, unit = _kind_from_unit(scope["unit"])
        # 两个百分点值之差仍是百分点；百分比之差是百分点（10%→40% 是 +30 个百分点）。
        diff_kind, diff_unit = (("pct_points", "pct") if kind_u in {"percent", "pct_points"}
                                else ("level", unit))
        diff = vals_of[si][j] - vals_of[si][i]
        direction = 0 if diff == 0 else (1 if diff > 0 else -1)
        a, b = cat_labels[i] if i < len(cat_labels) else i + 1, cat_labels[j] if j < len(cat_labels) else j + 1
        q = Quantity(value=abs(diff), kind=diff_kind, unit=diff_unit, direction=direction,
                     origin="diff", label=name, note=f"{p0 or a}→{p1 or b} {label}差额", assumption=order_assumption)
        out.append(_stamp_scope(q, scope, period=p1, period_source=s1, series_index=si))
        if vals_of[si][i] != 0 and kind_u == "level":
            # 变化率只对水平值有语义：百分点值的「相对变化」不是观众会读的指标。
            pct = diff / abs(vals_of[si][i]) * 100.0
            q = Quantity(value=abs(pct), kind="percent", unit="%", direction=direction,
                         origin="change", label=name, note=f"{p0 or a}→{p1 or b} {label}变化率",
                         assumption=order_assumption)
            out.append(_stamp_scope(q, scope, period=p1, period_source=s1, series_index=si))

    vals_of: Dict[int, List[float]] = {}
    for name, vals, scope, si in parsed:
        vals_of[si] = vals
        traces.append((name, cat_labels, vals, scope["unit"], ordered))
        kind_u, unit = _kind_from_unit(scope["unit"])
        for i, v in enumerate(vals):
            period, period_source = _period_at(i, scope)
            q = Quantity(value=v, kind=kind_u, unit=unit, direction=0, origin="raw",
                         label=name, note=f"{name}@{period or (cat_labels[i] if i < len(cat_labels) else i + 1)}")
            out.append(_stamp_scope(q, scope, period=period, period_source=period_source, series_index=si))
        if kind in {"pie", "donut"}:
            continue
        for i in range(len(vals) - 1):
            _change_pair(name, scope, si, i, i + 1, "")
        if len(vals) >= 3:
            _change_pair(name, scope, si, 0, len(vals) - 1, "累计")

    for name, vals, scope, si in parsed:
        total = sum(abs(v) for v in vals)
        if kind in {"pie", "donut"} and total:
            # 饼/环由图型声明\"部分—整体\"，PowerPoint 显示的正是这些份额：无需额外假设。
            for i, v in enumerate(vals):
                period, period_source = _period_at(i, scope)
                q = Quantity(value=abs(v) / total * 100.0, kind="percent", unit="%", origin="share",
                             label=cat_labels[i] if i < len(cat_labels) else name, note=f"{name} 切片占比")
                out.append(_stamp_scope(q, scope, period=period, period_source=period_source, series_index=si))

    if kind in {"bar", "line"} and len(parsed) >= 2:
        units = {(_parse_unit(scope["unit"]).scale, _parse_unit(scope["unit"]).base, scope["basis"]) for _n, _v, scope, _i in parsed}
        if len(units) == 1 and _kind_from_unit(parsed[0][2]["unit"])[0] == "level":
            axis_scope = parsed[0][2]
            for k in range(min(len(v) for _, v, _s, _i in parsed)):
                col_total = sum(v[k] for _, v, _s, _i in parsed)
                period, period_source = _period_at(k, axis_scope)
                q = Quantity(value=col_total, kind="level", unit=axis_scope["unit"], origin="total",
                             label="合计", note=f"{period or k + 1} 多系列合计", assumption=parts_assumption)
                out.append(_stamp_scope(q, axis_scope, period=period, period_source=period_source))
                if col_total:
                    for name, vals, scope, si in parsed:
                        q = Quantity(value=abs(vals[k] / col_total) * 100.0, kind="percent", unit="%", origin="share",
                                     label=name, note=f"{period or k + 1}·{name} 占比", assumption=parts_assumption)
                        out.append(_stamp_scope(q, scope, period=period, period_source=period_source, series_index=si))
    for name, vals, scope, si in parsed:
        kind_u, unit = _kind_from_unit(scope["unit"])
        if kind_u == "percent" or kind in {"pie", "donut"}:
            continue  # 百分比/份额相加没有语义；饼图合计就是 100% 的分母本身
        q = Quantity(value=sum(vals), kind=kind_u, unit=unit, origin="total",
                     label=name, note=f"{name} 系列合计", assumption=additive_assumption)
        out.append(_stamp_scope(q, scope, series_index=si))
    return out, traces


def _signed_value(q: Quantity) -> float:
    """定量表述的带符号值：direction 记录符号（原始值自身带符号时 direction 为 0）。"""
    return q.value * q.direction if q.direction else q.value


def _values_equal(q: Quantity, d: Quantity) -> Tuple[bool, bool]:
    """返回 (按主张显示精度是否相等, 是否经过数量级换算)。

    比较的是**带符号**的值：「5 亿元」不对上图上的 -5 亿元，「-5 亿元」也不对上 +5 亿元——
    绝对值相同而符号相反是两个不同的数，这正是方向性误判的来源。已知数量级总是参与换算，
    绝不回退到裸数字比较。主张写到几位小数就按几位比较（±半个末位）：「5.2 亿」与 5.24 亿
    一致，与 5.26 亿不一致。
    """
    qu, du = _parse_unit(q.unit), _parse_unit(d.unit)
    a, b = _signed_value(q) * qu.scale, _signed_value(d) * du.scale
    scale_used = abs(qu.scale - du.scale) > 1e-12 * max(qu.scale, du.scale)
    half_ulp = 0.5 * 10 ** (-q.decimals) if q.decimals is not None else 0.0
    tolerance = half_ulp * qu.scale + 1e-9 * max(1.0, abs(b))
    return abs(a - b) <= tolerance, scale_used


@dataclass
class QuantityVerdict:
    quantity: Quantity
    status: str          # verified | review | conflict | skipped
    mode: str            # ok | unit | direction | period | basis | no_candidate | ambiguous | assumption | missing
    detail: str
    matched: Optional[Quantity] = None   # 命中的数据侧来源（用于回查期间/单位/口径）


def _claim_periods_of(q: Quantity) -> Set[str]:
    return _extract_periods(q.clause) | _extract_periods(q.token)


def _claim_vs_data_period_state(q: Quantity, d: Quantity) -> str:
    """主张期间相对数据点：na | ok | missing | conflict。

    变化量允许命中起点或终点；任一端有确定交集即可确认。无法判断年份/粒度时记 missing，
    不能被当成已验证，也不能被当成矛盾。
    """
    claim_periods = _claim_periods_of(q)
    if not claim_periods:
        return "na"
    data_tokens = {d.period} if d.period else set()
    if d.origin in {"diff", "change"}:
        data_tokens |= _extract_periods(d.note)
    data_tokens.discard("")
    if not data_tokens:
        return "missing"
    relations = {_period_relation(claim, data) for claim in claim_periods for data in data_tokens}
    if "overlap" in relations:
        return "ok"
    if relations == {"disjoint"}:
        return "conflict"
    return "missing"


def _claim_vs_data_period_conflict(q: Quantity, d: Quantity) -> bool:
    """主张期间与数据点是否确定冲突（仅 conflict；信息不足由 missing 承担）。"""
    return _claim_vs_data_period_state(q, d) == "conflict"


def _describe_source(d: Quantity) -> str:
    bits = [d.label or "图表数据", d.origin]
    if d.unit:
        bits.append(d.unit)
    if d.period:
        bits.append(d.period)
    if d.basis:
        bits.append(d.basis)
    return "·".join(bits)


def _source_name_in_text(text: str, name: str) -> bool:
    """仅把明确出现的名称当作弱身份线索；同名仍不证明同指标或同口径。"""
    text, name = str(text or ""), str(name or "").strip()
    if not text or not name:
        return False
    if re.fullmatch(r"[A-Za-z0-9_]+", name):
        return bool(re.search(r"(?<![A-Za-z0-9_])" + re.escape(name) + r"(?![A-Za-z0-9_])", text, re.IGNORECASE))
    return name in text


def _mentions_source(q: Quantity, d: Quantity) -> bool:
    if _source_name_in_text(q.source_text, d.label):
        return True
    # 单系列分类图：类别名可作为线索；Q1/年份只是期间，不能冒充指标身份。
    if d.origin == "raw" and "@" in d.note:
        category = d.note.rsplit("@", 1)[-1]
        return not _is_temporal_label(category) and _source_name_in_text(q.source_text, category)
    return False


def _unitless_levels(data: List[Quantity]) -> List[Quantity]:
    """图表没标单位、按比例存储的水平值。

    百分比／百分点主张在找不到同量纲候选时，要用它做最后一级降级比对（相差百倍尺度
    的可能），因此两个阶段都要读它；独立成函数避免口径漂移。
    """
    return [d for d in data if d.kind == "level" and not d.unit]


def _semantic_pool(q: Quantity, data: List[Quantity]) -> List[Quantity]:
    """按语义池筛候选：变化量主张只认差额/变化率，水平主张只认原始值/单元格。

    幅值相同而语义不同不得互相顶替——否则「下滑 5 亿元」会命中图上 +5 亿元的差额。
    合计与份额属于需要假设的派生，留给命中后的 assumption 路径解释。
    """
    kind_gate = {"percent": {"percent", "pct_points"}, "pct_points": {"pct_points"}, "level": {"level"}}[q.kind]
    origins = {"diff", "change"} if q.change else {"raw", "cell", "total", "share"}
    pool = [d for d in data if d.kind in kind_gate and d.origin in origins]
    if q.kind in {"percent", "pct_points"}:
        pool += _unitless_levels(data)  # 未标单位时，按比例存储的值仍可比对（记需确认）
    elif q.direction and not q.change:
        # 显式带符号的水平量（「净利 -5 亿元」）既可能命中带符号的水平值，也可能与同号
        # 差额一致；带符号比较已排除反号值，两种读法都不构成矛盾。
        pool += [d for d in data if d.kind == "level" and d.origin in {"diff", "change"}]
    return pool


def _partition_by_unit(q: Quantity, pool: List[Quantity]) -> Tuple[
        List[Quantity], List[Tuple[Quantity, str]], List[Quantity], List[Quantity]]:
    """把候选按量纲分成 可比／维度未知／换算未确认／不可比 四类。

    只有 same 进可比池；missing 与 unknown_dimension 必须让结论降级为「无法核对」，
    不能因为数字相同就当一致。
    """
    q_unit = _parse_unit(q.unit)
    comparable: List[Quantity] = []
    unknown: List[Tuple[Quantity, str]] = []
    unconfirmed: List[Quantity] = []
    incompatible: List[Quantity] = []
    for d in pool:
        # 百分比对百分点留在可比池里，由命中后的 assumption 路径解释，不当成单位冲突。
        if q.kind == "percent" and d.kind == "pct_points":
            comparable.append(d)
            continue
        relation = _unit_relation(q_unit, _parse_unit(d.unit))
        if relation == "same":
            comparable.append(d)
        elif relation in {"missing", "unknown_dimension"}:
            unknown.append((d, relation))
        elif relation == "unconfirmed":
            unconfirmed.append(d)
        else:
            incompatible.append(d)
    return comparable, unknown, unconfirmed, incompatible


def _match_quantity(q: Quantity, data: List[Quantity]) -> QuantityVerdict:
    """对单条主张数值做语义、量纲、数量级、方向、期间、口径感知的匹配。

    判定次序即证据强度：可确认一致 → 只与带假设的派生值一致 → 与本页数据直接矛盾 →
    无法核对。同值但量纲未知、换算未确认或口径不明，一律不标已验证。

    先分语义池：**变化量主张**（「增长 5 亿元」「提升 3 个百分点」）只能由差额 / 变化率
    （diff/change）证明；**水平主张**（「Q2 营收 5 亿元」「占比 30%」）只能由原始值 /
    单元格证明，合计与份额是需要假设的派生。幅值相同而语义不同，不得互相顶替——否则
    「下滑 5 亿元」会命中图上 +5 亿元的差额而被判为一致。
    """
    if q.kind == "level" and not q.unit:
        return QuantityVerdict(q, "skipped", "missing", "无单位裸数，未参与核对")

    pool = _semantic_pool(q, data)
    comparable, unknown, unconfirmed, incompatible = _partition_by_unit(q, pool)
    q_unit = _parse_unit(q.unit)

    dir_pool = comparable
    dir_blocked: List[Quantity] = []
    # 只有变化量主张才按方向过滤候选（「增长 3%」不得命中 -3 的差额）。显式带符号的
    # 水平量（「利润率 -3%」「净利 -120 万元」）的符号是值的一部分，由带符号比较排除
    # 反号值即可——方向过滤会连无方向的原始值一起挡掉。
    directional = q.change and q.direction != 0
    if directional:
        dir_pool = [d for d in comparable if d.direction == q.direction]
        dir_blocked = [d for d in comparable if d.direction != 0 and d.direction != q.direction]

    claim_periods = _claim_periods_of(q)

    def _period_state(d: Quantity) -> str:
        if not claim_periods:
            return "na"
        return _claim_vs_data_period_state(q, d)

    hits: List[Tuple[Quantity, bool]] = []
    period_conflicts: List[Quantity] = []
    missing_period: List[Quantity] = []
    for d in dir_pool:
        equal, scaled = _values_equal(q, d)
        if not equal:
            continue
        state = _period_state(d)
        if state == "conflict":
            period_conflicts.append(d)
        elif state == "missing":
            missing_period.append(d)
        else:
            hits.append((d, scaled))

    clean = [h for h in hits if not h[0].assumption]
    if clean:
        clean.sort(key=lambda x: (not _mentions_source(q, x[0]), x[1], x[0].origin not in {"raw", "cell"}))
        best, scaled = clean[0]
        periods = {h[0].period for h in clean if h[0].period}
        if not claim_periods and len(periods) > 1:
            return QuantityVerdict(q, "review", "ambiguous",
                                   f"同值出现在不同期间（{'/'.join(sorted(periods))}），主张未标明期间，不能把同值当作同一观察点")
        bases = {h[0].basis for h in clean if h[0].basis}
        if len(bases) > 1:
            return QuantityVerdict(q, "review", "ambiguous",
                                   f"同值出现在不同统计口径（{'/'.join(sorted(bases))}），不能把同值当作同一指标")
        labels = {h[0].label for h in clean if h[0].label}
        if len(labels) > 1 and q.clause and not any(lb in q.clause for lb in labels):
            if len({h[0].label for h in clean if h[0].value == best.value}) > 1:
                return QuantityVerdict(q, "review", "ambiguous",
                                       f"多个来源同值（{'/'.join(sorted(labels))}），无法确定对应对象")
        if not _mentions_source(q, best):
            return QuantityVerdict(q, "review", "identity",
                                   f"数值与 {_describe_source(best)} 一致，但主张未明确点名该系列或非时间类别；"
                                   "同值不足以确认指标身份，请核对来源", matched=best)
        if q.comparative and best.origin in {"diff", "change"}:
            # 同比/环比是口径承诺：图上的相邻期间差额不能自动充当同一比较对象。
            return QuantityVerdict(q, "review", "comparative",
                                   f"与{best.note or best.origin}数值一致，但主张使用同比/环比口径，"
                                   f"图表未证明同一比较对象（如去年同期）", matched=best)
        if best.kind == "pct_points" and q.kind == "percent":
            return QuantityVerdict(q, "review", "assumption",
                                   f"与百分点口径「{best.note or best.label}」数值一致，需确认百分比/百分点表述",
                                   matched=best)
        extra = []
        if scaled:
            extra.append("经数量级换算")
        exact = abs(_signed_value(q) * q_unit.scale - _signed_value(best) * _parse_unit(best.unit).scale) <= 1e-9 * max(1.0, abs(_signed_value(best)))
        if not exact:
            extra.append(f"按主张精度（{q.decimals or 0} 位小数）四舍五入一致")
        if best.origin in {"total", "share", "diff", "change"}:
            extra.append(f"派生自{best.note or best.origin}")
        if best.unit_source:
            extra.append(f"单位来自{best.unit_source}")
        if best.period:
            extra.append(f"期间{best.period}")
        if best.basis:
            extra.append(f"口径{best.basis}")
        detail = f"对应{_describe_source(best)}" + ("（" + "，".join(extra) + "）" if extra else "")
        return QuantityVerdict(q, "verified", "ok", detail, matched=best)
    if hits:
        best = hits[0][0]
        return QuantityVerdict(q, "review", "assumption",
                               f"只与派生值「{best.note}」一致；该派生{best.assumption}，数据本身不能证明，需确认后才能成立",
                               matched=best)

    if period_conflicts:
        d = next((item for item in period_conflicts if _mentions_source(q, item)), period_conflicts[0])
        if not _mentions_source(q, d) or d.assumption:
            return QuantityVerdict(q, "review", "identity",
                                   f"同值位于 {d.period}，但指标身份或派生前提未确认；不能断言错期", matched=d)
        return QuantityVerdict(
            q, "conflict", "period",
            f"主张期间 {'/'.join(sorted(claim_periods))} 与数据期间 {d.period}"
            f"（{d.label or d.note}）不一致，同值不能当作同一观察点",
            matched=d,
        )
    if missing_period:
        d = missing_period[0]
        return QuantityVerdict(q, "review", "missing",
                               f"数值与 {_describe_source(d)} 相同，但数据侧缺少可核对的期间，不能把同值当作已验证",
                               matched=d)
    if dir_blocked:
        # 同单位并不代表同一指标；同一系列若既上升又下降，也无法仅凭某一差额断言反向。
        named_opposite = [d for d in dir_blocked if _mentions_source(q, d) and not d.assumption
                          and not _claim_vs_data_period_conflict(q, d)]
        named_same = [d for d in dir_pool if _mentions_source(q, d) and not _claim_vs_data_period_conflict(q, d)]
        names = {d.label for d in named_opposite + named_same if d.label}
        sample = "、".join(f"{'+' if d.direction > 0 else '-'}{d.value:g}" for d in dir_blocked[:3])
        if named_opposite and not named_same and len(names) == 1:
            named_sample = "、".join(f"{'+' if d.direction > 0 else '-'}{d.value:g}" for d in named_opposite[:3])
            return QuantityVerdict(q, "conflict", "direction",
                                   f"主张点名的系列只有相反方向变化（{named_sample}），请核对指标与期间")
        return QuantityVerdict(q, "review", "direction",
                               f"存在反向变化（{sample}），但指标身份或比较区间不唯一，不能断言主张矛盾")

    # 同量纲、数字相同而数量级不同：典型的\"万/亿\"错位，是与本页数据的直接矛盾
    for d in comparable:
        du = _parse_unit(d.unit)
        if abs(q_unit.scale - du.scale) > 1e-12 * max(q_unit.scale, du.scale) and \
                _values_equal(q, Quantity(value=d.value, kind=d.kind, unit=q.unit, direction=d.direction))[0]:
            if not _mentions_source(q, d):
                return QuantityVerdict(q, "review", "identity",
                                       f"数字相同但单位数量级不同；尚未建立主张与 {d.label or d.note} 的指标身份，不能断言矛盾",
                                       matched=d)
            return QuantityVerdict(q, "conflict", "unit",
                                   f"数字同为 {d.value:g}，但主张单位为「{q.unit}」、图表单位为「{d.unit}」，数量级相差 "
                                   f"{max(q_unit.scale, du.scale) / min(q_unit.scale, du.scale):g} 倍", matched=d)

    if unconfirmed and not comparable:
        units = "/".join(sorted({d.unit for d in unconfirmed}))
        return QuantityVerdict(q, "review", "unconfirmed",
                               f"主张单位「{q.unit}」与图表单位（{units}）之间没有已确认的换算关系，无法核对；"
                               f"请改用同一单位，或在数据中给出换算后的值")
    for d, relation in unknown:
        du = _parse_unit(d.unit)
        if d.unit and abs(q_unit.scale - du.scale) > 1e-12 * max(q_unit.scale, du.scale) and \
                _values_equal(q, Quantity(value=d.value, kind=d.kind, unit=q.unit, direction=d.direction))[0]:
            return QuantityVerdict(q, "review", "unit",
                                   f"数字同为 {d.value:g}，但主张写「{q.unit}」、图表为「{d.unit}」，数量级相差 "
                                   f"{max(q_unit.scale, du.scale) / min(q_unit.scale, du.scale):g} 倍；且一侧量纲未知，无法判定一致",
                                   matched=d)
        if _values_equal(q, d)[0]:
            if relation == "missing":
                side = "数据" if not d.unit else "主张"
            else:
                side = "主张" if not q_unit.base else "数据"
            reason = "单位缺失" if relation == "missing" else "只写了数量级、量纲未知"
            return QuantityVerdict(q, "review", "missing",
                                   f"按已知数量级与 {d.label or '图表数据'} 的 {d.value:g}{d.unit} 一致，但{side}侧{reason}，"
                                   f"不能把同值当作已验证", matched=d)
    for d in incompatible:
        if _values_equal(q, d)[0]:
            return QuantityVerdict(q, "review", "unit",
                                   f"数值与 {d.label or '图表数据'} 的 {d.value:g}{d.unit} 相同但量纲不同（主张「{q.unit}」）："
                                   f"若指同一指标则单位有误；若是另一指标，请注明来源", matched=d)
    if incompatible and not comparable and not unknown and not unconfirmed:
        return QuantityVerdict(q, "skipped", "unit",
                               f"与本图单位（{'/'.join(sorted({d.unit for d in incompatible}))}）不同量纲，不在本图核对范围")

    if directional and not any(d.direction != 0 for d in comparable):
        return QuantityVerdict(q, "review", "no_candidate", "图表无同量纲的变化量，无法核对方向性主张")
    if q.comparative:
        return QuantityVerdict(q, "review", "no_candidate",
                               "主张使用同比/环比等比较口径，图表未提供同口径数据，无法核对")
    if q.kind in {"percent", "pct_points"}:
        for d in _unitless_levels(data):
            if abs(d.value * 100.0 - q.value) <= 1e-6 or abs(d.value - q.value / 100.0) <= 1e-9:
                return QuantityVerdict(q, "review", "assumption",
                                       f"图表未标注单位，{d.value:g} 与主张 {q.value:g}% 可能相差百倍尺度，需确认是否按比例存储",
                                       matched=d)
    return QuantityVerdict(q, "review", "no_candidate",
                           "本页图表/表格中找不到与之一致的原始值或派生值；若数字来自其他来源，请写入 page.evidence 并具名复核")


def verify_page_claim(page: Dict[str, Any]) -> List[QuantityVerdict]:
    """核对 page.claim 与**全部可见文本元素**；任一文本正确都不能替另一文本免责。

    正文、图注、页脚里的定量表述与标题同等对待：只要写在本页且能被本页数据核对，
    就不能让错误数字藏在非标题元素里。无单位年份等裸数仍由 _match_quantity 跳过。
    """
    all_elements = [el for el in (page.get("elements") or []) if isinstance(el, dict)]
    elements = [el for el in all_elements if el.get("type") == "chart"]
    claim = page.get("claim") or ""
    claim_texts: List[Tuple[str, str]] = [("claim", claim)]
    for el in all_elements:
        if el.get("type") not in TEXT_ELEMENT_TYPES:
            continue
        text = str(el.get("content") or "")
        if text.strip() and text.strip() != claim.strip():
            claim_texts.append((f"{el.get('type')}:{el.get('id')}", text))
    quantities: List[Quantity] = []
    for source, text in claim_texts:
        for q in _parse_claim_quantities(text):
            q.source, q.source_text = source, text
            quantities.append(q)
    if not elements:
        return []

    data: List[Quantity] = []
    traces: List[Tuple[str, List[str], List[float], str, bool]] = []
    data_periods: Set[str] = set()
    for el in elements:
        chart = el.get("chart")
        if not isinstance(chart, dict):
            continue
        dq, tr = _collect_data_quantities(chart)
        data.extend(dq)
        traces.extend(tr)
        for text in (chart.get("period"), chart.get("title"), chart.get("basis")):
            data_periods |= _extract_periods(str(text or ""))
        payload = chart.get("data") if isinstance(chart.get("data"), dict) else {}
        for c in (payload.get("categories") or []):
            data_periods |= _extract_periods(str(c))
        for series in (payload.get("series") or []):
            if isinstance(series, dict):
                data_periods |= _extract_periods(str(series.get("period") or ""))
        for item in dq:
            data_periods |= _extract_periods(item.period)
    verdicts: List[QuantityVerdict] = []
    for q in quantities:
        v = _match_quantity(q, data)
        local_periods = _claim_periods_of(q)
        if v.status == "verified" and v.matched is not None and local_periods:
            period_state = _claim_vs_data_period_state(q, v.matched)
            if period_state == "conflict":
                v = QuantityVerdict(
                    q, "conflict", "period",
                    f"主张期间 {'/'.join(sorted(local_periods))} 与数据期间 {v.matched.period}"
                    f"（{v.matched.label or v.matched.note}）不一致，同值可能来自其他期间",
                    matched=v.matched,
                )
            elif period_state == "missing":
                v = QuantityVerdict(
                    q, "review", "missing",
                    f"主张含期间 {'/'.join(sorted(local_periods))}，命中数据期间不足或无法与主张对齐，不能把同值当作同期间",
                    matched=v.matched,
                )
        # 期间交叉核对：图表有期间信息且主张期间不在其间 → 冲突；主张有期间而图表完全没有 → 不臆断为通过
        if v.status in {"conflict", "review"} and v.mode == "no_candidate":
            claim_periods = _extract_periods(q.source_text)
            missing = {t for t in claim_periods if not any(_periods_overlap(t, c) for c in data_periods)} if data_periods else set()
            if missing:
                v = QuantityVerdict(q, "review", "period",
                                    f"主张期间 {'/'.join(sorted(claim_periods))} 超出图表期间 {'/'.join(sorted(data_periods))}，"
                                    f"本页数据无法核对")
        verdicts.append(v)

    # 图注/来源/正文中的期间与图表期间完全无交集 → 口径错配（只在与全部图表期间都不相交时提示，
    # 允许"对比去年同期"这类部分重叠的合法写法）
    if data_periods:
        for el in (page.get("elements") or []):
            if not isinstance(el, dict) or el.get("type") not in {"caption", "footer", "body", "text"}:
                continue
            text_periods = _extract_periods(str(el.get("content") or ""))
            if not text_periods:
                continue
            if not any(_periods_overlap(t, c) for t in text_periods for c in data_periods):
                verdicts.append(QuantityVerdict(
                    Quantity(value=0.0, kind="level", origin="claim", source=f"{el.get('type')}:{el.get('id')}",
                             token="/".join(sorted(text_periods)), clause=str(el.get("id") or "")),
                    "review", "period",
                    f"元素 {el.get('id')} 标注期间 {'/'.join(sorted(text_periods))} 与图表期间 "
                    f"{'/'.join(sorted(data_periods))} 无交集，请核对图注/来源口径",
                ))

    # 「从 X 到 Y」配对：序列中 Y 必须晚于 X，且方向与叙述一致
    sources = [v for v in verdicts if v.quantity.pair_role == "source"]
    targets = [v for v in verdicts if v.quantity.pair_role == "target"]
    for s in sources:
        for t in targets:
            if s.quantity.source != t.quantity.source:
                continue
            for name, cats, vals, unit, ordered in traces:
                # 不能让另一指标的同值、或不同数量级的裸数字，触发不可豁免的顺序矛盾。
                if not _source_name_in_text(s.quantity.source_text, name):
                    continue
                if any(_unit_relation(_parse_unit(q.unit), _parse_unit(unit)) != "same"
                       for q in (s.quantity, t.quantity)):
                    continue
                try:
                    i = next(k for k, x in enumerate(vals)
                             if _values_equal(s.quantity, Quantity(value=x, kind=s.quantity.kind, unit=unit))[0])
                    j = next(k for k, x in enumerate(vals)
                             if _values_equal(t.quantity, Quantity(value=x, kind=t.quantity.kind, unit=unit))[0])
                except StopIteration:
                    continue
                label_i = cats[i] if i < len(cats) else str(i + 1)
                label_j = cats[j] if j < len(cats) else str(j + 1)
                if j <= i:
                    if ordered:
                        # 类别自带期间：主张把较晚的值写成起点、较早的值写成终点，
                        # 方向与文字可以自洽，但与数据的时间顺序直接矛盾（不是"无法核对"）。
                        t.status, t.mode = "conflict", "direction"
                        t.detail = (f"「从…到…」与 {name} 序列的时间顺序（{label_i}→{label_j}）相反："
                                    "主张把较晚的值写成起点、较早的值写成终点")
                    else:
                        t.status, t.mode = "review", "direction"
                        t.detail = (f"「从…到…」的先后顺序与 {name} 序列（{label_i}→{label_j}）不一致；"
                                    "类别未标期间，顺序无法由图证明，请复核比较对象")
                sign = 0 if vals[j] == vals[i] else (1 if vals[j] > vals[i] else -1)
                if t.quantity.pair_direction and sign and t.quantity.pair_direction != sign:
                    if ordered:
                        t.status, t.mode = "conflict", "direction"
                        t.detail = f"叙述方向与 {name} 序列实际方向（{'上升' if sign > 0 else '下降'}）相反"
                    else:
                        t.status, t.mode = "review", "direction"
                        t.detail = f"{name} 的类别不构成已知先后，不能由排列顺序断定变化方向"
                break
    return verdicts


def _validate_claim_evidence(page: Dict[str, Any], errors: List[CompileError]) -> None:
    """逐项输出数值筛查结论：与本页数据矛盾记 E008（不可复核），无法核对记 E016（可具名复核）。

    每项独立定位到可见文本中的原文片段，具名复核只能覆盖它点名的那一项。
    """
    pid = page.get("id", "?")
    try:
        verdicts = verify_page_claim(page)
    except Exception as exc:  # 筛查程序自身缺陷：如实阻断，而不是当作"没有问题"
        errors.append(CompileError(
            "E008", ErrorLevel.HARD, f"page[{pid}].claim", "主张—数据筛查可执行",
            f"筛查程序异常（实现缺陷）：{type(exc).__name__}: {exc}",
            "数值筛查未能完成，不得推定一致；修复实现或人工核对后再交付",
        ))
        return
    kind_txt = {"percent": "百分比", "pct_points": "百分点", "level": "数值"}
    for v in verdicts:
        if v.status not in {"conflict", "review"}:
            continue
        q = v.quantity
        source = "claim" if q.source == "claim" else q.source.replace(":", "(") + ")"
        token = q.token or f"{q.value:g}"
        location = f"page[{pid}].{source}「{token}」"
        actual = f"[{v.mode}] {token}" + ("" if v.mode == "period" and not q.unit and not q.value else f"（{kind_txt[q.kind]}）") + f"：{v.detail}"
        if v.status == "conflict":
            errors.append(CompileError(
                "E008", ErrorLevel.HARD, location,
                "可见定量表述不得与本页图表/表格数据矛盾（方向、期间、同量纲数量级）", actual,
                "页面文字与它自己的数据互相矛盾：改正文字或数据；E008 不可具名复核，也不被风险批准覆盖",
            ))
        else:
            errors.append(CompileError(
                "E016", ErrorLevel.SOFT, location,
                "可见定量表述应能由本页数据核对（量纲、数量级、期间与派生假设均确认）", actual,
                "筛查无法确认一致≠已证伪：补齐单位/期间、改用可核对的表述，或在 evidence 写明来源后具名复核",
            ))


# ---------------------------------------------------------------------------
# 文本装载能力估算（E010）：只是估算，真实结论以 G-PRE 渲染测量为准
# ---------------------------------------------------------------------------

TEXT_MARGIN_X_IN = POLICY["typography"]["margin_in"]["x"] * 2
TEXT_MARGIN_Y_IN = POLICY["typography"]["margin_in"]["y"] * 2


def _char_em(ch: str) -> float:
    """粗略的字符宽度（em）：全角/CJK 记 1.0，ASCII 数字与字母约 0.55，空格 0.3。"""
    if ch.isspace():
        return 0.3
    code = ord(ch)
    if code < 0x2E80:  # 拉丁/希腊/西里尔等
        if ch.isdigit():
            return 0.56
        if ch.isupper():
            return 0.66
        return 0.52
    return 1.0


def text_fit_risk(el: Dict[str, Any]) -> Optional[str]:
    """估算文本能否装入声明框：逐片段按字宽贪心折行，行高取最大片段字号。装得下返回 None。

    这只是估算：字宽与"单倍行距"（line_pitch_em）都是近似，中文字体的行高常高于该值；
    真实结论以 G-PRE 渲染测量为准，测量到同一文本框后本估算降为 INFO。
    """
    w, h = float(el["w"]), float(el["h"])
    if not el["content"].strip():
        return None
    usable = max(1.0, (w - TEXT_MARGIN_X_IN) * 72.0)
    paragraphs: List[List[float]] = [[]]
    max_pt = 0.0
    for spec in run_styles(el):
        spacing = float(spec["letter_spacing_pt"] or 0.0)
        max_pt = max(max_pt, spec["size"])
        for ch in spec["text"]:
            if ch == "\n":
                paragraphs.append([])
            else:
                paragraphs[-1].append(_char_em(ch) * spec["size"] + spacing)
    lines = 0
    for widths in paragraphs:
        count, acc = 1, 0.0
        for width in widths:
            if acc + width > usable and acc > 0:
                count, acc = count + 1, width
            else:
                acc += width
        lines += count
    pitch = max_pt * text_style(el)["line_spacing"] * float(POLICY["typography"]["line_pitch_em"])
    needed_in = (lines * pitch + (len(paragraphs) - 1) * float(el.get("space_after_pt") or 0.0)) / 72.0 + TEXT_MARGIN_Y_IN
    if needed_in <= h * POLICY["typography"]["text_fit_tolerance"]:
        return None
    return f"估算需 {lines} 行、约 {needed_in:.2f} in 高，文本框仅 {h:.2f} in（行高按 {max_pt:g}pt 计）"


# ---------------------------------------------------------------------------
# 解析与校验
# ---------------------------------------------------------------------------

def _reject_nonstandard_json_number(token: str) -> None:
    raise ValueError(f"JSON 不允许非有限数值：{token}")


def load_plan(path: str) -> Tuple[Optional[Dict[str, Any]], List[CompileError]]:
    errors: List[CompileError] = []
    try:
        with open(path, "r", encoding="utf-8-sig") as f:  # 兼容 Windows 记事本写入的 BOM
            plan = json.load(f, parse_constant=_reject_nonstandard_json_number)
    except FileNotFoundError:
        errors.append(CompileError("E001", ErrorLevel.HARD, path, "plan 文件存在", "文件不存在", "无法读取 plan 文件"))
        return None, errors
    except (OSError, UnicodeDecodeError) as exc:
        errors.append(CompileError("E001", ErrorLevel.HARD, path, "plan 文件可读且为 UTF-8",
                                   f"{type(exc).__name__}: {exc}", "检查路径、权限与文件编码"))
        return None, errors
    except (json.JSONDecodeError, ValueError) as exc:
        errors.append(CompileError("E001", ErrorLevel.HARD, path, "plan 为合法 JSON", f"JSON 解析失败：{exc}", "plan 格式非法"))
        return None, errors
    if not isinstance(plan, dict):
        errors.append(CompileError("E001", ErrorLevel.HARD, path, "plan 为 JSON 对象", f"实际类型：{type(plan).__name__}", "plan 顶层必须是对象"))
        return None, errors
    return plan, errors


def _schema_location(path) -> str:
    parts = list(path)
    if not parts:
        return "$"
    out = "$"
    for part in parts:
        out += f"[{part}]" if isinstance(part, int) else f".{part}"
    return out


def _nonfinite_errors(value: Any, path: str = "$") -> List[CompileError]:
    """jsonschema 的 number 会接受 Python float('nan')；JSON/PPTX 不接受。"""
    errors: List[CompileError] = []
    if isinstance(value, float) and not math.isfinite(value):
        errors.append(CompileError(
            "E001", ErrorLevel.HARD, path, "所有数值均为有限 JSON number", repr(value),
            "NaN/Infinity 不得进入几何、图表或编译载荷",
        ))
    elif isinstance(value, dict):
        for key, item in value.items():
            errors.extend(_nonfinite_errors(item, f"{path}.{key}"))
    elif isinstance(value, list):
        for index, item in enumerate(value):
            errors.extend(_nonfinite_errors(item, f"{path}[{index}]"))
    return errors


def _validate_schema(plan: Dict[str, Any]) -> List[CompileError]:
    errors = _nonfinite_errors(plan)
    raw_schema_errors = list(PLAN_VALIDATOR.iter_errors(plan))
    schema_errors = []
    for candidate in raw_schema_errors:
        path = list(candidate.absolute_path)
        has_more_specific = any(
            other is not candidate
            and len(other.absolute_path) > len(path)
            and list(other.absolute_path)[:len(path)] == path
            for other in raw_schema_errors
        )
        # oneOf/anyOf 只有在候选分支本身因更深层字段失败时才算被更具体的错误解释；
        # 同层的"二选一必填"不能被页面里无关的深层错误掩盖。
        explained_deeper = any(len(c.absolute_path) > len(path) for c in candidate.context)
        if candidate.validator in {"oneOf", "anyOf"} and has_more_specific and explained_deeper:
            continue
        schema_errors.append(candidate)
    schema_errors.sort(key=lambda e: (-len(e.absolute_path), tuple(str(p) for p in e.absolute_path), e.message))
    for error in schema_errors:
        path = list(error.absolute_path)
        location = _schema_location(error.absolute_path)
        missing = None
        if error.validator == "required":
            match = re.match(r"'([^']+)' is a required property", error.message)
            if match:
                missing = match.group(1)
                location += f".{missing}"
        code = "E001"
        if missing in {"chart", "data"} and ("chart" in path or missing == "chart"):
            code = "E004"
        elif "chart" in path:
            code = "E007"
        elif path and path[-1] in {"x", "y", "w", "h"} and error.validator in {"minimum", "exclusiveMinimum"}:
            code = "E002"
        actual = error.message
        if error.validator in {"oneOf", "anyOf"}:
            # jsonschema 的原文会整段复述实例（整页 JSON）；改为引用 Schema 对该规则的说明与候选约束。
            rule = error.schema.get("description") if isinstance(error.schema, dict) else None
            actual = "同时满足多个互斥候选" if error.validator == "oneOf" and not error.context else "不满足任何候选"
            if not error.context:
                actual += "：" + " / ".join(json.dumps(branch, ensure_ascii=False) for branch in error.validator_value)
            if rule:
                actual = f"{rule}；{actual}"
        elif error.validator == "not" and error.validator_value == {}:
            actual = f"字段「{path[-1]}」在此元素类型/图型上没有运行时语义，不得出现（实际值 {error.instance!r}）"
        elif path and path[-1] == "kind" and "chart" in path:
            actual = f"kind={error.instance!r}：{actual}"
        elif "values" in path and error.validator == "type":
            actual = f"系列含不可解析数值 {error.instance!r}：{actual}"
        if error.context:
            actual += "；候选约束：" + " / ".join(e.message for e in error.context[:2])
        errors.append(CompileError(
            code, ErrorLevel.HARD, location,
            "plan 必须符合 scripts/plan.schema.json", actual,
            "结构契约失败；修正字段、类型、枚举或跨字段条件后再进入 BUILD",
        ))
    return errors


def _validate_handoff(plan: Dict[str, Any], errors: List[CompileError]) -> None:
    """只做 Schema 表达不了的交接检查：跨权威源的可复核代码、页内引用与配色身份。

    非空白、空 claim 的未决说明、证据状态与 evidence/open_questions 的对应关系都由 Schema 唯一定义。
    """
    for index, record in enumerate(plan.get("design_reviews", [])):
        if record["code"] not in POLICY["diagnostics"]["reviewable_codes"]:
            errors.append(CompileError("E001", ErrorLevel.HARD, f"deck.design_reviews[{index}]",
                                       "仅可复核的诊断码能具名复核", record["code"],
                                       "硬错、与本页数据的矛盾（E008）与合规不得用复核记录豁免"))
    palette = plan.get("palette") or []
    folded = [c.upper() for c in palette]
    if len(set(folded)) != len(folded):
        errors.append(CompileError(
            "E001", ErrorLevel.HARD, "deck.palette", "配色在忽略大小写后仍须各不相同",
            repr(palette), "同一颜色写两遍会让两个系列/类别共用一色而无法区分",
        ))
    for pi, page in enumerate(plan["pages"]):
        loc = f"page[{pi}]({page['id']})"
        focus_el = str(page.get("focus_element") or "").strip()
        if focus_el and focus_el not in {el["id"] for el in page["elements"]}:
            errors.append(CompileError(
                "E001", ErrorLevel.HARD, f"{loc}.focus_element",
                "focus_element 必须是本页存在元素 id",
                repr(focus_el), "把主角指到承载它的那个元素；确属并列或跨多个元素时改用 focus_unmapped 写明理由，说不出理由说明这一页还没有主角",
            ))


def chart_color_count(chart: Dict[str, Any]) -> int:
    """图表实际消费的独立颜色数：编译着色与契约校验共用同一口径。"""
    kind, data = chart["kind"], chart["data"]
    if kind == "table":
        return 0
    if kind in {"pie", "donut"}:
        return len(data["categories"])
    return len(data["series"])


def _validate_chart_contract(el: Dict[str, Any], loc: str, palette: List[str], palette_source: str,
                             errors: List[CompileError]) -> None:
    """Schema 管形状与按图型可用的字段；这里只查需要比对数据长度、跨字段语义或配色存量的约束。"""
    chart = el["chart"]
    kind = chart["kind"]
    data = chart["data"]

    def _err(code, level, where, condition, actual, message):
        errors.append(CompileError(code, level, f"{loc}.chart{where}", condition, actual, message))

    if "headers" in data:
        headers = data["headers"]
        for ri, row in enumerate(data["rows"]):
            if len(row) != len(headers):
                _err("E007", ErrorLevel.HARD, f".data.rows[{ri}]", "表格每行单元格数必须等于 headers",
                     f"实际 {len(row)}，headers {len(headers)}", "表格载荷不完整，禁止截断或补造单元格")
        column_count = len(headers)
        series_list: List[Dict[str, Any]] = []
    else:
        categories = data["categories"]
        series_list = data["series"]
        column_count = 1 + len(series_list)
        labels = [str(c) for c in categories]
        if any(not label.strip() for label in labels) or len(set(labels)) != len(labels):
            _err("E007", ErrorLevel.HARD, ".data.categories", "类别标签非空且字符串化后唯一",
                 repr(categories), "类别身份不清会导致数值错配")
        for si, series in enumerate(series_list):
            if len(series["values"]) != len(categories):
                _err("E007", ErrorLevel.HARD, f".data.series[{si}].values", "系列长度必须等于类别数",
                     f"长度 {len(series['values'])}，类别数 {len(categories)}", "禁止丢弃、插值或补造数据点")
            if kind in {"pie", "donut"} and sum(series["values"]) <= 0:
                _err("E007", ErrorLevel.HARD, f".data.series[{si}].values", "饼/环图必须具有正的有效总量",
                     repr(series["values"]), "零总量无法计算占比")
            for field, detail in _declared_scope_conflicts(chart, series):
                _err("E008", ErrorLevel.HARD, f".data.series[{si}].{field}",
                     "图表级与系列级声明不得互相矛盾（图题只能显示一种单位/期间）", detail,
                     "统一声明，或删去图表级缺省并把差异写进系列名；不得让图题标注与所画数据不符")
        highlight = chart.get("highlight_category")
        if highlight is not None and (len(series_list) != 1 or highlight >= len(categories)):
            _err("E007", ErrorLevel.HARD, ".highlight_category", "highlight_category 仅用于单系列柱图且指向存在的类别",
                 f"series={len(series_list)}，index={highlight}，categories={len(categories)}",
                 "无法把强调准确执行到声称的那一项；请修正或撤回高亮声明")

    weights = chart.get("column_weights")
    if weights is not None and len(weights) != column_count:
        _err("E007", ErrorLevel.HARD, ".column_weights", "列权重数量必须等于实际列数",
             f"权重 {len(weights)} 个，列 {column_count} 个", "按列逐项给出权重；数量不符时不会被静默忽略")

    title_conflict = _title_unit_conflict(chart)
    if title_conflict:
        _err("E008", ErrorLevel.HARD, ".title", "图题括注的单位须与数据单位一致", title_conflict,
             "观众会同时看到两种单位：改图题或改数据单位声明")

    if len(series_list) > 1 and kind != "table":
        if chart.get("shared_basis") is not True:
            # 未声明与显式 false 都是口径未获确认：同一事实边界，同一个 HARD 级别。
            _err("E008", ErrorLevel.HARD, ".shared_basis", "多系列共轴前必须显式确认 shared_basis=true",
                 repr(chart.get("shared_basis")),
                 "不同或未知统计口径不得仅因可绘制就自动共轴；请确认口径、分拆或改用表格")
        scopes = [_resolve_series_scope(chart, s) for s in series_list]
        undeclared = [s["series_name"] for s in scopes if not s["unit"]]
        declared = {(_parse_unit(s["unit"]).scale, _parse_unit(s["unit"]).base) for s in scopes if s["unit"]}
        if undeclared and declared:
            _err("E008", ErrorLevel.HARD, ".data.series", "共轴系列要么都声明单位，要么都不声明",
                 f"未声明单位的系列：{undeclared}", "在图表级声明共同单位，或为每个系列声明；不得让坐标轴单位处于半知状态")
        elif len(declared) > 1:
            _err("E008", ErrorLevel.HARD, ".data.series", "共轴系列必须具有相同量纲与数量级",
                 repr(sorted({s["unit"] for s in scopes})), "不得静默换算或混合单位；先在数据中完成有依据的换算，或分拆呈现")
        bases = {s["basis"] for s in scopes if s["basis"]}
        if chart.get("shared_basis") is True and len(bases) > 1:
            _err("E008", ErrorLevel.HARD, ".shared_basis", "shared_basis=true 时各系列已声明的口径必须相同",
                 repr(sorted(bases)), "权威声明互相矛盾；确认统计口径后再编译")

    needed = chart_color_count(chart)
    if needed > len(palette):
        _err("E007", ErrorLevel.HARD, "", f"图表需要 {needed} 个独立颜色",
             f"可用配色 {len(palette)} 个（来自 {palette_source}）",
             "由艺术指导补齐配色或减少类别；编译器不循环复用颜色")


def substantive_overlap(a: Dict[str, Any], b: Dict[str, Any]) -> float:
    """两个元素框的实质重叠面积（sq in），不构成实质重叠时为 0。阈值只来自 policy.geometry。"""
    g = POLICY["geometry"]
    ox = max(0.0, min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"]))
    oy = max(0.0, min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"]))
    if ox <= g["overlap_min_edge_in"] or oy <= g["overlap_min_edge_in"]:
        return 0.0
    area, min_area = ox * oy, min(a["w"] * a["h"], b["w"] * b["h"])
    return area if area > g["overlap_area_sq_in"] or area / min_area > g["overlap_ratio"] else 0.0


def _validate_image_handoff(page: Dict[str, Any], assets: Dict[str, Any],
                            slide_w_in: float, slide_h_in: float,
                            errors: List[CompileError]) -> None:
    """E012（交接完整性）：只查可由 plan 几何确定之事——大图/叠字图是否交代了图像简报。

    简报写得好坏、图片是否真的安静，程序不判；但"文字实质压在图片上却从未交代落字区"
    是出图前就注定的事故，必须在 BUILD 前被看见。
    """
    cfg = POLICY["image"]
    slide_area = slide_w_in * slide_h_in
    texts = [e for e in page["elements"] if e["type"] in TEXT_ELEMENT_TYPES]
    for el in page["elements"]:
        if el["type"] != "image":
            continue
        loc = f"page[{page['id']}].element({el['id']})"
        record = assets.get(el["image_id"])
        if record is None:
            continue  # E011 已在 _validate_asset 阻断
        role = el.get("image_role")
        area_ratio = (float(el["w"]) * float(el["h"])) / slide_area if slide_area else 0.0
        overlaps_text = any(substantive_overlap(el, t) for t in texts)
        large = area_ratio >= cfg["large_area_ratio"]
        if not (large or overlaps_text or role in {"background", "subject", "evidence"}):
            continue
        brief = record.get("image_brief") or {}
        def _filled(key: str) -> bool:
            return bool(str(brief.get(key) or "").strip())
        missing = []
        if not _filled("visual_intent"):
            missing.append("visual_intent")
        # evidence 是最需要可核查的一类素材：无论面积大小，主体景别与淘汰条件都不能缺。
        if (large or role in {"subject", "detail", "evidence"}) and not _filled("subject_and_shot"):
            missing.append("subject_and_shot")
        if overlaps_text and not _filled("text_safe_zone"):
            missing.append("text_safe_zone")
        if role == "evidence" and not brief.get("reject_if"):
            missing.append("reject_if")
        if missing:
            errors.append(CompileError(
                "E012", ErrorLevel.SOFT, f"{loc}.image_brief",
                "大图、主图、证据图或叠字图必须以 image_brief 交代意图、主体景别、落字安静区与淘汰条件",
                f"缺 {missing}（面积占比 {area_ratio:.0%}，叠字={overlaps_text}，role={role}）",
                "先按最终落位框补齐简报再生成/选材；不得先有图后凑版面",
            ))
        elif overlaps_text and str(brief.get("text_safe_zone", "")).strip().lower().startswith("none"):
            errors.append(CompileError(
                "E012", ErrorLevel.SOFT, f"{loc}.image_brief.text_safe_zone",
                "文字实质压图时不得同时声明无安静区", str(brief.get("text_safe_zone")),
                "要么移动文字/图片消除叠压，要么按最终裁切如实交代安静区位置",
            ))


def _validate_asset(el: Dict[str, Any], assets: Dict[str, Any], loc: str, errors: List[CompileError]) -> None:
    image_id = el["image_id"]
    record = assets.get(image_id)
    if record is None:
        errors.append(CompileError(
            "E011", ErrorLevel.HARD, f"{loc}.image_id",
            "图片必须引用 deck.assets 中同名的素材记录", image_id,
            "素材来源、授权、隐私与生成属性未交接，G-COMPLY 不得推定通过",
        ))
        return
    if el.get("image_role") == "evidence" and not record["used_as_evidence"]:
        errors.append(CompileError(
            "E011", ErrorLevel.HARD, f"{loc}.image_role",
            "image_role=evidence 必须由素材记录 used_as_evidence=true 承认",
            f"assets[{image_id}].used_as_evidence=false（generated={record['generated']}）",
            "证据用途要在素材记录登记并接受合规核对；生成图不得登记为证据（Schema 禁止）",
        ))
    for field in ("license_status", "privacy_status"):
        status = record[field]
        if status != "cleared":
            errors.append(CompileError(
                "E011", ErrorLevel.HARD if status == "blocked" else ErrorLevel.SOFT,
                f"deck.assets[{image_id}].{field}", "素材授权/隐私必须为 cleared 才可交付",
                status, "待确认项为 NOT VERIFIED；blocked 为硬阻断，均不得由风险批准绕过",
            ))


def resolved_palette(plan: Dict[str, Any]) -> List[str]:
    """图表消费的配色：plan 声明优先，否则是 policy 的唯一默认；编译与校验共用。"""
    return list(plan["palette"]) if "palette" in plan else list(POLICY["defaults"]["chart_palette"])


def _validate_page_numbers(plan: Dict[str, Any], errors: List[CompileError]) -> None:
    """编译前的页码规则、覆盖范围与编号合法性检查（E017）。

    Schema 管字段形状；这里管跨对象事实：例外要有理由、省略页要真实存在、
    页码框在画布内、页码不与正文/形状/图表的声明几何实质重叠。页码与背景图的关系
    交给 G-PRE/G-SEM 的真实合成背景对比核验，不在这里用几何猜测。
    """
    cfg = page_number_config(plan)
    pages = plan["pages"]
    ids = [p["id"] for p in pages]
    loc = "deck.page_numbers"
    if cfg["corner"] not in {"top_right", "bottom_right"}:
        errors.append(CompileError("E017", ErrorLevel.HARD, loc, "页码角区必须是右上或右下",
                                   cfg["corner"], "整稿只能选择一个角区；不得逐页切换"))
    if cfg["format"] not in {"plain", "zero_padded"}:
        errors.append(CompileError("E017", ErrorLevel.HARD, loc, "页码格式必须是 plain 或 zero_padded",
                                   cfg["format"], "整稿格式必须一致"))
    if cfg["start"] < 1:
        errors.append(CompileError("E017", ErrorLevel.HARD, loc, "起始页码 ≥ 1",
                                   str(cfg["start"]), "自定义起始号必须显式声明且为正整数"))
    if (not cfg["enabled"] or cfg["omit_pages"]) and not cfg["omit_reason"].strip():
        errors.append(CompileError("E017", ErrorLevel.HARD, loc, "省略页码必须写出任务理由",
                                   f"enabled={cfg['enabled']}, omit_pages={cfg['omit_pages']}",
                                   "页码是功能性信息；只有任务明确要求例外时才允许省略并留痕"))
    if len(set(cfg["omit_pages"])) != len(cfg["omit_pages"]):
        errors.append(CompileError("E017", ErrorLevel.HARD, loc, "omit_pages 不得重复",
                                   repr(cfg["omit_pages"]), "重复页面 id 会让省略范围不可追溯"))
    for pid in cfg["omit_pages"]:
        if pid not in ids:
            errors.append(CompileError("E017", ErrorLevel.HARD, loc, "omit_pages 必须指向真实页面",
                                       pid, "引用不存在的页面等于没有绑定任何例外"))
    sw, sh = _slide_dimensions(plan.get("aspect_ratio"))
    for i, (label, el) in enumerate(page_number_specs(plan)):
        if el is None:
            continue
        where = f"page[{ids[i]}].page_number"
        if (el["x"] < 0 or el["y"] < 0 or el["x"] + el["w"] > sw + BOUND_TOL_IN
                or el["y"] + el["h"] > sh + BOUND_TOL_IN):
            errors.append(CompileError("E017", ErrorLevel.HARD, where, "页码框必须完整位于画布内",
                                       f"({el['x']:.2f},{el['y']:.2f},{el['w']:.2f}x{el['h']:.2f}) / {sw}x{sh}",
                                       "调整 policy 的 margin/box；不得靠缩小或裁切规避"))
        for other in pages[i]["elements"]:
            if other["type"] == "image":  # 背景/主体图由真实渲染的对比与裁切核验
                continue
            if substantive_overlap(el, other) and not other.get("overlap_allowed"):
                errors.append(CompileError(
                    "E017", ErrorLevel.HARD, where, "页码不得与正文、形状或图表实质重叠",
                    f"与 {other['id']}（{other['type']}）重叠；页码文字 {label!r}",
                    "移动内容或明确声明 overlap_allowed；不得删除页码来腾空间"))


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _validate_generation(plan: Dict[str, Any], errors: List[CompileError],
                         asset_root: Optional[str] = None) -> None:
    """图像生成素材的登记完整性（E011）：未完成登记不得进入交付，缺记录不得冒充完成。

    智能体负责真实调用图像模型并写记录；编译器只做确定性核对，不发起任何网络请求。
    给出 asset_root 时，登记 file 必须解析到页面 image_id 的**同一物理文件**，且声明的
    sha256 必须等于该文件的实际哈希——只比文件名或 stem 不足以证明来源。
    """
    assets = plan.get("assets", {})
    used = {el["image_id"] for page in plan["pages"] for el in page["elements"] if el["type"] == "image"}
    for image_id, record in assets.items():
        gen = record.get("image_generation")
        loc = f"deck.assets[{image_id}].image_generation"
        if gen is None:
            if record.get("generated") is True and image_id in used:
                errors.append(CompileError(
                    "E011", ErrorLevel.SOFT, loc, "生成素材必须登记实际图像生成调用记录",
                    "generated=true 但缺少 image_generation",
                    "补写 tool/prompt/file/status/sha256；未登记时 G-COMPLY 记 NOT VERIFIED，不得视为已完成素材"))
            continue
        status = gen.get("status")
        if status != "generated":
            if image_id in used:
                errors.append(CompileError(
                    "E011", ErrorLevel.HARD, loc, "未完成的图像生成素材不得进入交付",
                    f"status={status!r}", "先完成生成、保存与登记，或从页面移除该图；不得用占位图冒充真实素材"))
            continue
        if record.get("generated") is not True or record.get("used_as_evidence") is True:
            errors.append(CompileError(
                "E011", ErrorLevel.HARD, loc, "已生成素材必须 generated=true 且 used_as_evidence=false",
                f"generated={record.get('generated')}, used_as_evidence={record.get('used_as_evidence')}",
                "生成图不是证据；登记与实际素材属性必须一致"))
        declared = str(gen.get("file") or "")
        actual_image = _resolve_image_path(image_id, asset_root) if asset_root else None
        if actual_image is not None:
            declared_path = Path(declared)
            if not declared_path.is_absolute():
                declared_path = Path(asset_root).resolve() / declared_path
            try:
                same_file = declared_path.resolve() == actual_image.resolve()
            except OSError:
                same_file = False
            if not same_file:
                errors.append(CompileError(
                    "E011", ErrorLevel.HARD, loc, "生成登记文件必须与页面 image_id 指向同一文件",
                    f"file={declared!r} → {declared_path}，image_id={image_id!r} 实际解析为 {actual_image}",
                    "登记真实落盘文件（含扩展名与相对素材根的路径）；名称相似但扩展名不同的文件不能互相顶替"))
            elif gen.get("sha256"):
                try:
                    actual_sha = _file_sha256(actual_image)
                except OSError as exc:
                    errors.append(CompileError(
                        "E011", ErrorLevel.HARD, loc, "生成登记文件必须可读并核对实际哈希",
                        f"{actual_image}: {exc}", "修复文件权限或重新落盘后再登记"))
                else:
                    if actual_sha.lower() != str(gen["sha256"]).lower():
                        errors.append(CompileError(
                            "E011", ErrorLevel.HARD, loc, "生成登记 sha256 必须等于实际素材文件哈希",
                            f"file={actual_sha}，登记={gen['sha256']}",
                            "重新计算真实文件哈希后登记；不得沿用旧图或其他版本的哈希"))
        elif used and image_id in used:
            errors.append(CompileError(
                "E011", ErrorLevel.HARD, loc, "已使用生成素材必须能在素材根目录解析到真实文件",
                f"file={declared!r}, image_id={image_id!r}, assets={asset_root!r}",
                "确认 --assets 指向真实素材目录，并登记实际落盘文件"))
        elif Path(declared).name != Path(image_id).name and Path(declared).stem != Path(image_id).stem:
            # 无 asset_root 时仍拒绝明显不同的文件名；这只是弱检查，不能替代哈希闭环。
            errors.append(CompileError(
                "E011", ErrorLevel.HARD, loc, "生成登记文件必须与页面 image_id 指向同一文件",
                f"file={declared!r}, image_id={image_id!r}",
                "登记真实落盘文件名；验证器还会用 sha256 核对 PPTX 内嵌图片"))


def validate_plan(plan: Dict[str, Any], asset_root: Optional[str] = None) -> List[CompileError]:
    """写 PPTX 之前的全部契约检查；任何 HARD 都意味着不应生成交付文件。

    asset_root 可选：提供时才能核对生成素材登记文件是否真实存在、与 image_id 同一文件、
    且声明 sha256 等于实际文件哈希；不提供时只做结构检查，不能把登记当作已证明。
    """
    errors = _validate_schema(plan)
    if errors:
        return errors

    _validate_handoff(plan, errors)
    _validate_page_numbers(plan, errors)
    _validate_generation(plan, errors, asset_root)
    w_limit, h_limit = _slide_dimensions(plan.get("aspect_ratio"))
    palette = resolved_palette(plan)

    seen_page_ids: Set[str] = set()
    for pi, page in enumerate(plan["pages"]):
        pid = page["id"]
        if pid in seen_page_ids:
            errors.append(CompileError(
                "E001", ErrorLevel.HARD, f"page[{pi}].id", "页面 id 整稿唯一", pid,
                "重复页面身份会破坏对象追溯",
            ))
        seen_page_ids.add(pid)
        elements = page["elements"]
        _validate_elements(page, errors, w_limit, h_limit, plan.get("assets", {}), palette,
                           "plan.palette" if "palette" in plan else "policy 默认配色")
        _validate_image_handoff(page, plan.get("assets", {}), w_limit, h_limit, errors)
        _validate_claim_evidence(page, errors)
    return errors


# 文本里出现字面转义序列（反斜杠 + n/t/r）是内容 bug：作者想要换行，却把 "\\n"
# 当成两个可见字符写进了 content。它会在 PPTX 与 PDF 里显示成 "\n" 而非换行，
# 单靠渲染事后看图才发现的缺陷，应在 BUILD 前确定性阻断。真正换行用 U+000A。
_LITERAL_ESCAPE_RE = re.compile(r"\\[ntr]")


def _validate_elements(
    page: Dict[str, Any],
    errors: List[CompileError],
    slide_w_in: float,
    slide_h_in: float,
    assets: Dict[str, Any],
    palette: List[str],
    palette_source: str,
) -> None:
    pid = page["id"]
    elements = page["elements"]
    placed: List[Dict[str, Any]] = []
    seen_element_ids: Set[str] = set()

    for ei, el in enumerate(elements):
        eid = el["id"]
        loc = f"page[{pid}].element[{ei}]({eid})"
        if eid in seen_element_ids:
            errors.append(CompileError(
                "E001", ErrorLevel.HARD, f"{loc}.id", "元素 id 页内唯一", eid,
                "重复元素身份会破坏 plan→PPTX 对象追溯",
            ))
        seen_element_ids.add(eid)
        if eid == PAGE_NUMBER_ID:
            errors.append(CompileError(
                "E017", ErrorLevel.HARD, f"{loc}.id", "元素 id 不得占用编译器保留的页码对象名",
                eid, "页码由整稿 page_numbers 契约派生；手写同名元素会造成对象身份与编号冲突",
            ))

        etype = el["type"]
        x, y, w, h = (float(el[k]) for k in ("x", "y", "w", "h"))
        if x + w > slide_w_in + BOUND_TOL_IN or y + h > slide_h_in + BOUND_TOL_IN:
            errors.append(CompileError(
                "E002", ErrorLevel.HARD, loc, "元素在幻灯片边界内",
                f"({x:.2f},{y:.2f},{w:.2f}x{h:.2f}) 超出边界 {slide_w_in}x{slide_h_in}",
                "元素超出幻灯片可绘制区域，阻断编译",
            ))
        else:
            placed.append(el)

        if etype in TEXT_ELEMENT_TYPES:
            runs = el.get("runs")
            if runs:
                joined = "".join(str(r.get("text", "")) for r in runs)
                content = el.get("content") or ""
                if "\n" in content or joined != content:
                    errors.append(CompileError(
                        "E001", ErrorLevel.HARD, f"{loc}.runs",
                        "runs 拼接必须与 content 逐字相等且 content 为单段",
                        f"runs 拼接={joined!r}，content={content!r}",
                        "双源文本必有一处漂移；以 content 为唯一事实源重新切分 runs",
                    ))
            # 字面 \n / \t / \r 是可见的错误字符，不是换行。逐一检查 content 与每个
            # run.text，避免它们被当成正常文本写进原生对象、直到渲染才暴露。
            visible_texts = [("content", str(el.get("content") or ""))]
            if runs:
                visible_texts += [(f"runs[{ri}].text", str(r.get("text") or ""))
                                  for ri, r in enumerate(runs)]
            for field, value in visible_texts:
                m = _LITERAL_ESCAPE_RE.search(value)
                if m:
                    errors.append(CompileError(
                        "E001", ErrorLevel.HARD, f"{loc}.{field}",
                        "可见文本不得含字面转义序列（\\n / \\t / \\r）",
                        f"发现 {m.group(0)!r}，位于 {value[max(0, m.start()-8):m.start()+8]!r}",
                        "换行请写真正的换行字符 U+000A（JSON 中表示为单个换行），不要把反斜杠加字母"
                        "这两个可见字符当作换行；否则 PPTX/PDF 会把它们显示成字面文字",
                    ))
                    break
            risk = text_fit_risk(el)
            if risk:
                errors.append(CompileError(
                    "E010", ErrorLevel.SOFT, f"{loc}.content",
                    "文本应能装入声明的文本框（不靠缩字号硬塞）", risk,
                    "文本可能溢出文本框：优先编辑删减、调整断行或放大框体，勿压缩字号与行距",
                ))
            floor = min_pt_for(etype)
            smallest = min(style["size"] for style in run_styles(el))
            if smallest < floor:
                role = next(name for name, sizes in MIN_PT_SETS.items() if etype in sizes)
                errors.append(CompileError(
                    "E005", ErrorLevel.SOFT, f"{loc}.font_size",
                    f"{etype} 属「{role}」角色，字号 ≥ {floor}pt",
                    f"实际最小：{smallest:g}pt",
                    "投影下不可读：放大字号或改换元素角色；不能靠压字距、缩行距或挪到页脚掩盖",
                ))

        if etype == "chart":
            _validate_chart_contract(el, loc, palette, palette_source, errors)
        elif etype == "image":
            _validate_asset(el, assets, loc, errors)

    for i, a in enumerate(placed):
        for b in placed[i + 1:]:
            area = substantive_overlap(a, b)
            # z_index 相同则按元素声明顺序叠放；只有实际下层可以授权承载上层。
            default_z = float(POLICY["defaults"]["element"]["z_index"])
            lower = a if float(a.get("z_index", default_z)) <= float(b.get("z_index", default_z)) else b
            if not area or lower.get("overlap_allowed"):
                continue
            errors.append(CompileError(
                "E002", ErrorLevel.HARD, f"page[{pid}].overlap",
                "实质重叠必须由下层承载元素显式 overlap_allowed=true",
                f"{a['id']} 与 {b['id']} 重叠面积约 {area:.2f} sq in",
                "不得按全屏面积推断背景，也不得由上层元素自行豁免；请修正几何或明确图层关系",
            ))


# ---------------------------------------------------------------------------
# PPTX 渲染
# ---------------------------------------------------------------------------

def _hex_to_rgb(hex_color: str) -> RGBColor:
    """颜色必须已由 plan 或 policy 解析为 #RRGGBB；非法或缺失明确报错，不得换色继续交付。"""
    if not isinstance(hex_color, str) or not re.fullmatch(r"#[0-9A-Fa-f]{6}", hex_color):
        raise ValueError(f"非法颜色：{hex_color!r}；应为 #RRGGBB")
    h = hex_color[1:]
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _rel_luminance(hex_color: str) -> float:
    """计算 HEX 颜色的 WCAG 相对亮度（0.0=纯黑，1.0=纯白）。"""
    if not isinstance(hex_color, str) or not re.fullmatch(r"#[0-9A-Fa-f]{6}", hex_color):
        raise ValueError(f"非法颜色：{hex_color!r}；应为 #RRGGBB")
    h = hex_color[1:]
    chans = []
    for idx in (0, 2, 4):
        s = int(h[idx:idx + 2], 16) / 255.0
        chans.append(s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4)
    return 0.2126 * chans[0] + 0.7152 * chans[1] + 0.0722 * chans[2]


def page_background(page: Dict[str, Any]) -> str:
    """页面实际底色：plan 声明优先，否则 policy 默认。编译器总是显式写入，不依赖模板继承。"""
    return str(page.get("background") or POLICY["defaults"]["background"]).upper()


# ---------------------------------------------------------------------------
# 页码：整稿级契约 → 每页一个确定性派生的原生文本对象
#
# 位置/样式/编号只在 policy.page_number 与 plan.page_numbers 定义一次；编译器按 pages
# 顺序生成，验证器用同一组派生函数核对实际 PPTX 对象与真实渲染。页码不是作者可删的装饰。
# ---------------------------------------------------------------------------

PAGE_NUMBER_ID = "_page_number"   # 保留 id：作者元素不得占用，避免对象名碰撞


def page_number_config(plan: Dict[str, Any]) -> Dict[str, Any]:
    """整稿页码契约：plan 的声明优先，缺省值只来自 policy。"""
    cfg = POLICY["page_number"]
    override = plan.get("page_numbers") or {}
    return {
        "corner": str(override.get("corner", cfg["corner"])),
        "format": str(override.get("format", cfg["format"])),
        "start": int(override.get("start", cfg["start"])),
        "enabled": bool(override.get("enabled", cfg["enabled"])),
        "omit_pages": list(override.get("omit_pages") or []),
        "omit_reason": str(override.get("omit_reason") or ""),
    }


def page_number_label(plan: Dict[str, Any], index: int) -> Optional[str]:
    """第 index 页（0 起）的页码文字；省略页返回 None。补零宽度按末页号取整稿一致。"""
    cfg = page_number_config(plan)
    if not cfg["enabled"] or plan["pages"][index]["id"] in cfg["omit_pages"]:
        return None
    number = cfg["start"] + index
    if cfg["format"] == "plain":
        return str(number)
    last = cfg["start"] + len(plan["pages"]) - 1
    width = max(int(POLICY["page_number"]["zero_pad_min_width"]), len(str(last)))
    return f"{number:0{width}d}"


def page_number_element(plan: Dict[str, Any], index: int) -> Optional[Dict[str, Any]]:
    """派生页码文本元素（与 plan 元素同构）。编译写入与交付核对共用同一份规格。"""
    label = page_number_label(plan, index)
    if label is None:
        return None
    page = plan["pages"][index]
    cfg = POLICY["page_number"]
    sw, sh = _slide_dimensions(plan.get("aspect_ratio"))
    w, h = float(cfg["box_in"]["w"]), float(cfg["box_in"]["h"])
    x = sw - float(cfg["margin_in"]["x"]) - w
    y = (float(cfg["margin_in"]["y"]) if page_number_config(plan)["corner"] == "top_right"
         else sh - float(cfg["margin_in"]["y"]) - h)
    tone = page.get("page_number_tone")
    dark = ((tone == "dark") if tone else
            _rel_luminance(page_background(page)) < POLICY["chart"]["dark_background_luminance"])
    return {
        "id": PAGE_NUMBER_ID, "type": "footer", "content": label,
        "x": x, "y": y, "w": w, "h": h,
        "font_size": float(cfg["font_size_pt"]),
        "color": str(cfg["ink"]["dark_page" if dark else "light_page"]).upper(),
        "bold": False, "align": "left", "z_index": float("inf"),
    }


def page_number_specs(plan: Dict[str, Any]) -> List[Tuple[Optional[str], Optional[Dict[str, Any]]]]:
    """每页的 (可见页码文字, 派生元素)；省略页为 (None, None)。"""
    return [(None if el is None else el["content"], el)
            for el in (page_number_element(plan, i) for i in range(len(plan["pages"])))]


def chart_tone(bg_hex: str) -> Tuple[bool, Dict[str, str]]:
    """按页面底色选择图表墨阶档（深/浅）；编译与交付核对共用一个判据。"""
    dark = _rel_luminance(bg_hex) < POLICY["chart"]["dark_background_luminance"]
    return dark, POLICY["chart"]["style"]["dark" if dark else "light"]


def _apply_fill_alpha(shape, alpha) -> None:
    """为形状的纯色填充写入 a:alpha（不透明度）；alpha=None 时不写。"""
    if alpha is None:
        return
    srgb = shape.fill._xPr.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
    if srgb is None:  # 调用方必须先写纯色填充；找不到就是实现漂移，不能静默丢掉透明度
        raise RuntimeError("fill_alpha 需要已写入的 a:solidFill/a:srgbClr")
    for old in srgb.findall(qn("a:alpha")):
        srgb.remove(old)
    el = srgb.makeelement(qn("a:alpha"), {"val": str(int(round(float(alpha) * 100000)))})
    srgb.append(el)


def _set_cell_hairlines(
    cell,
    top_hex: Optional[str] = None,
    top_pt: float = 0.0,
    bottom_hex: Optional[str] = None,
    bottom_pt: float = 0.0,
) -> None:
    """按 ECMA-376 CT_TableCellProperties 元素顺序写入无竖线、纯水平发丝线与透明底场域。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)

    def _make_ln(tag: str, color_hex: Optional[str], width_pt: float):
        ln = tcPr.makeelement(qn(tag), {"w": str(int(round(width_pt * 12700))), "cap": "flat"})
        if color_hex and width_pt > 0:
            sf = ln.makeelement(qn("a:solidFill"), {})
            sf.append(sf.makeelement(qn("a:srgbClr"), {"val": color_hex.lstrip("#").upper()}))
            ln.append(sf)
        else:
            ln.append(ln.makeelement(qn("a:noFill"), {}))
        return ln

    # 严格置于 fill 元素之前，符合 OOXML XSD 序列（lnL, lnR, lnT, lnB -> noFill）
    cell.fill.background()
    no_fill = tcPr.find(qn("a:noFill"))
    borders = [
        _make_ln("a:lnL", None, 0.0),
        _make_ln("a:lnR", None, 0.0),
        _make_ln("a:lnT", top_hex, top_pt),
        _make_ln("a:lnB", bottom_hex, bottom_pt),
    ]
    if no_fill is not None:
        idx = list(tcPr).index(no_fill)
        for offset, b_el in enumerate(borders):
            tcPr.insert(idx + offset, b_el)
    else:
        for b_el in borders:
            tcPr.append(b_el)


def _apply_run_font(run, font_name: str, size_pt: float, color_hex: Optional[str], bold: bool = False,
                    letter_spacing_pt: Optional[float] = None) -> None:
    """按 ECMA-376 规范同时写入西文 <a:latin> 与东亚 <a:ea> 字体定义；字距写入 rPr/@spc（1/100 pt）。"""
    run.font.size = Pt(size_pt)
    run.font.color.rgb = _hex_to_rgb(color_hex)
    if bold:
        run.font.bold = True
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    if letter_spacing_pt is not None:
        rPr.set("spc", str(int(round(letter_spacing_pt * 100))))
    existing_ea = rPr.find(qn("a:ea"))
    if existing_ea is not None:
        rPr.remove(existing_ea)
    ea = rPr.makeelement(qn("a:ea"), {"typeface": font_name})
    latin = rPr.find(qn("a:latin"))
    if latin is not None:
        latin.addnext(ea)
    else:
        rPr.append(ea)


def _ensure_ea_in_txpr(element, font_name: str) -> None:
    """在 c:txPr 的每个 a:defRPr 上补写 a:ea，使图表文字（图例/坐标轴/直标）也拿到东亚字体。

    python-pptx 的 font.name 只写 <a:latin>；中文图表标签若缺 a:ea，目标环境会回退到
    另一套字族，和页面其余文字明显不同族——这正是"整稿统一"最容易被忽略的破口。
    """
    if element is None:
        return
    for txPr in element.iter(qn("c:txPr")):
        for defRPr in txPr.iter(qn("a:defRPr")):
            for old in defRPr.findall(qn("a:ea")):
                defRPr.remove(old)
            defRPr.append(defRPr.makeelement(qn("a:ea"), {"typeface": font_name}))


def plain_text_style(size: float, color: str, font: str, bold: bool = False) -> Dict[str, Any]:
    """派生文字对象（表题、诊断占位）的显式式样；字段与 text_style 同构。"""
    return {"size": float(size), "bold": bold, "font": font, "color": color.upper(),
            "line_spacing": float(POLICY["typography"]["line_spacing"]["body"]),
            "letter_spacing_pt": None, "valign": "top"}


def _add_textbox(
    slide,
    box: Tuple[float, float, float, float],
    content: str,
    style: Dict[str, Any],
    *,
    align: Optional[str] = None,
    runs: Optional[List[Dict[str, Any]]] = None,
    space_after_pt: Optional[float] = None,
    object_name: Optional[str] = None,
):
    x, y, w, h = box
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if object_name:
        tb.name = object_name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(POLICY["typography"]["margin_in"]["x"])
    tf.margin_right = Inches(POLICY["typography"]["margin_in"]["x"])
    tf.margin_top = Inches(POLICY["typography"]["margin_in"]["y"])
    tf.margin_bottom = Inches(POLICY["typography"]["margin_in"]["y"])
    tf.vertical_anchor = {"middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}.get(style["valign"], MSO_ANCHOR.TOP)
    # 声明的几何就是设计决定：关闭"自动调整形状以适应文本"，避免目标环境把框缩小/撑大后
    # 让 G-PRE 的墨迹边界检查与实际版面脱节。装不下由 E010/G-PRE 报出来，不靠静默缩放。
    tf.auto_size = MSO_AUTO_SIZE.NONE

    for li, line in enumerate(str(content).split("\n")):
        p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
        p.line_spacing = style["line_spacing"]
        if space_after_pt is not None:
            p.space_after = Pt(space_after_pt)
        if align:
            p.alignment = {
                "left": PP_ALIGN.LEFT,
                "center": PP_ALIGN.CENTER,
                "right": PP_ALIGN.RIGHT,
                "justify": PP_ALIGN.JUSTIFY,
            }[align]
        # 单段内联 runs 的拼接一致性已在校验阶段硬保证；每个片段的式样已由 run_styles 解析完毕
        for spec in (runs if runs and li == 0 else [dict(style, text=line)]):
            run = p.add_run()
            run.text = spec["text"]
            _apply_run_font(run, spec["font"], spec["size"], spec["color"], bold=spec["bold"],
                            letter_spacing_pt=spec["letter_spacing_pt"])


def _add_shape(slide, x: float, y: float, w: float, h: float,
               color: str, shape_kind: str,
               object_name: Optional[str] = None,
               fill_alpha: Optional[float] = None):
    shape_map = {
        "rect": MSO_SHAPE.RECTANGLE,
        "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
    }
    sp = slide.shapes.add_shape(shape_map[shape_kind], Inches(x), Inches(y), Inches(w), Inches(h))
    if object_name:
        sp.name = object_name
    sp.fill.solid()
    sp.fill.fore_color.rgb = _hex_to_rgb(color)
    _apply_fill_alpha(sp, fill_alpha)
    sp.line.fill.background()
    sp.shadow.inherit = False
    style_el = sp._element.find(qn("p:style"))
    if style_el is not None:
        eff = style_el.find(qn("a:effectRef"))
        if eff is not None:
            eff.set("idx", "0")


def _resolve_image_path(image_id: str, asset_root: str) -> Optional[Path]:
    root = Path(asset_root).resolve()
    raw = Path(image_id)
    candidates: List[Path] = []
    if raw.suffix:
        candidates.append((root / raw).resolve())
    for ext in (".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"):
        candidates.append((root / f"{image_id}{ext}").resolve())
    for cand in candidates:
        try:
            cand.relative_to(root)
        except ValueError:
            continue
        if cand.is_file():
            return cand
    return None


def _expected_image_crops(src_w: float, src_h: float, w: float, h: float,
                          focus: Optional[Dict[str, Any]] = None) -> Tuple[float, float, float, float]:
    """等比裁切的唯一计算口径，编译器与交付验证器共用。返回 (left, right, top, bottom) 比例。

    focus 是最终可见窗口中心在源图中的归一化位置，缺省 (0.5, 0.5) 即历史居中裁切；
    窗口在源图边界处钳制。源/目标比例在 0.5% 内视为同幅，不裁切（保持历史行为）。
    """
    default_focus = POLICY["defaults"]["element"]["crop_focus"]
    fx = float(default_focus["x"]) if not focus else float(focus.get("x", default_focus["x"]))
    fy = float(default_focus["y"]) if not focus else float(focus.get("y", default_focus["y"]))
    if src_w <= 0 or src_h <= 0 or w <= 0 or h <= 0:
        return 0.0, 0.0, 0.0, 0.0
    src_ratio = src_w / src_h
    dst_ratio = w / h
    if src_ratio > dst_ratio * 1.005:  # 源更宽：裁左右，保留窗口宽 = dst/src
        keep = dst_ratio / src_ratio
        left = min(max(fx - keep / 2.0, 0.0), 1.0 - keep)
        return left, 1.0 - keep - left, 0.0, 0.0
    if dst_ratio > src_ratio * 1.005:  # 目标更宽：裁上下，保留窗口高 = src/dst
        keep = src_ratio / dst_ratio
        top = min(max(fy - keep / 2.0, 0.0), 1.0 - keep)
        return 0.0, 0.0, top, 1.0 - keep - top
    return 0.0, 0.0, 0.0, 0.0


def _add_image(slide, x: float, y: float, w: float, h: float,
               image_id: Optional[str], asset_root: str,
               object_name: Optional[str] = None,
               crop_focus: Optional[Dict[str, Any]] = None) -> Tuple[bool, str, Dict[str, Any]]:
    """返回 (成功, 错误信息, 执行事实)。执行事实仅供 E012 技术风险披露，不评价图像优劣。"""
    if not image_id:
        return False, "image_id 为空", {}
    img_path = _resolve_image_path(image_id, asset_root)
    if img_path is None:
        return False, "文件未找到或路径越出素材根目录", {}
    try:
        pic = slide.shapes.add_picture(str(img_path), Inches(x), Inches(y), Inches(w), Inches(h))
        if object_name:
            pic.name = object_name
        img_w, img_h = pic.image.size
        info: Dict[str, Any] = {"src_px": [img_w, img_h]}
        if img_w > 0 and img_h > 0 and w > 0 and h > 0:
            cl, cr, ct, cb = _expected_image_crops(img_w, img_h, w, h, crop_focus)
            pic.crop_left, pic.crop_right, pic.crop_top, pic.crop_bottom = cl, cr, ct, cb
            visible_w = img_w * (1.0 - cl - cr)
            visible_h = img_h * (1.0 - ct - cb)
            info["effective_ppi"] = min(visible_w / w, visible_h / h)
            info["crop_loss"] = 1.0 - (visible_w * visible_h) / (img_w * img_h)
        return True, "", info
    except Exception as exc:  # 读取/裁切失败必须保留真实根因
        return False, f"{type(exc).__name__}: {exc}", {}


def _compose_chart_caption(chart: Dict[str, Any]) -> str:
    if chart.get("show_title", POLICY["defaults"]["element"]["show_title"]) is False:
        return ""  # 抑重复图题：单位/期间须由作者在原生说明元素中交代，编译器不替补也不静默丢失声明
    title = str(chart.get("title") or "").strip()
    meta: List[str] = []
    unit, period = _effective_chart_meta(chart)
    if unit and unit not in title:
        meta.append(unit)
    if period and period not in title:
        meta.append(period)
    suffix = " · ".join(meta)
    if title and suffix:
        return f"{title}（{suffix}）"
    return title or suffix


def _normalize_table_data(chart: Dict[str, Any]) -> Tuple[List[str], List[List[str]]]:
    """表格两种合法载荷共用一个规范化函数，编译器与 G-SEM 不得各自解释。"""
    data = chart["data"]
    if "headers" in data:
        headers = [str(c) for c in data["headers"]]
        rows = [[str(c) for c in row] for row in data["rows"]]
        if any(len(row) != len(headers) for row in rows):
            raise ValueError("表格行长度与 headers 不一致")
        return headers, rows

    categories = [str(c) for c in data["categories"]]
    series_list = data["series"]
    caption_unit, _ = _effective_chart_meta(chart)
    # 首列是类别行名：表头留空，编译器不替作者编造"指标/类别"之类的标签。
    headers = [""]
    for series in series_list:
        name, unit = str(series["name"]), _meta_text(series.get("unit"))
        # 系列自带单位且图题未统一标注时写进列头：声明过的单位必须在成品上可见。
        if unit and unit not in name and not (caption_unit and _same_unit(_parse_unit(unit), _parse_unit(caption_unit))):
            name = f"{name}（{unit}）"
        headers.append(name)
    if any(len(s["values"]) != len(categories) for s in series_list):
        raise ValueError("表格系列长度与类别数不一致")
    rows = [
        [cat] + [str(series["values"][ci]) for series in series_list]
        for ci, cat in enumerate(categories)
    ]
    return headers, rows


def _table_caption_height(chart: Dict[str, Any], h: float) -> float:
    cfg = POLICY["chart"]["table_caption"]
    if _compose_chart_caption(chart) and h > cfg["min_element_height_in"]:
        return min(cfg["max_height_in"], h * cfg["height_ratio"])
    return 0.0


def _table_geometry(chart: Dict[str, Any], x: float, y: float, w: float, h: float) -> Tuple[float, float, float, float]:
    """返回原生表格对象的实际几何；标题是独立文本对象，不能冒充表格本体。"""
    cfg = POLICY["chart"]["table_caption"]
    cap_h = _table_caption_height(chart, h)
    if cap_h:
        return x, y + cap_h + cfg["gap_in"], w, max(cfg["min_table_height_in"], h - cap_h - cfg["gap_in"])
    return x, y, w, h


def _add_table(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    chart: Dict[str, Any],
    heading_font: str,
    body_font: str,
    bg_hex: str,
    object_name: Optional[str] = None,
) -> bool:
    headers, rows = _normalize_table_data(chart)  # 形状已由契约保证；异常即实现缺陷，不吞掉
    _dark, style = chart_tone(bg_hex)
    ink_primary = style["title"]
    ink_body = style["body"]
    rule_strong = style["rule_strong"]
    rule_subtle = style["rule_subtle"]

    caption = _compose_chart_caption(chart)
    cap_h = _table_caption_height(chart, h)
    _, table_y, _, table_h = _table_geometry(chart, x, y, w, h)
    if cap_h:
        _add_textbox(slide, (x, y, w, cap_h), caption,
                     plain_text_style(POLICY["chart"]["font_pt"]["title"], ink_primary, heading_font, bold=True),
                     align="left", object_name=f"{object_name}.caption" if object_name else None)

    n_rows = 1 + len(rows)
    n_cols = len(headers)
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(table_y), Inches(w), Inches(table_h))
    if object_name:
        tbl_shape.name = object_name
    tbl = tbl_shape.table
    weights = chart.get("column_weights")  # 长度已在契约层硬校验
    if weights:
        # 逐列取整会让列宽之和偏离框宽（EMU 级漂移）：余数归入末列，总宽精确等于声明几何
        frame, total_weight = int(Inches(w)), float(sum(weights))
        widths = [int(frame * float(weight) / total_weight) for weight in weights]
        widths[-1] += frame - sum(widths)
        for ci, width in enumerate(widths):
            tbl.columns[ci].width = width

    # 清除内置带状/外框样式，回归“场域 > 对齐 > 发丝线 > 字体层级”无竖框三线秩序
    tblPr = tbl._tbl.find(qn("a:tblPr"))
    if tblPr is not None:
        for attr in ("bandRow", "bandCol", "firstRow", "lastRow", "firstCol", "lastCol"):
            tblPr.set(attr, "0")
        style_id = tblPr.find(qn("a:tableStyleId"))
        if style_id is not None:
            style_id.text = POLICY["chart"]["table_style_id"]

    col_aligns = [
        PP_ALIGN.RIGHT
        if ci > 0 and all(bool(_extract_numbers(r[ci])) and len(r[ci].strip()) <= 12 for r in rows)
        else PP_ALIGN.LEFT
        for ci in range(n_cols)
    ]

    for ci, text in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.margin_left = Inches(0.12)
        cell.margin_right = Inches(0.12)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        _set_cell_hairlines(cell, top_hex=rule_strong, top_pt=0.75, bottom_hex=rule_strong, bottom_pt=1.0)
        p = cell.text_frame.paragraphs[0]
        p.alignment = col_aligns[ci]
        run = p.add_run()
        run.text = text
        _apply_run_font(run, heading_font, POLICY["chart"]["font_pt"]["table_header"], ink_primary, bold=True)

    last_ri = len(rows)
    for ri, row in enumerate(rows, start=1):
        is_last = ri == last_ri
        for ci, text in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.margin_left = Inches(0.12)
            cell.margin_right = Inches(0.12)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            _set_cell_hairlines(
                cell,
                bottom_hex=rule_strong if is_last else rule_subtle,
                bottom_pt=0.75 if is_last else 0.5,
            )
            p = cell.text_frame.paragraphs[0]
            p.alignment = col_aligns[ci]
            run = p.add_run()
            run.text = text
            _apply_run_font(run, body_font, POLICY["chart"]["font_pt"]["table_body"], ink_primary if ci == 0 else ink_body, bold=(ci == 0))
    return True


def _contrast_ratio(a_hex: str, b_hex: str) -> float:
    """WCAG 相对亮度对比度（1–21），用于判断系列色在当前页面底色上的可辨识度。"""
    la, lb = _rel_luminance(a_hex), _rel_luminance(b_hex)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


# 项目技术风险阈值，不是 WCAG AA 合规判定；整体无障碍需要独立验证。
PALETTE_MIN_CONTRAST = POLICY["chart"]["min_contrast"]


def _label_ink(fill_hexes: List[str], bg_hex: Optional[str]) -> str:
    """数据标注墨色：选"在最差扇区上也最可读"的一档，而不是固定深色或白色。

    扇区颜色深浅不一（如米棕 + 墨绿 + 深蓝），固定深色字在深蓝扇区上几乎不可读。
    这里取深墨/白两档，比较各自对全部扇区的**最小对比度**，取较大者。
    """
    bg = bg_hex or POLICY["defaults"]["background"]
    # 候选墨色只有 policy 一处定义：深墨、浅墨，再加页面底色本身
    # （深底页时底色即最亮/最暗的可用墨色）。不在这里另写第二套 HEX。
    ink_options = POLICY["chart"]["label_ink"]
    candidates = list(ink_options)
    if bg.upper() not in {c.upper() for c in candidates}:
        candidates.append(bg)
    fills = [c for c in fill_hexes if isinstance(c, str)] or [bg]
    best, best_worst = ink_options[0], -1.0
    for ink in candidates:
        worst = min(_contrast_ratio(ink, f) for f in fills)
        if worst > best_worst:
            best, best_worst = ink, worst
    return best


def _pick_series_colors(palette: List[str], bg_hex: str, kind: str, count: int) -> Tuple[List[str], List[str]]:
    """按位置取色（数量与去重已在契约层校验）；对比不足只披露，绝不重排、复用或回退主题色。"""
    picked = palette[:count]
    floor = PALETTE_MIN_CONTRAST[kind]
    notes = [f"位置 {i + 1} 保留 {color}：对比度 {ratio:.2f}:1 < 技术复核阈值 {floor:.1f}:1；未换色，不代表 WCAG 合规"
             for i, color in enumerate(picked) if (ratio := _contrast_ratio(color, bg_hex)) < floor]
    return picked, notes


DEEMPHASIS_TARGET_MARGIN = 0.05


def _mix_toward(a_hex: str, b_hex: str, t: float) -> str:
    """在 sRGB 字节空间线性插值：t=0 为 a、t=1 为 b。用于最小幅度地推深或提亮。"""
    a = [int(a_hex.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(b_hex.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


def chart_deemphasis_ink(bg_hex: str, kind: str = "bar") -> Tuple[str, Optional[str]]:
    """反强调（上下文）图色：以 policy 意图色为锚，按**实际页底色**兜底到可辨阈值。

    为什么不能只给一个常量：对比度是「两色之间」的量，而底色是逐页声明的。同一枚中性灰
    在 #FFFFFF 上够 2.11:1，在米白 #F2EDE4 上只剩 1.81:1——任何固定 HEX 都不可能同时成立。
    这里只做一件事：保持色相，朝背景的反方向移动**最小**距离，使对比度达到
    min_contrast[kind] + 余量。不动阈值、不改语义、不引入第二套调色板。

    返回 (落色, 披露说明)；未调整时说明为 None。BUILD 与 VERIFY 必须调用同一个函数，
    否则两处各算一遍就会漂移——这正是 E009 与「高亮/退色身份漂移」互相打架的根因。
    """
    floor = float(PALETTE_MIN_CONTRAST.get(kind, PALETTE_MIN_CONTRAST.get("bar", 2.0)))
    target = floor + DEEMPHASIS_TARGET_MARGIN
    dark, _style = chart_tone(bg_hex)
    anchor = POLICY["chart"]["deemphasis"]["dark" if dark else "light"]
    if _contrast_ratio(anchor, bg_hex) >= target:
        return anchor, None

    def _first_clear(end_hex: str) -> Optional[float]:
        """朝该端点插值，达到 target 所需的最小 t；端点本身都不够则返回 None。"""
        if _contrast_ratio(_mix_toward(anchor, end_hex, 1.0), bg_hex) < target:
            return None
        low, high = 0.0, 1.0
        for _ in range(18):
            mid = (low + high) / 2.0
            if _contrast_ratio(_mix_toward(anchor, end_hex, mid), bg_hex) >= target:
                high = mid
            else:
                low = mid
        return high

    # 黑白两端都试，取移动更小的那端：底色落在中间明度时，靠明暗直觉选方向会选错。
    options = [(t, end) for end in ("#000000", "#FFFFFF") if (t := _first_clear(end)) is not None]
    if not options:
        return anchor, (f"反强调色 {anchor} 在底色 {bg_hex} 上够不到 {floor:.1f}:1，"
                        "黑白两端都不足以拉开差距；请更换页面底色或显式声明系列色")
    high, end = min(options)
    best = _mix_toward(anchor, end, high)
    # 量化到字节后可能刚好掉回阈值下方，按 1/255 步进补足（上限 8 步）。
    for _ in range(8):
        if _contrast_ratio(best, bg_hex) >= target:
            break
        high = min(1.0, high + 1.0 / 255.0)
        best = _mix_toward(anchor, end, high)
    ratio = _contrast_ratio(best, bg_hex)
    note = (f"反强调色按底色兜底：意图色 {anchor} → {best}（对底色 {bg_hex} 为 {ratio:.2f}:1，"
            f"阈值 {floor:.1f}:1）；只为可辨，未改变强调/退隐的语义")
    return best, note


def _add_chart(
    slide,
    box: Tuple[float, float, float, float],
    chart: Dict[str, Any],
    heading_font: str,
    body_font: str,
    palette: List[str],
    bg_hex: str,
    notes: List[str],
    info_notes: Optional[List[str]] = None,
    object_name: Optional[str] = None,
) -> bool:
    """绘制已通过契约校验的原生图表/表格。python-pptx 的真实失败保留根因并返回 False。"""
    x, y, w, h = box
    kind = chart["kind"]
    if kind == "table":
        return _add_table(slide, x, y, w, h, chart, heading_font, body_font, bg_hex, object_name)
    data = chart["data"]
    series_list = data["series"]
    dark, style = chart_tone(bg_hex)
    ink_title = style["title"]
    ink_tick = style["tick"]
    rule_grid = style["grid"]

    try:
        chart_data = CategoryChartData()
        chart_data.categories = list(data["categories"])
        for s in series_list:
            chart_data.add_series(str(s["name"]), [float(v) for v in s["values"]])
        xl_type = {
            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE,
            "donut": XL_CHART_TYPE.DOUGHNUT,
        }[kind]
        gframe = slide.shapes.add_chart(xl_type, Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
        if object_name:
            gframe.name = object_name
        chart_obj = gframe.chart

        # 去除图表区底色与外框，融入页面连续场域（python-pptx 默认的绘图区本无填充）。
        # CT_ChartSpace 要求 c:spPr 紧随 c:chart；追加到末尾会落在 txPr/externalData 之后而违反 OOXML 顺序。
        cs = chart_obj._chartSpace
        for old in cs.findall(qn("c:spPr")):
            cs.remove(old)
        cs_spPr = cs.makeelement(qn("c:spPr"), {})
        cs_spPr.append(cs_spPr.makeelement(qn("a:noFill"), {}))
        cs_ln = cs_spPr.makeelement(qn("a:ln"), {})
        cs_ln.append(cs_ln.makeelement(qn("a:noFill"), {}))
        cs_spPr.append(cs_ln)
        cs.find(qn("c:chart")).addnext(cs_spPr)

        caption = _compose_chart_caption(chart)
        chart_obj.has_title = bool(caption)
        if caption:
            tf = chart_obj.chart_title.text_frame
            tf.clear()
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT if POLICY["chart"].get("title_align", "left") == "left" else PP_ALIGN.CENTER
            run = para.add_run()
            run.text = caption
            _apply_run_font(run, heading_font, POLICY["chart"]["font_pt"]["title"], ink_title, bold=True)
            if POLICY["chart"].get("title_align", "left") == "left":
                # 自动布局会把图题居中在整个图表区；显式 manualLayout 把它钉在绘图区左缘，
                # 与页面文字轴对齐。目标 PowerPoint 上若被忽略，退化回默认居中，不会更差（G-TGT 未验证）。
                title_el = chart_obj._chartSpace.find(qn("c:chart")).find(qn("c:title"))
                if title_el is not None:
                    for old_layout in title_el.findall(qn("c:layout")):
                        title_el.remove(old_layout)
                    layout = title_el.makeelement(qn("c:layout"), {})
                    manual = layout.makeelement(qn("c:manualLayout"), {})
                    for tag, val in (("c:layoutTarget", "inner"), ("c:xMode", "edge"), ("c:yMode", "edge"),
                                     ("c:x", "0"), ("c:y", "0"), ("c:w", "1"), ("c:h", "0.12")):
                        manual.append(manual.makeelement(qn(tag), {"val": val}))
                    layout.append(manual)
                    # c:layout 必须紧随 c:tx
                    tx = title_el.find(qn("c:tx"))
                    if tx is not None:
                        tx.addnext(layout)
                    else:
                        title_el.insert(0, layout)

        # 图例只在多系列或饼/环（类别靠颜色区分）时保留，单系列柱线图不叠加重复标签
        want_legend = len(series_list) > 1 or kind in {"pie", "donut"}
        chart_obj.has_legend = want_legend
        if want_legend:
            chart_obj.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart_obj.legend.include_in_layout = False
            chart_obj.legend.font.name = body_font
            chart_obj.legend.font.size = Pt(POLICY["chart"]["font_pt"]["legend"])
            chart_obj.legend.font.color.rgb = _hex_to_rgb(ink_tick)
            _ensure_ea_in_txpr(chart_obj.legend._element, body_font)

        hide_value_axis = (chart.get("axis_labels", POLICY["defaults"]["element"]["axis_labels"]) is False
                           and kind in {"bar", "line"})
        for axis_attr in (("category_axis", "value_axis") if kind in {"bar", "line"} else ()):
            axis = getattr(chart_obj, axis_attr)
            axis.major_tick_mark = XL_TICK_MARK.NONE  # 刻度短线不承载信息：读数靠网格或直标
            axis.minor_tick_mark = XL_TICK_MARK.NONE
            if hide_value_axis and axis_attr == "value_axis":
                # 数值已逐点直标（契约硬保证）：整根数值轴退场——只去标签而留轴线与刻度是噪声
                axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE
                axis.has_major_gridlines = False
                axis.visible = False
                continue
            axis.tick_labels.font.name = body_font
            axis.tick_labels.font.size = Pt(POLICY["chart"]["font_pt"]["axis"])
            axis.tick_labels.font.color.rgb = _hex_to_rgb(ink_tick)
            _ensure_ea_in_txpr(axis._element, body_font)
            axis.format.line.color.rgb = _hex_to_rgb(rule_grid)
            axis.format.line.width = Pt(0.75)
            if axis_attr == "value_axis":
                axis.has_major_gridlines = True
                axis.major_gridlines.format.line.color.rgb = _hex_to_rgb(rule_grid)
                axis.major_gridlines.format.line.width = Pt(0.5)

        # 调色板身份稳定；低对比仅披露，不由 BUILD 重排语义色
        eligible, palette_notes = _pick_series_colors(palette, bg_hex, kind, chart_color_count(chart))
        notes.extend(palette_notes)
        highlight = chart.get("highlight_category")
        if kind == "bar":
            chart_obj.plots[0].gap_width = int(POLICY["chart"]["bar_gap_width"])
        if eligible:
            if kind == "bar" and highlight is not None and len(series_list) == 1:
                # 编辑型强调：只有承载主张的那根柱保留首色，其余退至中性上下文色。
                deemphasis, deemphasis_note = chart_deemphasis_ink(bg_hex, kind)
                if deemphasis_note:
                    (info_notes if info_notes is not None else notes).append(deemphasis_note)
                for pi_, point in enumerate(chart_obj.plots[0].series[0].points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = _hex_to_rgb(eligible[0] if pi_ == highlight else deemphasis)
            elif kind == "bar":
                for si, ser in enumerate(chart_obj.series):
                    ser.format.fill.solid()
                    ser.format.fill.fore_color.rgb = _hex_to_rgb(eligible[si])
            elif kind == "line":
                for si, ser in enumerate(chart_obj.series):
                    hex_c = _hex_to_rgb(eligible[si])
                    ser.format.line.color.rgb = hex_c
                    ser.marker.format.fill.solid()
                    ser.marker.format.fill.fore_color.rgb = hex_c
                    ser.marker.format.line.color.rgb = hex_c
            elif kind in {"pie", "donut"}:
                # 类别由颜色区分，必须逐扇区着色；同时给出数值标注，避免"只靠颜色编码"
                points = chart_obj.plots[0].series[0].points
                for pi_, point in enumerate(points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = _hex_to_rgb(eligible[pi_])
        if kind in {"pie", "donut"}:
            # 饼/环的百分比标注由图型决定（Schema 不允许 data_labels 声明）
            plot = chart_obj.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_percentage = True
            dl.show_value = False
            dl.show_category_name = False
            dl.show_series_name = False
            dl.show_legend_key = False
            dl.number_format = "0%"
            dl.number_format_is_linked = False
            dl.font.size = Pt(POLICY["chart"]["font_pt"]["data_label"])
            dl.font.name = body_font
            dl.font.color.rgb = _hex_to_rgb(_label_ink(eligible or [], bg_hex))
            _ensure_ea_in_txpr(dl._element, body_font)
        elif kind in {"bar", "line"} and chart.get("data_labels"):
            plot = chart_obj.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_value = True
            dl.show_category_name = False
            dl.show_series_name = False
            dl.show_legend_key = False
            dl.font.size = Pt(POLICY["chart"]["font_pt"]["data_label"])
            dl.font.name = body_font
            dl.font.color.rgb = _hex_to_rgb(ink_title)
            dl.number_format = POLICY["chart"]["data_label_format"]
            dl.number_format_is_linked = False
            _ensure_ea_in_txpr(dl._element, body_font)
            if kind == "bar":
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
        return True
    except Exception as exc:  # python-pptx 的真实失败：保留根因，由调用方记 E007 并阻断
        notes.append(f"图表渲染异常：{type(exc).__name__}: {exc}")
        return False


def _same_resolved(a: Path, b: Path) -> bool:
    try:
        return a.resolve() == b.resolve()
    except OSError:
        return False


def _output_collision_errors(out_path: str, plan_path: Optional[str], plan: Dict[str, Any],
                             asset_root: str) -> List[CompileError]:
    """拒绝把输出写到输入 plan 或任何已引用素材上，避免一次误参数覆盖源文件。"""
    errors: List[CompileError] = []
    out = Path(out_path)
    if plan_path and _same_resolved(out, Path(plan_path)):
        errors.append(CompileError(
            "E001", ErrorLevel.HARD, "deck.output", "输出 PPTX 不得覆盖输入 plan",
            f"output={out_path!r} 与 plan={plan_path!r} 解析为同一路径",
            "换一个输出文件名；输入与交付物必须物理分离",
        ))
    used = {el["image_id"] for page in plan.get("pages", []) for el in page.get("elements", [])
            if isinstance(el, dict) and el.get("type") == "image"}
    for image_id in sorted(used):
        image_path = _resolve_image_path(image_id, asset_root)
        if image_path is not None and _same_resolved(out, image_path):
            errors.append(CompileError(
                "E001", ErrorLevel.HARD, "deck.output", "输出 PPTX 不得覆盖已引用素材",
                f"output={out_path!r} 与素材 {image_id!r} 解析为同一路径",
                "换一个输出文件名；素材是证据输入，不能被生成物覆盖",
            ))
    return errors


def _atomic_save_presentation(prs, path: Path) -> None:
    """同目录临时文件 + 原子替换：失败/中断不会截断上一份合格交付物。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=path.suffix, dir=str(path.parent))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        prs.save(str(temp_path))
        os.replace(temp_path, path)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass


def compile_plan(
    plan: Dict[str, Any],
    out_path: str,
    asset_root: str = ".",
    strict: bool = False,
    plan_path: Optional[str] = None,
) -> Tuple[bool, List[CompileError]]:
    """返回 (是否成功, 错误列表)。

    HARD 在写文件前返回；strict=True 时未处置的 SOFT 也判失败。失败时**不覆盖**目标
    out_path，而是把带诊断占位的 PPTX 写到 `<stem>.error.pptx`，保留上一份合格交付物。
    """
    collisions = _output_collision_errors(out_path, plan_path, plan, asset_root)
    if collisions:
        return False, collisions
    errors = validate_plan(plan, asset_root)
    if any(e.level == ErrorLevel.HARD for e in errors):
        return False, errors
    slide_w_in, slide_h_in = _slide_dimensions(plan.get("aspect_ratio"))

    fonts = plan.get("fonts", {})
    heading_font = fonts.get("heading", DEFAULT_FONTS["heading"])
    body_font = fonts.get("body", DEFAULT_FONTS["body"])
    uses_default_palette = "palette" not in plan
    palette = resolved_palette(plan)
    placeholder = POLICY["chart"]["placeholder"]

    def _diagnostic_placeholder(slide, box, object_name: str, message: str) -> None:
        """运行期失败只画诊断占位并阻断交付；它不是可交付内容。"""
        x, y, w, h = box
        _add_shape(slide, x, y, w, h, placeholder["fill"], "rect", object_name=f"{object_name}.error")
        _add_textbox(slide, (x + 0.05, y + 0.05, max(0.1, w - 0.1), max(0.1, h - 0.1)), message,
                     plain_text_style(min_pt_for("caption"), placeholder["text"], body_font),
                     align="left", object_name=f"{object_name}.error.message")

    prs = Presentation()
    prs.slide_width = Inches(slide_w_in)
    prs.slide_height = Inches(slide_h_in)
    blank_layout = prs.slide_layouts[6]

    for page_index, page in enumerate(plan["pages"]):
        slide = prs.slides.add_slide(blank_layout)
        bg_hex = page_background(page)
        # 总是显式写入底色：交付核对与对比度计算基于文件里真实存在的值，而不是模板继承的假设。
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = _hex_to_rgb(bg_hex)

        pid = page["id"]
        # Python sort 稳定；z_index 相同时保持 THINK 交接的元素顺序。
        default_z = float(POLICY["defaults"]["element"]["z_index"])
        for el in sorted(page["elements"], key=lambda e: float(e.get("z_index", default_z))):
            etype = el["type"]
            eid = el["id"]
            object_name = f"{pid}.{eid}"
            box = tuple(float(el[k]) for k in ("x", "y", "w", "h"))
            x, y, w, h = box

            if etype in TEXT_ELEMENT_TYPES:
                _add_textbox(
                    slide, box, el["content"], text_style(el, fonts),
                    align=el.get("align"),
                    runs=run_styles(el, fonts) if el.get("runs") else None,
                    space_after_pt=el.get("space_after_pt"),
                    object_name=object_name,
                )
            elif etype == "image":
                ok, image_error, image_info = _add_image(
                    slide, x, y, w, h, el["image_id"], asset_root,
                    object_name=object_name,
                    crop_focus=el.get("crop_focus"),
                )
                if ok:
                    image_cfg = POLICY["image"]
                    ppi = image_info.get("effective_ppi")
                    if ppi is not None and ppi < image_cfg["min_effective_ppi"]:
                        errors.append(CompileError(
                            "E012", ErrorLevel.SOFT, f"page[{pid}].element[{eid}].image",
                            f"有效分辨率 ≥ {image_cfg['min_effective_ppi']} ppi（技术下限）",
                            f"实际约 {ppi:.0f} ppi 于 {w:.2f}×{h:.2f} in 落位",
                            "投影下将明显发虚：换更高分辨率源图、缩小落位或按最终画幅重出；不评价构图优劣",
                        ))
                    loss = image_info.get("crop_loss")
                    if loss is not None and loss > image_cfg["max_crop_loss"]:
                        errors.append(CompileError(
                            "E012", ErrorLevel.SOFT, f"page[{pid}].element[{eid}].image",
                            f"裁切损失 ≤ {image_cfg['max_crop_loss']:.0%}（技术下限）",
                            f"实际丢失源图约 {loss:.0%} 面积",
                            "主体大概率被切伤：用 crop_focus 保主体，或直接按最终画幅比例生成/选材；有意取局部细节可具名复核",
                        ))
                else:
                    _diagnostic_placeholder(slide, box, object_name, f"[缺失资产：{el['image_id']}]")
                    errors.append(CompileError(
                        "E003", ErrorLevel.HARD, f"page[{pid}].element[{eid}].image",
                        f"资产 {el['image_id']} 可读且能置入 PPTX", image_error,
                        "图片资源不可用，已绘制仅供诊断的占位框；交付仍阻断",
                    ))
            elif etype == "chart":
                chart_notes: List[str] = []
                chart_info: List[str] = []
                ok = _add_chart(slide, box, el["chart"], heading_font, body_font, palette, bg_hex,
                                chart_notes, info_notes=chart_info, object_name=object_name)
                if ok:
                    if uses_default_palette and el["chart"]["kind"] != "table":
                        errors.append(CompileError(
                            "E009", ErrorLevel.INFO, f"page[{pid}].element[{eid}].chart",
                            "缺省色来自唯一技术策略", "未提供 palette；使用 policy.json 默认图表色",
                            "仅披露执行默认，不证明品牌/色彩语义适配；仍须整稿复核",
                        ))
                    for note in chart_info:
                        errors.append(CompileError(
                            "E009", ErrorLevel.INFO, f"page[{pid}].element[{eid}].chart",
                            "反强调色须在页面底色上可辨", note,
                            "系统按 min_contrast 兜底保证可辨；若需完全自定义请显式声明系列色",
                        ))
                    for note in chart_notes:
                        errors.append(CompileError(
                            "E009", ErrorLevel.SOFT, f"page[{pid}].element[{eid}].chart",
                            "图表配色与标注须在页面底色上可辨且语义一致", note,
                            "图表已渲染但存在配色/标注风险，请人工复核渲染结果",
                        ))
                else:
                    _diagnostic_placeholder(slide, box, object_name, "[图表渲染失败]")
                    errors.append(CompileError(
                        "E007", ErrorLevel.HARD, f"page[{pid}].element[{eid}].chart",
                        f"图表 {eid} 可由 python-pptx 原生绘制",
                        f"kind={el['chart']['kind']!r} 渲染失败" + (f"；原因：{chart_notes[-1]}" if chart_notes else ""),
                        "图表/表格渲染失败，已绘制仅供诊断的占位框；交付仍阻断",
                    ))
            elif etype in {"shape", "divider"}:
                _add_shape(slide, x, y, w, h, el["color"],
                           el.get("shape_kind", POLICY["defaults"]["element"]["shape_kind"]),
                           object_name=object_name,
                           fill_alpha=el.get("fill_alpha"))
            else:  # Schema 已排除；若权威源与实现漂移，必须显式失败
                raise RuntimeError(f"Schema 允许但编译器未实现的元素类型：{etype!r}")

        # 页码是整稿契约派生的功能性对象：最后绘制（最上层），
        # 不因风格或布局紧张而被省略。编号与 pages 顺序同源。
        pn = page_number_element(plan, page_index)
        if pn is not None:
            _add_textbox(
                slide, (pn["x"], pn["y"], pn["w"], pn["h"]), pn["content"],
                text_style(pn, fonts), align=pn["align"],
                object_name=f"{pid}.{PAGE_NUMBER_ID}",
            )

    success = not any(e.level == ErrorLevel.HARD for e in errors)
    if strict and unresolved_reviews(plan, errors):
        success = False
    out_file = Path(out_path)
    # 失败产物另存 .error.pptx：既保留诊断现场，也不覆盖上一份已通过的文件。
    target = out_file if success else out_file.with_name(f"{out_file.stem}.error{out_file.suffix}")
    try:
        _atomic_save_presentation(prs, target)
    except OSError as exc:
        return False, errors + [CompileError(
            "E001", ErrorLevel.HARD, "deck.output", "PPTX 可原子写入目标路径",
            f"{target}: {type(exc).__name__}: {exc}",
            "检查目录权限、磁盘空间与目标文件是否被占用；不得留下截断文件",
        )]
    return success, errors


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="编译 slide plan 为 PPTX 并做硬错误检查")
    parser.add_argument("plan", help="slide plan JSON 路径")
    parser.add_argument("output", help="输出 PPTX 路径")
    parser.add_argument("--assets", default=".", help="素材根目录（默认当前目录）")
    parser.add_argument("--strict", action="store_true",
                        help="严格模式：未完成具名复核的 SOFT 问题视为失败")
    args = parser.parse_args(argv)

    # 记录既有诊断文件的状态：本次失败只能指向本次写出的诊断，不得把上一次遗留的旧文件说成本次产物。
    _target = Path(args.output)
    _diag_path = _target.with_name(f"{_target.stem}.error{_target.suffix}")
    try:
        _diag_before = _diag_path.stat().st_mtime_ns if _diag_path.is_file() else None
    except OSError:
        _diag_before = None

    plan, load_errors = load_plan(args.plan)
    errors: List[CompileError] = list(load_errors)
    if plan is None:
        for e in errors:
            print(str(e), file=sys.stderr)
        return 1

    success, comp_errors = compile_plan(plan, args.output, asset_root=args.assets,
                                       strict=args.strict, plan_path=args.plan)
    errors.extend(comp_errors)

    if errors:
        print(f"\n共发现 {len(errors)} 个问题：", file=sys.stderr)
        for e in errors:
            print(str(e), file=sys.stderr)

    if not success:
        diagnostic = _diag_path
        try:
            _diag_after = diagnostic.stat().st_mtime_ns if diagnostic.is_file() else None
        except OSError:
            _diag_after = None
        if _diag_after is not None and _diag_after != _diag_before:
            detail = f"诊断文件仅供排查，不得作为交付物：{diagnostic}"
        elif _diag_after is not None:
            detail = (f"失败发生在创建文件之前，本次未写出 PPTX；正式输出未被覆盖。"
                      f"目录中的 {diagnostic} 是更早一次失败遗留的旧诊断文件，同样不得交付。")
        else:
            detail = "失败发生在创建文件之前，未写出 PPTX；正式输出未被覆盖。"
        print(f"\nBLOCK — Delivery：存在未解决的强制条件。{detail}", file=sys.stderr)
        return 1
    print(f"编译成功：{args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
