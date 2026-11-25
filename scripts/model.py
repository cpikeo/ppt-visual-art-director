#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""plan 模型与共用工具：诊断、载入、度量、颜色、几何、成品包校验。

compile / verify / revise 共用这一份实现；这里不依赖 python-pptx，可单独校验 plan。
"""

from __future__ import annotations

import collections
import hashlib
import json
import math
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree


HERE = Path(__file__).resolve().parent

LIMITS: Dict[str, Any] = json.loads((HERE / "limits.json").read_text(encoding="utf-8-sig"))

SCHEMA: Dict[str, Any] = json.loads((HERE / "plan.schema.json").read_text(encoding="utf-8-sig"))

TEXT_TYPES = ("title", "body", "text", "number", "caption", "footer")

DASH = {"solid": "solid", "dash": "dash", "dot": "sysDot", "dash_dot": "dashDot", "long_dash": "lgDash"}

EMU = 914400.0

PT_EMU = 12700.0

Box = Tuple[float, float, float, float]

# OOXML 命名空间。自己拼簇名而不用 pptx.oxml.ns.qn：那会把 pptx 提升为导入期依赖，
# 而本模块要先能在"只校验 plan"的场景下单独运行。
_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return "{%s}%s" % (_NS[prefix], local)

# 解析来自外部的 PPTX 部件一律走这个解析器：不展开实体、不联网，防 XXE 与实体膨胀。
SAFE_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False, huge_tree=False)

def parse_xml(data: bytes):
    return etree.fromstring(data, SAFE_PARSER)

# ---------------------------------------------------------------- 诊断


class BadOp(Exception):
    """revise.py 的操作无法执行：定位不到对象、参数不合契约。不猜，直接停。"""

@dataclass
class Issue:
    level: str          # HARD | SOFT | INFO
    code: str
    where: str
    message: str

    def __str__(self) -> str:
        return f"[{self.level}] {self.code} {self.where}: {self.message}"

class Report:
    def __init__(self) -> None:
        self.items: List[Issue] = []

    def add(self, level: str, code: str, where: str, message: str) -> None:
        self.items.append(Issue(level, code, where, message))

    def hard(self, code: str, where: str, message: str) -> None:
        self.add("HARD", code, where, message)

    def soft(self, code: str, where: str, message: str) -> None:
        self.add("SOFT", code, where, message)

    def info(self, code: str, where: str, message: str) -> None:
        self.add("INFO", code, where, message)

    def extend(self, other: "Report") -> None:
        self.items.extend(other.items)

    @property
    def blocked(self) -> bool:
        return any(i.level == "HARD" for i in self.items)

    def count(self, level: str) -> int:
        return sum(1 for i in self.items if i.level == level)

    def to_dict(self) -> List[Dict[str, str]]:
        return [{"level": i.level, "code": i.code, "where": i.where, "message": i.message} for i in self.items]

def _common_tail(msgs: List[str]) -> str:
    """多条文案的共同后缀，截到子句边界（：；，。），太短（<24 字）视为没有。"""
    tail = os.path.commonprefix([m[::-1] for m in msgs])[::-1]
    for i, ch in enumerate(tail):
        if ch in "：；，。":
            tail = tail[i:]
            break
    else:
        return ""
    return tail if len(tail) >= 24 else ""


def show(items: List[Issue]) -> None:
    """打印诊断，减少常驻输出与重读成本（JSON 报告 Report.to_dict 仍逐条保留）：
    ① 同级别、同代码、同文案、仅位置不同 → 合并一行；
    ② 同代码、细节不同但共有一段样板说明（如 7 条 W-PUNCT）→ 逐条只留“位置：细节”，样板只印一次。"""
    by_code: Dict[Any, List[Issue]] = {}
    for i in items:
        by_code.setdefault((i.level, i.code), []).append(i)
    for (level, code), g in by_code.items():
        uniq = {i.message for i in g}
        tail = _common_tail(sorted(uniq)) if len(uniq) >= 3 else ""
        if tail:
            print(f"[{level}] {code}（×{len(g)}）{tail[1:] if tail[0] == '：' else tail}")
            for i in g:
                print(f"  {i.where}: {i.message[:-len(tail)]}")
            continue
        groups: Dict[str, List[Issue]] = {}
        for i in g:
            groups.setdefault(i.message, []).append(i)
        for message, gg in groups.items():
            if len(gg) == 1:
                print(gg[0])
            else:
                print(f"[{level}] {code} {'、'.join(i.where for i in gg)}（×{len(gg)}）: {message}")

# ---------------------------------------------------------------- 载入


def _reject_constant(name: str) -> Any:
    raise ValueError(f"不允许 {name}（JSON 标准没有 NaN / Infinity；它们会让几何与图表静默出错）")


def _reject_duplicate_keys(pairs: List[Tuple[str, Any]]) -> Dict[str, Any]:
    seen: Dict[str, Any] = {}
    for key, value in pairs:
        if key in seen:
            raise ValueError(f"对象里键 {key!r} 重复出现（标准解析器会静默取最后一个，改动会无声失效）")
        seen[key] = value
    return seen


def strict_json_loads(raw: str) -> Any:
    """plan / edit / outline 共用的严格 JSON 读取：拒绝 NaN/Infinity、重复键，深度炸弹按 ValueError 报告。"""
    try:
        return json.loads(raw, parse_constant=_reject_constant, object_pairs_hook=_reject_duplicate_keys)
    except RecursionError:
        raise ValueError("JSON 嵌套过深，拒绝解析")


def load_plan(path: Path, rep: Report) -> Optional[Dict[str, Any]]:
    try:
        raw = path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        rep.hard("E-SCHEMA", str(path), f"无法读取：{exc}")
        return None
    try:
        plan = strict_json_loads(raw)
    except ValueError as exc:   # JSONDecodeError 是 ValueError 的子类
        rep.hard("E-SCHEMA", str(path), f"JSON 无效：{exc}")
        return None
    try:
        from jsonschema import Draft202012Validator
    except ImportError:
        rep.hard("E-SCHEMA", "环境", "缺少 jsonschema：pip install -r requirements.txt")
        return plan
    validator = Draft202012Validator(SCHEMA)
    errors = sorted(validator.iter_errors(plan), key=lambda e: list(map(str, e.absolute_path)))
    for err in errors:
        where = "/".join(str(p) for p in err.absolute_path) or "(根)"
        rep.hard("E-SCHEMA", where, err.message[:300])
    if isinstance(plan, dict) and not errors:
        for where in _xml_illegal(plan):
            rep.hard("E-SCHEMA", where, "含 XML 不允许的控制字符（\\x00–\\x08、\\x0b、\\x0c、\\x0e–\\x1f 等）：删掉它，换行用 \\n")
    return plan if isinstance(plan, dict) else None


_XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


def _xml_illegal(node: Any, path: str = "") -> List[str]:
    """找出含 XML 1.0 非法字符的字符串路径（最多 5 条）。在载入期拦下，好过写包末段 E-RUNTIME 或写出 _x0001_ 乱码。"""
    if isinstance(node, str):
        return [path or "(根)"] if _XML_BAD.search(node) else []
    items = node.items() if isinstance(node, dict) else enumerate(node) if isinstance(node, list) else ()
    out: List[str] = []
    for k, v in items:
        out += _xml_illegal(v, f"{path}/{k}" if path else str(k))
        if len(out) >= 5:
            break
    return out[:5]

# ---------------------------------------------------------------- 度量模型


def canvas(plan: Dict[str, Any]) -> Tuple[float, float]:
    return tuple(LIMITS["canvas"][plan.get("aspect_ratio") or LIMITS["canvas"]["default"]])  # type: ignore

def _em_width(ch: str) -> float:
    if ch in " \u00a0\t":
        return 0.30
    code = ord(ch)
    if code < 128:
        if ch.isdigit():
            return 0.58          # 粗体数字偏宽；按 0.55 估，100pt 的“41%”会被低估到刚好放下
        if ch in "%@&":
            return 0.90          # 百分号、@、& 接近一个字宽；按标点的 0.34 估会漏掉大字号数字的折行
        return 0.55 if ch.isalnum() else 0.34
    if ch in "×÷±":
        return 1.00          # 实测：衬线 × 约 1.0 em；按 0.9 估，“2.3×”在 138pt 时被低估而折行
    if ch in "，。、；：？！（）【】《》「」『』“”‘’—…·～":
        return 1.0
    if code >= 0x2E80:
        return 1.0
    return 0.6

def _is_cjk(ch: str) -> bool:
    return ord(ch) >= 0x2E80 and ch not in "，。、；：？！（）【】《》「」『』“”‘’—…·～"


def _boundary_em(text: str) -> float:
    """中文与数字/西文相接处，渲染器（LibreOffice 必然、PowerPoint 视设置）会自动加约 0.1 em 的间距。
    估宽不计它，“估 1 行、实 2 行”会成批出现（渲染实测：「84%，只比一店低4个点」多出一行）。"""
    n = 0
    for a, b in zip(text, text[1:]):
        if (_is_cjk(a) and b.isascii() and (b.isalnum() or b == "%")) or \
           (_is_cjk(b) and a.isascii() and (a.isalnum() or a == "%")):
            n += 1
    return 0.1 * n


def _para_width_em(seg: str, pieces: List[Dict[str, Any]], pt: float) -> float:
    """一段的宽度，以"基准字号 1 em"计。逐 run 按各自字号折算：
    标题里放一个 96pt 数字加 24pt 单位，按整段取同一个字号估高会差出好几行。"""
    if not pieces:
        return sum(_em_width(c) for c in seg) + _boundary_em(seg)
    total = 0.0
    for piece in pieces:
        size = float(piece.get("font_size") or pt)
        text = str(piece.get("text", ""))
        total += (sum(_em_width(c) for c in text) + _boundary_em(text)) * size / pt
    return total

LIST_INDENT_IN = 0.32  # 原生列表的悬挂缩进（英寸）；估宽与 compile 共用这一个数


def list_indent(el: Dict[str, Any]) -> float:
    return LIST_INDENT_IN if el.get("list") else 0.0


def fit_lines(el: Dict[str, Any], st: Dict[str, Any], box_w: float) -> List[int]:
    """逐段行数。按每段最大字号与逐 run 折算的宽度算，不用元素级字号一刀切。"""
    inset = st["pad"][0]
    usable = max(box_w - 2 * inset - list_indent(el), 0.05)
    paras = paragraph_runs(el)
    if paras is None:
        paras = [(str(el.get("content", "")), [])]
    out: List[int] = []
    for seg, pieces in paras:
        base = max([float(p.get("font_size") or st["size"]) for p in pieces] or [st["size"]])
        cap = usable / (base / 72.0)
        out.append(max(1, math.ceil(_para_width_em(seg, pieces, base) / cap - 1e-9)))
    return out or [1]

def line_pitch_em(family: Optional[str]) -> float:
    """单倍行距的行高（em）：按字体族名子串匹配 limits.json，未知字体取默认。"""
    cfg = LIMITS["type"]["line_pitch_em"]
    key = str(family or "").lower().replace(" ", "")
    for name, value in cfg["by_font"].items():
        if name in key:
            return float(value)
    return float(cfg["default"])

def fit_height_in(el: Dict[str, Any], st: Dict[str, Any], box_w: float) -> float:
    """框高需求（英寸）。估不准时以真实渲染为准。

    按 LibreOffice 渲染实测校准：
    首行只有字形本身占可见高度，行盒里超出字形的部分是字体的 ascent/descent 余量，
    不会把墨迹推出框外；只有首行之后的每一行才按整条行盒推进。旧模型把首行也按
    pitch 计入，等于多算一整行的行距余量——11.0.0 在该样本上 5 条 W-FIT 告警里
    没有一条在渲染中形成可见溢出（最大 +3.1pt），而真正越框的 channels/note3（横向
    +11pt）它又不看宽度、一次都没响。现在：首行 base×行距，其余行 base×pitch×行距，
    再加上下内边距；纵向告警回到"框确实装不下可见文字"的语义上。
    """
    pitch = line_pitch_em(st.get("family"))
    ls = float(st["line_spacing"])
    paras = paragraph_runs(el) or [(str(el.get("content", "")), [])]
    lines = fit_lines(el, st, box_w)
    total = 2.0 * float(st["pad"][1])
    # 只有整个文本框的第一行按字形高度计；其后每一行（含后续段落的首行）都按整条行盒推进。
    # 旧版把每一段的首行都当"第一行"，\n 分段的文字（标题两行、要点列表）被系统性低估。
    first = True
    for (seg, pieces), count in zip(paras, lines):
        base = max([float(p.get("font_size") or st["size"]) for p in pieces] or [st["size"]])
        step = base * pitch * ls / 72.0
        if first:
            total += base * ls / 72.0 + max(count - 1, 0) * step
            first = False
        else:
            total += count * step
    space_after = float(el.get("space_after_pt") or 0)
    if space_after:
        total += space_after * max(len(paras) - 1, 1) / 72.0
    return total

def hex_rgb(hex_color: str) -> Tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

def rel_luminance(hex_color: str) -> float:
    def channel(v: int) -> float:
        s = v / 255.0
        return s / 12.92 if s <= 0.03928 else ((s + 0.055) / 1.055) ** 2.4
    r, g, b = hex_rgb(hex_color)
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)

def contrast_ratio(a: str, b: str) -> float:
    la, lb = rel_luminance(a), rel_luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)

def mix_toward(fg: str, bg: str, t: float) -> str:
    """把 fg 向 bg 混合 t（0=原色，1=bg），只用于图表退隐色。"""
    fa, ba = hex_rgb(fg), hex_rgb(bg)
    out = [round(a + (b - a) * t) for a, b in zip(fa, ba)]
    return "#" + "".join(f"{v:02X}" for v in out)

def is_dark(hex_color: str) -> bool:
    return rel_luminance(hex_color) < LIMITS["contrast"]["dark_background_luminance"]

def page_bg(page: Dict[str, Any]) -> str:
    return (page.get("background") or LIMITS["defaults"]["background"]).upper()

def deck_palette(plan: Dict[str, Any]) -> List[str]:
    return [c.upper() for c in (plan.get("palette") or LIMITS["defaults"]["palette"])]

def chart_palette(plan: Dict[str, Any], chart: Dict[str, Any]) -> List[str]:
    """图表取色序列：图自带 palette 优先，否则用整稿 palette。compile、checks、verify 共用这一处。"""
    return [c.upper() for c in (chart.get("palette") or [])] or deck_palette(plan)

def is_finite_number(v: Any) -> bool:
    """图表数值的唯一判定：真实有限数值。bool、字符串、None、NaN、±Infinity 一律不是。

    plan 校验（checks）与定点改稿（revise chart_data）共用这一处，两条路径对同一份数据结论一致。
    """
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)


def non_finite_series(series: List[Dict[str, Any]]) -> List[str]:
    """返回含非有限数值（values 或 points 坐标）的系列名。"""
    bad = []
    for s in series or []:
        flat = list(s.get("values") or []) + [c for pt in (s.get("points") or []) for c in pt]
        if not all(is_finite_number(v) for v in flat):
            bad.append(str(s.get("name")))
    return bad


def chart_data_problems(kind: str, categories: Optional[List[Any]], series: List[Dict[str, Any]], unit: str = "",
                        stacked: Any = None, data_labels: bool = False) -> List[Tuple[str, str, str]]:
    """图表数据契约的唯一实现：新建（checks.check_chart）与既有图表改数（revise.apply_chart）共用。

    返回 [(严重度 HARD|SOFT, 诊断码, 说明)]。规则：系列名与类别名唯一、数值有限、值数=类别数；
    饼/环单系列、不得为负、合计不得为 0、以 % 表示时合计应为 100；百分比堆叠每列合计应为 100。
    """
    out: List[Tuple[str, str, str]] = []
    names = [str(s.get("name")) for s in series]
    cnt = collections.Counter(names)
    if any(c > 1 for c in cnt.values()):
        out.append(("HARD", "E-CHART", f"系列名重复 {sorted(n for n, c in cnt.items() if c > 1)}：名称是颜色、组合线与图例的唯一键"))
    cats = [str(c) for c in (categories or [])]
    ccnt = collections.Counter(cats)
    if any(c > 1 for c in ccnt.values()):
        out.append(("HARD", "E-CHART", f"类别名重复 {sorted(n for n, c in ccnt.items() if c > 1)}：类别是饼/环取色与 color_map 的键"))
    for name in non_finite_series(series):
        out.append(("HARD", "E-CHART", f"系列「{name}」含非有限数值（NaN / Infinity / 布尔 / 非数字）"))
    if categories is None:
        return out
    lengths_ok = True
    for s in series:
        n = len(s.get("values") or [])
        if n != len(cats):
            lengths_ok = False
            out.append(("HARD", "E-CHART", f"系列「{s.get('name')}」有 {n} 个值，类别 {len(cats)} 个"))
    if kind in ("pie", "donut"):
        if len(series) != 1:
            out.append(("HARD", "E-CHART", f"{kind} 只能单系列，现有 {len(series)} 个"))
        for s in series:
            vals = [v for v in (s.get("values") or []) if is_finite_number(v)]
            if any(v < 0 for v in vals):
                out.append(("HARD", "E-CHART", "饼/环不能有负值"))
            if vals and sum(vals) <= 0:
                out.append(("HARD", "E-CHART", f"系列「{s.get('name')}」各扇区合计为 0：占比无定义，改用表格或补数据"))
            elif str(unit or s.get("unit") or "").strip() in ("%", "％", "百分比") and vals and abs(sum(vals) - 100) > 0.5:
                out.append(("SOFT", "W-SUM", f"系列「{s.get('name')}」各扇区合计 {sum(vals):g}%，不是 100%：补“其他”或核对口径；不要让读者自己发现缺口"))
    if stacked == "percent" and kind == "bar" and lengths_ok and not non_finite_series(series):
        bad = []
        for j, cat in enumerate(categories):
            total = sum(float(s["values"][j]) for s in series)
            if abs(total - 100) > 0.5:
                bad.append(f"{cat}={total:g}")
        if bad:
            out.append(("HARD" if data_labels else "SOFT", "E-CHART" if data_labels else "W-CHART",
                        f"百分比堆叠要求每列合计 100，实际 {bad}；图形会归一化而直标仍是原值"))
    return out


def chart_items(chart: Dict[str, Any]) -> List[str]:
    """这张图按什么顺序消费颜色：饼/环按**类别**（扇区），其余按**系列**。着色与图例共用这一份顺序。"""
    kind = chart.get("kind")
    data = chart.get("data") or {}
    if kind == "table":
        return []
    if kind in ("pie", "donut"):
        return [str(c) for c in (data.get("categories") or [])]
    return [str(s.get("name")) for s in (data.get("series") or [])]

def chart_color_map(plan: Dict[str, Any], chart: Dict[str, Any]) -> List[str]:
    """类别/系列 → 颜色的**唯一**映射。编译、校验、核验与自定义图例都读这里，不存在第二份实现。

    优先级（高到低）：
      1. `chart.color_map`  按名称显式指派：改了类别/系列名就对不上，编译期报 E-CHART；
      2. `chart.palette`    本图专用序列（整稿强调色只属于某类数据时用它隔离）；
      3. `plan.palette`     整稿数据色序列，按 chart_items 的顺序循环取用；
      4. `limits.json` 的 defaults.palette。

    主题强调色（accent1）不参与数据着色：主题只服务使用者在 PowerPoint 里新建的对象，
    把它注入图表会让"主题色覆盖数据色"这类错色无从追查。
    """
    items = chart_items(chart)
    if not items:
        return []
    explicit = {str(k): str(v).upper() for k, v in (chart.get("color_map") or {}).items()}
    base = chart_palette(plan, chart) or LIMITS["defaults"]["palette"]
    return [explicit.get(name, base[i % len(base)]) for i, name in enumerate(items)]

def chart_flags(chart: Dict[str, Any]) -> Dict[str, bool]:
    """原生图例与直标的**有效**开关：compile、checks、verify 只读这里，没有第二份缺省。

    缺省值（plan 里没写该键时）：
      - 饼/环：`data_labels` 开（饼带类别名与读数，环只有读数）、`legend` 关。
        旧版两个都缺省为关，饼/环会被编译成既无类别也无数值的图。
      - 柱/线/面积/散点：`data_labels` 关；`legend` 在两个以上系列时开。
      - 表格：两者都不适用，恒为关。
    显式写了的值（含 false）永远优先；是否放行由 `chart_readout_problems` 判断。
    """
    kind = chart.get("kind")
    if kind == "table":
        return {"legend": False, "data_labels": False}
    pie_like = kind in ("pie", "donut")
    n_series = len((chart.get("data") or {}).get("series") or [])
    legend = chart.get("legend")
    labels = chart.get("data_labels")
    return {
        "legend": bool(legend) if legend is not None else (False if pie_like else n_series > 1),
        "data_labels": bool(labels) if labels is not None else pie_like,
    }

def chart_readout_problems(chart: Dict[str, Any], has_legend_element: bool = False) -> List[Tuple[str, str]]:
    """饼/环的读数闭环：每个扇区必须能认出“是什么”，最好还能读到“多少”。

    返回 [(HARD|SOFT, 说明)]。checks 据此阻断，compile 不另立规则。
      - 类别可识别：饼 = 直标（带类别名）/ 原生图例 / legend 元素；环 = 原生图例 / legend 元素（环内直标只有数值）。
      - 数值可读：直标，或带 show_value 的 legend 元素。纯图例只给占比的视觉印象，提示而不阻断。
    """
    kind = chart.get("kind")
    if kind not in ("pie", "donut"):
        return []
    flags = chart_flags(chart)
    out: List[Tuple[str, str]] = []
    names_ok = flags["legend"] or has_legend_element or (kind == "pie" and flags["data_labels"])
    if not names_ok:
        out.append(("HARD", f"{kind} 没有任何类别读数：" + (
            "环内直标只有数值，需要原生图例（legend: true）或 legend 元素" if kind == "donut"
            else "需要直标（data_labels）、原生图例或 legend 元素")))
    elif not flags["data_labels"] and not has_legend_element:
        out.append(("SOFT", f"{kind} 没有数值直标：读者只能凭扇区面积估数；"
                            "开 data_labels，或用 legend 元素并设 show_value"))
    return out

def lightness(hex_color: str) -> float:
    """CIE L*（0–100）：颜色转成灰度后的感知明度。灰度打印、投影与色觉差异下，区分系列靠的是它。"""
    y = rel_luminance(hex_color)
    return 116.0 * (y ** (1 / 3)) - 16.0 if y > 216 / 24389 else y * 24389 / 27

def gray_conflicts(colors: List[str], floor: Optional[float] = None) -> List[Tuple[int, int, float]]:
    """全部两两比较（不只相邻），返回灰度明度差 ΔL* 低于下限的 (i, j, ΔL*)。"""
    floor = LIMITS["chart"]["min_series_delta_lstar"] if floor is None else floor
    out = []
    for i in range(len(colors)):
        for j in range(i + 1, len(colors)):
            d = abs(lightness(colors[i]) - lightness(colors[j]))
            if d < floor:
                out.append((i, j, d))
    return out

# ---------------------------------------------------------------- 图表视觉契约
# 颜色的职责分四类，各有唯一来源，compile / checks / verify / chart_audit 都读这里：
#   语义色  chart_color_map      系列、类别、强调项的身份色（数据）
#   结构色  chart_style          轴、网格、图例、图题的墨色，随**实际承载面**取深/浅
#   标注色  label_ink            直标的墨色，随**标签真正落在的表面**取：落在底上用结构色，
#                                落在色块里用深/白墨，线与点的直标在读得清时沿用系列色
#   承载面  chart_surface        图表后面真实的底：页面底色 + 其下覆盖整张图的色块（同 resolve_bg）
# 旧写法只看页面底色：图放在深色面板上、页面却是浅色时，轴、图例、直标全部错色。

LABEL_INK_MIN = 4.5          # 14pt 直标按正文算（limits.contrast.normal）
INSIDE_POSITIONS = {"ctr", "inEnd", "inBase"}

# PowerPoint 对 dLblPos 的取值限制按图表类型而不同；XSD 只认枚举，不认这张表，
# 写错不会被 XSD 拦下，但 PowerPoint 会提示“需要修复”。来源：ECMA-376 §21.2.2.48 与 python-pptx 分析文档。
def allowed_label_positions(plot_tag: str, grouping: str = "clustered") -> Optional[set]:
    """None 表示该类型不允许写 dLblPos（donut、area）；集合为允许值。"""
    if plot_tag == "barChart":
        return {"ctr", "inEnd", "inBase"} if grouping in ("stacked", "percentStacked") \
            else {"ctr", "inEnd", "inBase", "outEnd"}
    if plot_tag in ("lineChart", "scatterChart"):
        return {"t", "b", "l", "r", "ctr"}
    if plot_tag == "pieChart":
        return {"bestFit", "ctr", "inEnd", "outEnd"}
    return None

def chart_surface(page: Dict[str, Any], el: Dict[str, Any]) -> str:
    """图表（或图例）真实落在的底色。被图片或渐变承载时无法静态判定，退回页面底色并由 checks 提示。"""
    return resolve_bg(page, el) or page_bg(page)

def chart_surface_known(page: Dict[str, Any], el: Dict[str, Any]) -> bool:
    return resolve_bg(page, el) is not None

def chart_style(surface: str) -> Dict[str, str]:
    """结构色：取在该承载面上对比更高的那一套（不是只看亮度阈值，中灰底也能选对）。"""
    cfg = LIMITS["chart"]["style"]
    return cfg["light"] if contrast_ratio(cfg["light"]["body"], surface) >= contrast_ratio(cfg["dark"]["body"], surface) \
        else cfg["dark"]

def chart_mode(surface: str) -> str:
    return "light" if chart_style(surface) is LIMITS["chart"]["style"]["light"] else "dark"

def best_ink(surface: str) -> str:
    """色块里的字：先在深墨与白墨里选高对比的；两者都不到 4.5（中明度色块，如 #C55A11 上白字 4.33）
    就改用纯黑——黑与白对任何底色至少有一个 ≥ 4.58，所以总能读清。"""
    dark = LIMITS["chart"]["style"]["light"]["title"]
    best = max((dark, "#FFFFFF"), key=lambda c: contrast_ratio(c, surface))
    if contrast_ratio(best, surface) >= LABEL_INK_MIN:
        return best
    return max((best, "#000000", "#FFFFFF"), key=lambda c: contrast_ratio(c, surface))

def series_ink(color: str, surface: str, fallback: str) -> str:
    """线/点直标：能用系列色读得清（≥4.5）就用，保持归属；读不清退回结构色，可读性优先。"""
    return color.upper() if contrast_ratio(color, surface) >= LABEL_INK_MIN else fallback

def chart_unit(chart: Dict[str, Any]) -> str:
    """直标与图例的单位：图表级优先，否则取各系列声明的同一个单位。checks 保证多系列单位一致。"""
    unit = str(chart.get("unit") or "").strip()
    if unit:
        return unit
    units = {str(s["unit"]).strip() for s in ((chart.get("data") or {}).get("series") or []) if s.get("unit")}
    return units.pop() if len(units) == 1 else ""

def chart_label_plan(plan: Dict[str, Any], chart: Dict[str, Any], surface: str) -> List[Dict[str, Any]]:
    """直标的唯一决策：每个系列（饼/环为唯一系列）的位置、墨色、衬底、逐点墨色。

    位置决定墨色落在哪种表面上，所以两者必须一起定：
      柱（簇）  outEnd  → 落在承载面，结构色
      柱（堆叠）ctr     → 落在系列色块，深/白墨按该色块取
      面积      不写位置（PowerPoint 不允许），默认居中落在系列色块，深/白墨
      线、散点  t（多线时最后一条 b，避免相近的线标签叠压）→ 落在承载面；多系列沿用系列色（读得清时），否则结构色
      组合折线  t       → 可能压在柱体上，加承载面色衬底，墨色在衬底上取
      饼        outEnd  → 落在承载面，结构色
      环        不写位置（不允许），默认居中落在环带，逐扇区深/白墨
    """
    kind, stacked = chart.get("kind"), chart.get("stacked")
    items, colors = chart_items(chart), chart_color_map(plan, chart)
    body = chart_style(surface)["body"].upper()
    combo = set(chart.get("combo_line_series") or [])
    multi = len(items) > 1
    if kind == "pie":
        return [{"name": items[0] if items else "", "position": "outEnd", "ink": body, "plate": None, "points": None}]
    if kind == "donut":
        return [{"name": "", "position": None, "ink": body, "plate": None,
                 "points": [best_ink(c) for c in colors]}]
    out = []
    for k, (name, color) in enumerate(zip(items, colors)):
        entry = {"name": name, "position": "outEnd", "ink": body, "plate": None, "points": None}
        if kind == "line" or kind == "scatter":
            # 多条线靠得近时同在上方的标签会叠在一起（真实渲染里见过）：最后一条放到线下，其余在线上
            below = kind == "line" and multi and k == len(items) - 1
            entry.update(position="b" if below else "t", ink=series_ink(color, surface, body) if multi else body)
        elif kind == "area":
            entry.update(position=None, ink=best_ink(color))
        elif name in combo:
            entry.update(position="t", plate=surface.upper(), ink=series_ink(color, surface, body))
        elif stacked:
            entry.update(position="ctr", ink=best_ink(color))
        out.append(entry)
    return out

def page_numbers(plan: Dict[str, Any]) -> Dict[str, Any]:
    cfg = dict(plan.get("page_numbers") or {})
    cfg.setdefault("corner", LIMITS["page_number"]["corner"])
    cfg.setdefault("start", LIMITS["page_number"]["start"])
    cfg.setdefault("hide", [])
    return cfg

def pn_box(plan: Dict[str, Any], page: Dict[str, Any]) -> Dict[str, float]:
    w, h = canvas(plan)
    box = LIMITS["page_number"]["box_in"]
    mx, my = LIMITS["page_number"]["margin_in"]["x"], LIMITS["page_number"]["margin_in"]["y"]
    corner = page_numbers(plan)["corner"]
    return {"x": round(w - mx - box["w"], 4),
            "y": round(my if corner == "top_right" else h - my - box["h"], 4),
            "w": box["w"], "h": box["h"]}

# ---------------------------------------------------------------- 几何


def explicit_box(el: Dict[str, Any]) -> Optional[Box]:
    if all(k in el for k in ("x", "y", "w", "h")):
        return (float(el["x"]), float(el["y"]), float(el["w"]), float(el["h"]))
    return None

def anchor_point(box: Box, anchor: str, at: float = 0.5) -> Tuple[float, float]:
    x, y, w, h = box
    t = min(max(at, 0.0), 1.0)
    if anchor == "left":
        return x, y + h * t
    if anchor == "right":
        return x + w, y + h * t
    if anchor == "top":
        return x + w * t, y
    if anchor == "bottom":
        return x + w * t, y + h
    return x + w / 2.0, y + h / 2.0

def rotated_ancestors(page: Dict[str, Any]) -> Dict[str, Tuple[str, ...]]:
    """每个元素身上带旋转的祖先组，用于判断锚点会不会被组旋转带走。"""
    out: Dict[str, Tuple[str, ...]] = {}

    def walk(items: List[Dict[str, Any]], chain: Tuple[str, ...]) -> None:
        for el in items:
            eid = str(el.get("id"))
            out[eid] = chain
            if el.get("type") == "group":
                nxt = chain + (eid,) if float(el.get("rotation") or 0) else chain
                walk(el.get("elements") or [], nxt)

    walk(page.get("elements") or [], ())
    return out

def flatten(page: Dict[str, Any]) -> List[Dict[str, Any]]:
    """声明顺序展开：group 子元素就地展开并带上 group id，供校验与查背景用。"""
    out: List[Dict[str, Any]] = []

    def walk(items: List[Dict[str, Any]], gid: Optional[str]) -> None:
        for el in items:
            if el.get("type") == "group":
                walk(el.get("elements") or [], el.get("id"))
            else:
                out.append({"el": el, "group": gid})

    walk(page.get("elements") or [], None)
    return out

def _union(boxes: List[Box]) -> Optional[Box]:
    if not boxes:
        return None
    x0 = min(b[0] for b in boxes)
    y0 = min(b[1] for b in boxes)
    x1 = max(b[0] + b[2] for b in boxes)
    y1 = max(b[1] + b[3] for b in boxes)
    return (x0, y0, x1 - x0, y1 - y0)

def all_elements(page: Dict[str, Any]) -> List[Dict[str, Any]]:
    """页内全部元素，含 group 自身与其子元素（id 引用与锚点解析用这一份）。"""
    out: List[Dict[str, Any]] = []

    def walk(items: List[Dict[str, Any]]) -> None:
        for el in items:
            out.append(el)
            if el.get("type") == "group":
                walk(el.get("elements") or [])

    walk(page.get("elements") or [])
    return out

def page_geometry(page: Dict[str, Any]) -> Dict[str, Box]:
    """解析每页全部元素的盒子：显式坐标、path 点集、group 子元素并集。锚点在这里解析。"""
    by_id: Dict[str, Dict[str, Any]] = {str(el.get("id")): el for el in all_elements(page)}
    boxes: Dict[str, Box] = {}

    def resolve(eid: str, stack: Tuple[str, ...]) -> Optional[Box]:
        if eid in boxes:
            return boxes[eid]
        el = by_id.get(eid)
        if el is None or eid in stack:
            return None
        kind = el.get("type")
        box: Optional[Box] = None
        if kind == "path":
            pts = [point_xy(p, resolve, stack + (eid,)) for p in el.get("points") or []]
            if any(p is None for p in pts) or len(pts) < 2:
                box = None
            else:
                xs = [p[0] for p in pts]  # type: ignore[index]
                ys = [p[1] for p in pts]  # type: ignore[index]
                box = (min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
        elif kind == "group":
            kids = [resolve(str(c.get("id")), stack + (eid,)) for c in el.get("elements") or []]
            box = _union([b for b in kids if b is not None])
        else:
            box = explicit_box(el)
        if box is not None:
            boxes[eid] = box
        return box

    for eid in by_id:
        resolve(eid, ())
    return boxes

def point_xy(point: Any, resolve, stack: Tuple[str, ...]) -> Optional[Tuple[float, float]]:
    if isinstance(point, (list, tuple)) and len(point) == 2:
        return float(point[0]), float(point[1])
    if isinstance(point, dict):
        target = resolve(str(point.get("el")), stack)
        if target is None:
            return None
        return anchor_point(target, str(point.get("anchor", "center")), float(point.get("at", 0.5)))
    return None

def resolve_points(el: Dict[str, Any], boxes: Dict[str, Box]) -> List[Tuple[float, float]]:
    out: List[Tuple[float, float]] = []
    for point in el.get("points") or []:
        xy = point_xy(point, lambda eid, _s: boxes.get(eid), ())
        if xy is not None:
            out.append(xy)
    return out

def _intersection(a: Box, b: Box) -> Tuple[float, float]:
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    ix = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    return ix, iy

def substantive_overlap(a: Box, b: Box) -> bool:
    ix, iy = _intersection(a, b)
    if min(ix, iy) < LIMITS["geometry"]["overlap_min_edge_in"]:
        return False
    area = ix * iy
    smaller = min(a[2] * a[3], b[2] * b[3])
    if smaller <= 0:
        return False
    return area >= LIMITS["geometry"]["overlap_area_sq_in"] and area / smaller >= LIMITS["geometry"]["overlap_ratio"]

def default_asset_root(anchor: Path, explicit: Optional[str] = None) -> Path:
    """素材根目录：--assets 优先，否则取 anchor（plan 或 edit.json）同级 assets/，再否则取其所在目录。

    compile / verify / revise 共用这一个解析，三者不会各看各的目录。
    """
    if explicit:
        return Path(explicit).resolve()
    anchor = Path(anchor).resolve()
    sub = anchor.parent / "assets"
    return sub if sub.is_dir() else anchor.parent

def resolve_asset(asset_root: Path, image_id: str) -> Optional[Path]:
    """image_id → 素材文件：先按原名，再补常见扩展名。只认这两个规则，且结果必须留在素材根内。

    不做"同名不同扩展名"的兜底：那会在 hero.png 与 hero-2.jpg 并存时悄悄取到另一张图，
    而核验器用的是同一个解析函数，两边一起错，谁也发现不了。

    边界：绝对路径、含 .. 的路径、经符号链接指向素材根之外的文件一律视为找不到。
    plan 与 edit.json 来自外部，image_id 不能成为读取任意本机文件的入口。
    """
    name = str(image_id or "")
    if not name or "\x00" in name or Path(name).is_absolute() or ".." in Path(name.replace("\\", "/")).parts:
        return None
    root = Path(asset_root).resolve()

    def inside(candidate: Path) -> Optional[Path]:
        if not candidate.is_file():
            return None
        real = candidate.resolve()
        try:
            real.relative_to(root)
        except ValueError:
            return None
        return real

    direct = root / name
    hit = inside(direct)
    if hit:
        return hit
    if direct.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}:
        return None
    for ext in (".png", ".jpg", ".jpeg", ".webp", ".PNG", ".JPG", ".JPEG", ".WEBP"):
        hit = inside(root / (name + ext))
        if hit:
            return hit
    return None

# ---------------------------------------------------------------- 文本式样


def text_pad(el: Dict[str, Any]) -> Tuple[float, float]:
    """文本框内边距：缺省 0.05/0.03 in。贴网格轴线的文字写 pad_in: 0，否则光学边缘会偏 0.05in。"""
    if el.get("pad_in") is not None:
        pad = float(el["pad_in"])
        return pad, pad
    return LIMITS["type"]["inset_in"]["x"], LIMITS["type"]["inset_in"]["y"]

def text_style(plan: Dict[str, Any], page: Dict[str, Any], el: Dict[str, Any]) -> Dict[str, Any]:
    kind = el["type"]
    size = float(el.get("font_size") or LIMITS["type"]["default_pt"].get(kind, 18))
    fonts = plan.get("fonts") or {}
    if kind == "number":
        family = el.get("font") or fonts.get("number") or fonts.get("heading")
    else:
        family = el.get("font") or fonts.get("heading" if kind == "title" else "body")
    family = family or LIMITS["defaults"]["fonts"]["heading" if kind in ("title", "number") else "body"]
    color = el.get("color")
    if not color:
        color = LIMITS["defaults"]["ink_on_dark"] if is_dark(page_bg(page)) else LIMITS["defaults"]["ink"]
    pad_x, pad_y = text_pad(el)
    return {
        "size": size,
        "family": family,
        "color": (color or LIMITS["defaults"]["ink"]).upper(),
        "bold": bool(el.get("bold", LIMITS["type"]["bold"].get(kind, False))),
        "line_spacing": float(el.get("line_spacing") or LIMITS["type"]["line_spacing"].get(kind, 1.25)),
        "align": el.get("align", "left"),
        "valign": el.get("valign", "top"),
        "spacing": float(el.get("letter_spacing_pt") or 0),
        "pad": (pad_x, pad_y),
    }

def paragraph_runs(el: Dict[str, Any]) -> Optional[List[Tuple[str, List[Dict[str, Any]]]]]:
    """把 content 与 runs 归一成 [(段文本, [run…]), …]。

    runs 里的 \\n 就是分段：内联样式可以跨段，标题第一行深墨、第二行强调色不必拆成两个文本框。
    拼接与 content 不等时返回 None，由校验报错，不猜。
    """
    content = str(el.get("content", ""))
    paras = content.split("\n")
    runs = el.get("runs")
    if not runs:
        return [(seg, [{"text": seg}] if seg else []) for seg in paras]
    if "".join(str(r.get("text", "")) for r in runs) != content:
        return None
    buckets: List[List[Dict[str, Any]]] = [[]]
    for spec in runs:
        style = {k: v for k, v in spec.items() if k != "text"}
        for i, piece in enumerate(str(spec.get("text", "")).split("\n")):
            if i:
                buckets.append([])
            if piece:
                buckets[-1].append({**style, "text": piece})
    if len(buckets) != len(paras):
        return None
    return [(paras[i], buckets[i]) for i in range(len(paras))]

# ---------------------------------------------------------------- 校验


CARRIERS = ("shape", "divider", "path")

def resolve_bg(page: Dict[str, Any], target: Dict[str, Any]) -> Optional[str]:
    """文字实际落在什么颜色上：页面底色 + 其下覆盖 ≥90% 的色块；其下有图则无法静态判定。"""
    base = page_bg(page)
    box = explicit_box(target)
    if box is None:
        return base
    area = max(box[2] * box[3], 1e-6)
    for item in flatten(page):
        el = item["el"]
        if el is target:
            break
        below = explicit_box(el)
        if below is None or el.get("type") not in ("shape", "divider", "image"):
            continue
        ix, iy = _intersection(below, box)
        cover = ix * iy / area
        if el["type"] == "image" and cover >= 0.6:
            return None
        if el["type"] not in ("shape", "divider") or cover < 0.9:
            continue
        if el.get("gradient"):
            return None
        alpha = float(el.get("fill_alpha", 1.0))
        color = str(el.get("color") or LIMITS["defaults"]["background"]).upper()
        base = color if alpha >= 0.999 else mix_toward(base, color, alpha)
    return base

def image_probe(path: Path, w: float, h: float, focus: Optional[Dict[str, float]]):
    """返回（有效 ppi, 被裁掉的原图比例）。裁掉太多说明画幅选错，不是"裁切很紧"。"""
    try:
        from PIL import Image
        with Image.open(path) as img:
            px_w, px_h = img.size
    except Exception:
        return None, None
    if w <= 0 or h <= 0 or not px_w or not px_h:
        return None, None
    left, right, top, bottom = _crop_geometry(px_w, px_h, w, h, focus)
    loss = 1 - (1 - left - right) * (1 - top - bottom)
    # 有效分辨率按"最终留在画面里的那部分像素"算：裁得越紧，同样的框里 ppi 越高。
    return min(px_w * (1 - left - right) / w, px_h * (1 - top - bottom) / h), loss

def backdrop_image(page: Dict[str, Any], target: Dict[str, Any]) -> Optional[Tuple[Dict[str, Any], bool]]:
    """文字落在哪张图上：返回（图片元素，图与文字之间是否隔着色块/渐变蒙版）；没有图则 None。

    判定与 resolve_bg 一致：图覆盖文字框 ≥ 60%。取声明顺序里最靠上的那张。
    """
    box = explicit_box(target)
    if box is None:
        return None
    area = max(box[2] * box[3], 1e-6)
    found: Optional[Dict[str, Any]] = None
    scrim = False
    for item in flatten(page):
        el = item["el"]
        if el is target:
            break
        below = explicit_box(el)
        if below is None:
            continue
        ix, iy = _intersection(below, box)
        cover = ix * iy / area
        if el.get("type") == "image" and cover >= 0.6:
            found, scrim = el, False
        elif found is not None and el.get("type") in ("shape", "divider") and cover >= 0.3:
            scrim = True
    return (found, scrim) if found is not None else None

_LIN = [((v / 255.0) / 12.92 if v / 255.0 <= 0.03928 else (((v / 255.0) + 0.055) / 1.055) ** 2.4) for v in range(256)]

def _lum_rgb(rgb: Tuple[int, int, int]) -> float:
    return 0.2126 * _LIN[rgb[0]] + 0.7152 * _LIN[rgb[1]] + 0.0722 * _LIN[rgb[2]]

def zone_probe(path: Path, img_el: Dict[str, Any], text_box: Box, page_background: str) -> Optional[Dict[str, Any]]:
    """文字框正下方的真实像素：按裁切、duotone、透明度还原最终落点，返回相对亮度样本与纹理强度。

    这是静态近似（不含渲染引擎的色彩管理，也不含图上其他元素），用来在编译期发现“字落在亮暗交界或高纹理区”，
    结论仍以真实渲染为准。读不了图返回 None。
    """
    img_box = explicit_box(img_el)
    if img_box is None:
        return None
    try:
        from PIL import Image
        with Image.open(path) as src:
            src = src.convert("RGB")
            pw, ph = src.size
            l, r, t, b = _crop_geometry(pw, ph, img_box[2], img_box[3], img_el.get("crop_focus"))
            x0, y0 = max(img_box[0], text_box[0]), max(img_box[1], text_box[1])
            x1 = min(img_box[0] + img_box[2], text_box[0] + text_box[2])
            y1 = min(img_box[1] + img_box[3], text_box[1] + text_box[3])
            if x1 - x0 <= 0.02 or y1 - y0 <= 0.02:
                return None

            def src_px(fx: float, fy: float, total_w: float, total_h: float, lo_x: float, lo_y: float):
                return ((l + ((fx - lo_x) / total_w) * (1 - l - r)) * pw,
                        (t + ((fy - lo_y) / total_h) * (1 - t - b)) * ph)

            ax, ay = src_px(x0, y0, img_box[2], img_box[3], img_box[0], img_box[1])
            bx, by = src_px(x1, y1, img_box[2], img_box[3], img_box[0], img_box[1])
            n = int(LIMITS["image"]["zone"]["samples"])
            region = src.crop((int(max(ax, 0)), int(max(ay, 0)), max(int(bx), int(ax) + 1), max(int(by), int(ay) + 1)))
            region = region.resize((n, n), Image.BOX)
            pixels = list(region.get_flattened_data() if hasattr(region, "get_flattened_data") else region.getdata())
    except Exception:
        return None
    duo = img_el.get("duotone")
    alpha = float(img_el.get("alpha", 1.0))
    bg = hex_rgb(page_background)
    lums: List[float] = []
    grays: List[float] = []
    for px in pixels:
        rgb = px
        if duo:
            g = (0.2126 * px[0] + 0.7152 * px[1] + 0.0722 * px[2]) / 255.0
            d, lt = hex_rgb(duo["dark"]), hex_rgb(duo["light"])
            rgb = tuple(round(d[i] + (lt[i] - d[i]) * g) for i in range(3))
        if alpha < 0.999:
            rgb = tuple(round(bg[i] + (rgb[i] - bg[i]) * alpha) for i in range(3))
        lums.append(_lum_rgb(rgb))
        grays.append((0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]))
    mean = sum(grays) / len(grays)
    std = (sum((g - mean) ** 2 for g in grays) / len(grays)) ** 0.5
    return {"lums": lums, "gray_std": std, "gray_mean": mean}

def zone_contrast(ink_hex: str, probe: Dict[str, Any], limit: float) -> Tuple[float, float]:
    """（达标像素占比，最不利 5% 像素的对比度）。"""
    li = rel_luminance(ink_hex)
    ratios = sorted((max(li, v) + 0.05) / (min(li, v) + 0.05) for v in probe["lums"])
    ok = sum(1 for r in ratios if r >= limit) / len(ratios)
    return ok, ratios[max(int(len(ratios) * 0.05) - 1, 0)]

def corner_adj(radius_in: float, w: float, h: float) -> int:
    """圆角半径 → adj（短边的百分比，0.5 即胶囊）。编译与核验共用这一个换算。"""
    short = max(min(w, h), 1e-6)
    return int(round(min(max(radius_in / short, 0.0), 0.5) * 100000))

def _crop_geometry(src_w: float, src_h: float, w: float, h: float, focus: Optional[Dict[str, float]]):
    fx = 0.5 if not focus else float(focus.get("x", 0.5))
    fy = 0.5 if not focus else float(focus.get("y", 0.5))
    src_ratio, dst_ratio = src_w / src_h, w / h
    l = r = t = b = 0.0
    if src_ratio > dst_ratio:
        trim = 1 - (dst_ratio / src_ratio)
        l, r = trim * fx, trim * (1 - fx)
    elif src_ratio < dst_ratio:
        trim = 1 - (src_ratio / dst_ratio)
        t, b = trim * fy, trim * (1 - fy)
    return l, r, t, b

def rot_box(box: Box, rotation: Any) -> Box:
    """旋转后的外接框。声明的框是没转过的，页角与遮挡要按真实占位算。

    外接框的两边是 |w·cos| + |h·sin| 与 |w·sin| + |h·cos|：只用 cos/sin 会把
    100°–170° 这一类角度算成负数宽度。
    """
    import math

    if not rotation:
        return box
    x, y, w, h = box
    a = math.radians(float(rotation) % 180)
    c, s = abs(math.cos(a)), abs(math.sin(a))
    nw, nh = w * c + h * s, w * s + h * c
    return (x + w / 2 - nw / 2, y + h / 2 - nh / 2, nw, nh)

def pn_surface(plan: Dict[str, Any], page: Dict[str, Any], box: Dict[str, float]) -> str:
    """页码实际落在什么底色上：取最后绘制、且盖住页码中心的不透明色块或页面底色。

    图片与图表上无法靠一个色值判读，遇到就停在页面底色，由 page_number_tone 声明。
    """
    cx, cy = box["x"] + box["w"] / 2, box["y"] + box["h"] / 2
    boxes = page_geometry(page)
    surface = page_bg(page)
    for item in flatten(page):
        el = item["el"]
        if el["type"] in TEXT_TYPES or el["type"] in ("image", "chart", "path"):
            continue
        color = el.get("color")
        if not color:
            continue
        gb = boxes.get(str(el.get("id")))
        if gb is None:
            continue
        gb = rot_box(gb, el.get("rotation"))
        if gb[0] <= cx <= gb[0] + gb[2] and gb[1] <= cy <= gb[1] + gb[3]:
            surface = str(color).upper()
    return surface

def pn_ink(plan: Dict[str, Any], page: Dict[str, Any]) -> str:
    """页码墨色：声明优先；否则按页码实际落点，在两档墨里选能读的那个。

    只看 page.background 会漏掉"底色其实由一张铺满的色块给出"这种页——
    深蓝底上写深墨页码，等于没写。
    """
    box = pn_box(plan, page)
    tone = page.get("page_number_tone")
    if tone:
        return LIMITS["page_number"]["ink"][tone]
    surface = pn_surface(plan, page, box)
    base = "dark" if is_dark(page_bg(page)) else "light"
    ink = LIMITS["page_number"]["ink"][base]
    other = LIMITS["page_number"]["ink"]["light" if base == "dark" else "dark"]
    return ink if contrast_ratio(ink, surface) >= contrast_ratio(other, surface) else other

PPTX_MAIN_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"

def _xml_part_problems(path, infos) -> List[str]:
    """逐个 XML 部件：禁止 DTD/实体声明（XXE、实体膨胀），且必须可被安全解析；否则在入口以 E-FORMAT 拦下，而不是在深处抛栈。"""
    import zipfile
    out: List[str] = []
    budget = LIMITS["package"]["max_xml_total_mb"] * 1024 * 1024   # 全包 XML 解压总预算：单件不大也可能有上万件
    with zipfile.ZipFile(path) as zf:
        for i in infos:
            if not i.filename.lower().endswith((".xml", ".rels")) or i.is_dir():
                continue
            if i.file_size > LIMITS["package"]["max_xml_part_mb"] * 1024 * 1024:  # OOXML 的 XML 部件远小于此；不把超大部件整块读进内存
                out.append(f"部件 {i.filename} 解压后超过 {LIMITS['package']['max_xml_part_mb']} MB，拒绝解析")
                continue
            budget -= i.file_size
            if budget < 0:
                out.append(f"XML 部件解压总量超过 {LIMITS['package']['max_xml_total_mb']} MB，停止解析")
                break
            data = zf.read(i)
            if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
                out.append(f"部件 {i.filename} 含 DTD/实体声明，拒绝解析（OOXML 不使用）")
                continue
            try:
                root = parse_xml(data)
            except etree.XMLSyntaxError as exc:
                out.append(f"部件 {i.filename} 不是合法 XML：{str(exc)[:60]}")
                continue
            # UTF-16 等编码会绕过上面的字节匹配：以解析后的 docinfo 为准
            if root.getroottree().docinfo.doctype or root.getroottree().docinfo.internalDTD is not None:
                out.append(f"部件 {i.filename} 含 DTD/实体声明，拒绝解析（OOXML 不使用）")
    return out

def package_problems(path: Path, check_suffix: bool = True) -> List[str]:
    """文件是否真的是无宏 .pptx：扩展名、主部件内容类型与包内部件必须一致。

    PPTM（宏启用）与扩展名不符的包都在这里拦下；返回空列表才算合格。
    """
    import zipfile

    path = Path(path)
    problems: List[str] = []
    if check_suffix and path.suffix.lower() != ".pptx":
        problems.append(f"扩展名 {path.suffix or '(无)'} 不是 .pptx")
    try:
        with zipfile.ZipFile(path) as zf:
            infos = zf.infolist()
            names = [i.filename for i in infos]
            unpacked = sum(i.file_size for i in infos)
            ct = zf.getinfo("[Content_Types].xml")
            if ct.file_size > LIMITS["package"]["max_xml_part_mb"] * 1024 * 1024:
                return problems + ["[Content_Types].xml 过大，拒绝解析"]
            types = parse_xml(zf.read(ct))
    except (OSError, KeyError, zipfile.BadZipFile, etree.XMLSyntaxError) as exc:
        return problems + [f"不是可读的 OOXML 包：{type(exc).__name__}"]
    if any(n.startswith("/") or ".." in n.replace("\\", "/").split("/") for n in names):
        problems.append("包内含越界部件名（绝对路径或 ..）")
    if len(names) > LIMITS["package"]["max_entries"]:
        problems.append(f"包内部件 {len(names)} 个，超过 {LIMITS['package']['max_entries']}，拒绝处理")
    if unpacked > LIMITS["package"]["max_unpacked_mb"] * 1024 * 1024:
        problems.append(f"解压后超过 {LIMITS['package']['max_unpacked_mb']} MB，拒绝处理（疑似压缩炸弹）")
    ratio_cfg = LIMITS["package"]
    for i in infos:
        if (i.file_size > ratio_cfg["ratio_check_min_mb"] * 1024 * 1024
                and i.file_size / max(i.compress_size, 1) > ratio_cfg["max_part_ratio"]):
            problems.append(f"部件 {i.filename[:80]} 压缩比异常（{i.file_size // max(i.compress_size, 1)}:1），拒绝处理（疑似压缩炸弹）")
            break
    if any(not p.startswith("扩展名") for p in problems):
        return problems          # 越界、超限的包不再往下解压：核验本身不能成为炸弹的引爆点
    problems.extend(_xml_part_problems(path, infos))
    overrides = {o.get("PartName"): o.get("ContentType") or "" for o in types if o.get("PartName")}
    main_ct = overrides.get("/ppt/presentation.xml")
    if main_ct != PPTX_MAIN_CT:
        problems.append(f"主部件内容类型为 {main_ct}，不是 .pptx 格式（可能是宏启用或其他格式）")
    if any("vbaProject" in n for n in names) or any("macroEnabled" in v for v in overrides.values()):
        problems.append("含宏部件（vbaProject）或宏启用内容类型")
    problems.extend(package_integrity(path))
    return problems

def content_types(zf) -> Tuple[Dict[str, str], Dict[str, str]]:
    """([Content_Types].xml) → (Default：扩展名→类型, Override：/部件名→类型)。扩展名小写。"""
    root = parse_xml(zf.read("[Content_Types].xml"))
    defaults = {(d.get("Extension") or "").lower(): d.get("ContentType") or "" for d in root if etree.QName(d).localname == "Default"}
    overrides = {o.get("PartName"): o.get("ContentType") or "" for o in root if etree.QName(o).localname == "Override"}
    return defaults, overrides

def effective_content_types(zf) -> Dict[str, str]:
    """包内每个部件实际生效的内容类型（Override 优先，其次按扩展名的 Default；都没有则为空串）。

    Default 与 Override 的写法互换（python-pptx 保存时会重排）不改变语义，比较要比这个，不比 XML 文本。
    """
    defaults, overrides = content_types(zf)
    out: Dict[str, str] = {}
    for name in zf.namelist():
        if name == "[Content_Types].xml" or name.endswith("/"):
            continue
        ext = name.rsplit(".", 1)[-1].lower() if "." in name.rsplit("/", 1)[-1] else ""
        out[name] = overrides.get("/" + name) or defaults.get(ext, "")
    return out

def _rel_target(source: str, target: str) -> str:
    """关系目标 → 包内部件名。source 为关系所属的源部件（根关系为空串）。"""
    import posixpath
    target = target.split("#", 1)[0]
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))

def package_integrity(path: Path) -> List[str]:
    """OPC 包的结构完整性：机器能证明的“文件会不会被 PowerPoint 判损坏”的那一部分。

    查：CRC、重名部件、每个部件都有内容类型、Override 指向存在的部件、内部关系的目标存在、
    关系文件指向存在的源部件。不能证明：部件内容本身符合 schema、PowerPoint 是否提示修复（需实机）。
    """
    import zipfile

    out: List[str] = []
    try:
        with zipfile.ZipFile(Path(path)) as zf:
            bad = zf.testzip()
            if bad:
                out.append(f"部件 CRC 校验失败：{bad}")
            names = zf.namelist()
            dup = sorted(n for n, c in collections.Counter(names).items() if c > 1)
            if dup:
                out.append(f"包内有重名部件：{dup[:4]}")
            nameset = set(names)
            _defaults, overrides = content_types(zf)
            effective = effective_content_types(zf)
            untyped = sorted(n for n, t in effective.items() if not t and not n.endswith(".rels"))
            if untyped:
                out.append(f"部件没有内容类型（[Content_Types].xml 缺 Default/Override）：{untyped[:4]}")
            ghost = sorted(p for p in overrides if p.lstrip("/") not in nameset)
            if ghost:
                out.append(f"[Content_Types].xml 的 Override 指向不存在的部件：{ghost[:4]}")
            dangling: List[str] = []
            for name in names:
                if not name.endswith(".rels"):
                    continue
                source = name.replace("_rels/", "", 1)[: -len(".rels")] if "_rels/" in name else ""
                if source and source not in nameset and not name.startswith("_rels/"):
                    dangling.append(f"{name} 的源部件 {source} 不存在")
                for rel in parse_xml(zf.read(name)):
                    if rel.get("TargetMode") == "External":
                        continue
                    target = _rel_target(source, rel.get("Target") or "")
                    if target not in nameset:
                        dangling.append(f"{name} → {rel.get('Target')}")
            if dangling:
                out.append(f"关系指向不存在的部件：{dangling[:4]}")
    except (OSError, KeyError, zipfile.BadZipFile, etree.XMLSyntaxError):
        pass  # 打不开的情形由 package_problems 主体报告
    return out

def guard_paths(rep: "Report", inputs: Dict[str, Optional[Path]], outputs: List[Tuple[str, Optional[Path], str]],
                force: bool = False, dirs: Optional[Dict[str, Optional[Path]]] = None) -> None:
    """所有 CLI 共用的路径守卫：输入、输出、报告三类路径只在这一处判定。

    outputs = [(标签, 路径, 必须的扩展名)]。规则：输出不得与任何输入（含硬链接 / 符号链接指向的同一文件）
    相同；输出之间不得相同；扩展名必须匹配（报告只能是 .json，所以不可能覆盖图片或 PPTX）；不得经符号链接写出；
    已存在默认不覆盖。dirs = 专用输出目录（渲染目录），不得是任何输入文件所在目录——那是用户的工作目录，不是产物目录。
    """
    ins = {k: Path(v).resolve() for k, v in inputs.items() if v is not None}

    def same(a: Path, b: Path) -> bool:
        if a == b:
            return True
        try:
            return a.exists() and b.exists() and a.samefile(b)
        except OSError:
            return False

    seen: List[Tuple[str, Path]] = []
    for label, raw, suffix in outputs:
        if raw is None:
            continue
        raw = Path(raw)
        target = raw.resolve()
        if raw.is_symlink():
            rep.hard("E-EXIST", str(raw), f"{label}是符号链接，拒绝经链接写出")
        for in_label, in_path in ins.items():
            if same(target, in_path):
                rep.hard("E-EXIST", str(target), f"{label}与输入（{in_label}）是同一个文件，会覆盖它")
        for other_label, other in seen:
            if same(target, other):
                rep.hard("E-EXIST", str(target), f"{label}与{other_label}同路径，会互相覆盖")
        seen.append((label, target))
        if suffix and target.suffix.lower() != suffix:
            rep.hard("E-FORMAT", str(target), f"{label}只能是 {suffix} 文件；收到 {target.name!r}")
        try:
            if target.is_file() and target.stat().st_nlink > 1:
                rep.hard("E-EXIST", str(target), f"{label}是硬链接（链接数 {target.stat().st_nlink}），拒绝覆盖：它可能同时是别处的输入")
        except OSError:
            pass
        if target.is_dir():
            rep.hard("E-EXIST", str(target), f"{label}路径是目录")
        elif not target.parent.is_dir():
            rep.hard("E-EXIST", str(target.parent), f"{label}所在目录不存在，先建目录")
        elif target.exists() and not force:
            rep.hard("E-EXIST", str(target), f"{label}已存在，未写入；确认要替换就加 --force")
    for label, raw in (dirs or {}).items():
        if raw is None:
            continue
        target = Path(raw).resolve()
        if target.exists() and not target.is_dir():
            rep.hard("E-EXIST", str(target), f"{label}不是目录")
        for in_label, in_path in ins.items():
            if in_path.parent == target:
                rep.hard("E-EXIST", str(target), f"{label}是输入（{in_label}）所在的目录；产物须写进专用子目录，不碰工作目录")
                break


def place_atomic(tmp: Path, dest: Path, force: bool = False) -> None:
    """把已写好的临时文件放到交付名：所有 CLI（compile / revise / scaffold / verify）唯一的落点。

    默认不覆盖：os.link 在目标已存在时原子失败（FileExistsError），关闭“检查后、放入前”被并发进程抢占的窗口；
    force 才 os.replace。成功后临时文件已被移走或删除。
    """
    tmp, dest = Path(tmp), Path(dest)
    if force:
        os.replace(tmp, dest)
    else:
        os.link(tmp, dest)
        tmp.unlink()


def temp_beside(path: Path, suffix: str = ".part") -> Path:
    """与目标同目录的唯一临时文件（O_EXCL 创建，不可预测、不经符号链接、并发互不踩踏）。调用方负责清理。"""
    path = Path(path)
    fd, name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=suffix, dir=str(path.parent))
    os.close(fd)
    umask = os.umask(0)
    os.umask(umask)
    os.chmod(name, 0o666 & ~umask)   # mkstemp 默认 0600；交付物应取普通文件权限，而不是只有属主可读
    return Path(name)


def write_atomic(path: Path, data: bytes, force: bool = False, verify: Any = None) -> None:
    """所有 CLI 共用的落盘：同目录临时文件 → 写后回读校验 → 原子放入。

    默认不覆盖：用 os.link 放入（目标已存在则失败，关闭“检查后、写入前”被抢占的窗口）；force 才 os.replace。
    任何一步失败都清掉临时文件，已有目标一个字节不动。verify(bytes) 返回错误串或 None，写后对回读内容再验一遍。
    """
    path = Path(path)
    fd, tmp_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent))
    tmp = Path(tmp_name)
    try:
        with os.fdopen(fd, "wb") as fh:
            fh.write(data)
            fh.flush()
            os.fsync(fh.fileno())
        back = tmp.read_bytes()
        if hashlib.sha256(back).digest() != hashlib.sha256(data).digest():
            raise OSError("写后回读哈希不一致")
        err = verify(back) if verify else None
        if err:
            raise ValueError(f"写后校验失败：{err}")
        place_atomic(tmp, path, force)
    finally:
        if tmp.exists():
            tmp.unlink()


PENDING = "【待判断】"          # scaffold 起稿时的占位：设计判断还没做，不是理由
SCHEMA_VERSION = 2


def pending_fields(page: Dict[str, Any]) -> List[str]:
    """页面里仍处于“待判断”的关键设计字段（占位符、draft、缺失的构图理由）。"""
    out = []
    if page.get("design_state") == "draft":
        out.append("design_state=draft")
    for k in ("claim", "focus", "composition_rationale"):
        if PENDING in str(page.get(k) or ""):
            out.append(k)
    if not str(page.get("composition_rationale") or "").strip():
        out.append("composition_rationale（缺失）")
    return out


# 15.2 及以前 scaffold 会把配方说明原样写进 composition_rationale；迁移时识别并摘掉，免得被当成真实理由。
LEGACY_GRAMMARS = (
    "标题底边锚在 64% 线，上方整块安静区", "一句话占满视野，竖线是唯一装饰，让人停下来", "巨型序号压暗场，换气与分章",
    "关键数字压住页面，正文只写结论与限定", "图承担读数，旁注只写意义", "两个数字同基线同尺度，落差就是论点",
    "标题与条目同一顶线，竖线钉住两栏", "表格承担精确读数", "只表达先后，不画箭头", "半幅图承载空间，文字留在另一侧",
    "收在决定与下一步", "自由构图：关系决定几何，系统只供网格、色值与字号档",
)


def migrate_plan(plan: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str]]:
    """旧版 plan → 当前版本，返回 (新 plan, 变更说明)。只做可证的事：不凭标题臆造 focus，不删用户写的理由。"""
    import copy
    new, notes = copy.deepcopy(plan), []
    ver = new.get("schema_version", 1)
    if ver > SCHEMA_VERSION:
        raise ValueError(f"schema_version={ver} 比本工具（{SCHEMA_VERSION}）新，拒绝降级迁移")
    for page in new.get("pages") or []:
        pid = page.get("id", "?")
        if page.get("composition_rationale") in LEGACY_GRAMMARS:
            del page["composition_rationale"]
            page["design_state"] = "draft"
            notes.append(f"{pid}: composition_rationale 是配方默认说明，已移除并标 design_state=draft（待 AI 补真实构图理由）")
        if page.get("focus") and page.get("claim") and str(page["focus"]) == str(page["claim"]).split("\n")[0]:
            notes.append(f"{pid}: focus 与 claim 首行相同，疑为标题推断值；未改动，请复核")
    if ver != SCHEMA_VERSION:
        new["schema_version"] = SCHEMA_VERSION
        notes.append(f"schema_version {ver} → {SCHEMA_VERSION}")
    return new, notes


def utf8_stdio() -> None:
    """控制台代码页不是 UTF-8 时（Windows 中文终端很常见），中文与 ⏎ 会抛 UnicodeEncodeError。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass


def alt_text(el: Dict[str, Any]) -> str:
    """元素的替代文字：显式 alt 优先；图片缺省取 asset_reason，图表缺省取图题。装饰元素返回空串。"""
    if el.get("alt"):
        return str(el["alt"])
    if el.get("type") == "image":
        return str(el.get("asset_reason") or "")
    if el.get("type") == "chart":
        return str((el.get("chart") or {}).get("title") or "")
    return ""
