#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt-visual-art-director · 确定性 PPTX 编译器

职责：消费 slide plan（符合 references/contract.md），编译为原生可编辑 PPTX，
并执行客观硬错误与一致性检查（边界、非法尺寸、重叠授权、缺失资源、图表/表格渲染、字号可读性、图表-主张口径核对）。

设计边界：
- 不决定叙事、构图、风格或表达取舍；忠实执行已确定的设计意图（含东亚字体、页面底色、等比裁切、原生图表与原生表格）。
- 硬错误阻断通过并返回精确位置、触发条件、实际结果与错误级别。
- 设计与语义问题不由程序静默篡改，而是反馈给 AI 或人工复核。

用法：
    python scripts/compile_pptx.py plan.json out.pptx [--assets .] [--strict]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

try:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt
except ImportError as exc:  # pragma: no cover - 依赖缺失时给出清晰提示
    print(f"[FATAL] 缺少依赖 python-pptx：{exc}\n        请先安装：pip install python-pptx", file=sys.stderr)
    sys.exit(2)


# ---------------------------------------------------------------------------
# 错误模型
# ---------------------------------------------------------------------------

class ErrorLevel(Enum):
    HARD = "HARD"      # 阻断执行路径
    SOFT = "SOFT"      # 警告，不阻断，建议复核或修复


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


# ---------------------------------------------------------------------------
# 常量与契约默认值（唯一对应 references/contract.md）
# ---------------------------------------------------------------------------

SLIDE_W_IN_16_9 = 13.333
SLIDE_H_IN_16_9 = 7.5
SLIDE_W_IN_4_3 = 10.0
SLIDE_H_IN_4_3 = 7.5
BOUND_TOL_IN = 0.01

DEFAULT_FONTS = {"heading": "微软雅黑", "body": "微软雅黑"}
MIN_TITLE_PT = 28
MIN_BODY_PT = 18
MIN_CAPTION_PT = 12
DEFAULT_NUMBER_PT = 40
DEFAULT_CAPTION_PT = 13
MAX_ELEMENTS_PER_PAGE = 12

REQUIRED_DECK_FIELDS = ["deck_id", "title", "audience", "goal", "core_claim", "pages"]
REQUIRED_PAGE_FIELDS = ["id", "role", "claim", "elements"]
REQUIRED_ELEMENT_FIELDS = ["id", "type", "x", "y", "w", "h"]

VALID_ASPECT_RATIOS = {"16:9", "4:3"}
VALID_ROLES = {"open", "build", "breathe", "turn", "hold", "close"}
VALID_LAYOUTS = {"large_white", "editorial", "data_field", "image_story", "linear", "asymmetric"}
VALID_DENSITIES = {"sparse", "moderate", "dense"}
VALID_EMPHASES = {"low", "medium", "high"}
TEXT_ELEMENT_TYPES = {"title", "body", "text", "number", "caption", "footer"}
VALID_ELEMENT_TYPES = TEXT_ELEMENT_TYPES | {"image", "chart", "shape", "divider"}
VALID_CHART_KINDS = {"bar", "line", "pie", "donut", "table"}


def _slide_dimensions(aspect_ratio: Optional[str]) -> Tuple[float, float]:
    """返回实际幻灯片尺寸（英寸），避免 4:3 页面被 16:9 尺寸误判。"""
    if aspect_ratio == "4:3":
        return SLIDE_W_IN_4_3, SLIDE_H_IN_4_3
    return SLIDE_W_IN_16_9, SLIDE_H_IN_16_9


# ---------------------------------------------------------------------------
# 数值解析与图表口径推导
# ---------------------------------------------------------------------------

def _num(value: Any, where: Optional[str] = None) -> Optional[float]:
    if isinstance(value, bool) or value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


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


def _extract_chart_vectors(chart: Optional[Dict[str, Any]]) -> List[List[float]]:
    if not isinstance(chart, dict) or not isinstance(chart.get("data"), dict):
        return []
    data = chart["data"]
    vectors: List[List[float]] = []
    for s in data.get("series") or []:
        if isinstance(s, dict):
            vec = [fv for v in (s.get("values") or []) if (fv := _num(v)) is not None]
            if vec:
                vectors.append(vec)
    for row in data.get("rows") or []:
        if isinstance(row, list):
            vec = []
            for cell in row:
                fv = _num(cell)
                if fv is not None:
                    vec.append(fv)
                elif isinstance(cell, str):
                    vec.extend(_extract_numbers(cell))
            if vec:
                vectors.append(vec)
    return vectors


def _chart_data_numbers(chart: Optional[Dict[str, Any]]) -> List[float]:
    """提取图表/表格中的原始数值列表（用于错误报告展示）。"""
    return [v for vec in _extract_chart_vectors(chart) for v in vec]


def _chart_verifiable_numbers(chart: Optional[Dict[str, Any]]) -> Set[float]:
    """提取图表/表格原始数值及直接可验证的派生值（百分比换算、求和、差值、涨跌幅、饼图占比）。"""
    series_vectors = _extract_chart_vectors(chart)
    if not series_vectors:
        return set()
    candidates: Set[float] = set()
    all_vals = [v for vec in series_vectors for v in vec]
    for v in all_vals:
        candidates.update({v, abs(v), v * 100.0, abs(v) * 100.0})
    candidates.add(sum(all_vals))

    kind = (chart or {}).get("kind", "bar")
    # 多系列按同一类别纵向汇总与结构占比（如堆叠/簇状图 Q4 合计 118+86=204，占比 118/204≈58%）
    if len(series_vectors) >= 2:
        min_len = min(len(v) for v in series_vectors)
        for k in range(min_len):
            col_total = sum(v[k] for v in series_vectors)
            candidates.add(col_total)
            if col_total != 0:
                for v in series_vectors:
                    col_share = abs(v[k] / col_total) * 100.0
                    candidates.update({col_share, round(col_share, 1), round(col_share, 0)})

    for vec in series_vectors:
        s_total = sum(vec)
        candidates.add(s_total)
        if kind in {"pie", "donut"} and s_total != 0:
            for v in vec:
                share = abs(v / s_total) * 100.0
                candidates.update({share, round(share, 1), round(share, 0)})
        for i in range(len(vec)):
            for j in range(len(vec)):
                if i == j:
                    continue
                diff = abs(vec[j] - vec[i])
                candidates.add(diff)
                if vec[i] != 0:
                    pct = abs((vec[j] - vec[i]) / vec[i]) * 100.0
                    candidates.update({pct, round(pct, 1), round(pct, 0)})
    return candidates


# ---------------------------------------------------------------------------
# 解析与校验
# ---------------------------------------------------------------------------

def load_plan(path: str) -> Tuple[Optional[Dict[str, Any]], List[CompileError]]:
    errors: List[CompileError] = []
    try:
        with open(path, "r", encoding="utf-8") as f:
            plan = json.load(f)
    except FileNotFoundError:
        errors.append(CompileError("E001", ErrorLevel.HARD, path, "plan 文件存在", "文件不存在", "无法读取 plan 文件"))
        return None, errors
    except json.JSONDecodeError as exc:
        errors.append(CompileError("E001", ErrorLevel.HARD, path, "plan 为合法 JSON", f"JSON 解析失败：{exc}", "plan 格式非法"))
        return None, errors
    if not isinstance(plan, dict):
        errors.append(CompileError("E001", ErrorLevel.HARD, path, "plan 为 JSON 对象", f"实际类型：{type(plan).__name__}", "plan 顶层必须是对象"))
        return None, errors
    return plan, errors


def validate_plan(
    plan: Dict[str, Any],
    slide_w_in: Optional[float] = None,
    slide_h_in: Optional[float] = None,
) -> List[CompileError]:
    errors: List[CompileError] = []

    ar = plan.get("aspect_ratio")
    if ar is not None and ar not in VALID_ASPECT_RATIOS:
        errors.append(CompileError(
            "E001", ErrorLevel.SOFT, "deck.aspect_ratio",
            f"aspect_ratio ∈ {sorted(VALID_ASPECT_RATIOS)}",
            f"实际：{ar!r}",
            "画幅比例非法，将按默认 16:9 处理",
        ))
    default_w, default_h = _slide_dimensions(ar)
    w_limit = default_w if slide_w_in is None else slide_w_in
    h_limit = default_h if slide_h_in is None else slide_h_in

    for field in REQUIRED_DECK_FIELDS:
        if field not in plan or plan[field] in (None, ""):
            errors.append(CompileError(
                "E001", ErrorLevel.HARD, f"deck.{field}",
                "顶层必填字段非空",
                f"字段 {field} 缺失或为空",
                "演示缺少必要身份信息，无法编译",
            ))

    pages = plan.get("pages")
    if not isinstance(pages, list) or len(pages) == 0:
        errors.append(CompileError(
            "E001", ErrorLevel.HARD, "deck.pages",
            "pages 为非空数组",
            f"实际：{pages!r}",
            "演示至少需要一页",
        ))
        return errors

    seen_page_ids: set = set()
    for pi, page in enumerate(pages):
        loc = f"page[{pi}]"
        if not isinstance(page, dict):
            errors.append(CompileError("E001", ErrorLevel.HARD, loc, "页面为对象", f"实际类型：{type(page).__name__}", "页面条目必须是对象"))
            continue

        pid = page.get("id")
        if not pid:
            errors.append(CompileError("E001", ErrorLevel.HARD, loc, "页面 id 非空", "缺失 id", "页面必须有唯一标识"))
        elif pid in seen_page_ids:
            errors.append(CompileError("E001", ErrorLevel.HARD, loc, "页面 id 唯一", f"重复 id：{pid}", "页面 id 不得重复"))
        else:
            seen_page_ids.add(pid)

        for field in REQUIRED_PAGE_FIELDS:
            if field == "claim":
                if "claim" not in page or page["claim"] is None or not isinstance(page["claim"], str):
                    errors.append(CompileError(
                        "E001", ErrorLevel.HARD, f"{loc}.claim",
                        "页面 claim 字段存在且为字符串",
                        f"字段 claim 缺失或类型非法：{page.get('claim')!r}",
                        f"页面 {pid or loc} 缺少 claim 字段",
                    ))
                elif page["claim"].strip() == "" and not str(page.get("notes") or "").strip():
                    errors.append(CompileError(
                        "E001", ErrorLevel.SOFT, f"{loc}.claim",
                        "claim 留空时在 notes 记录 open question 或留空依据",
                        "claim 与 notes 均为空",
                        f"页面 {pid or loc} 未立主张且未在 notes 说明原因",
                    ))
            elif field not in page or page[field] in (None, ""):
                errors.append(CompileError(
                    "E001", ErrorLevel.HARD, f"{loc}.{field}",
                    "页面必填字段非空",
                    f"字段 {field} 缺失或为空",
                    f"页面 {pid or loc} 缺少必要字段 {field}",
                ))

        for enum_field, valid_set, desc in (
            ("role", VALID_ROLES, "页面职责"),
            ("layout", VALID_LAYOUTS, "构图方向"),
            ("density", VALID_DENSITIES, "信息密度"),
            ("emphasis", VALID_EMPHASES, "视觉分量"),
        ):
            val = page.get(enum_field)
            if val and val not in valid_set:
                errors.append(CompileError(
                    "E001", ErrorLevel.SOFT, f"{loc}.{enum_field}",
                    f"{enum_field} ∈ {sorted(valid_set)}",
                    f"实际：{val!r}",
                    f"{desc}使用未定义枚举值",
                ))

        elements = page.get("elements")
        if not isinstance(elements, list):
            errors.append(CompileError("E001", ErrorLevel.HARD, f"{loc}.elements", "elements 为数组", f"实际：{type(elements).__name__}", "页面必须有元素列表"))
            continue

        if len(elements) > MAX_ELEMENTS_PER_PAGE:
            errors.append(CompileError(
                "E006", ErrorLevel.SOFT, f"{loc}.elements",
                f"元素数量 ≤ {MAX_ELEMENTS_PER_PAGE}",
                f"实际：{len(elements)}",
                "页面元素较多，请核查是否存在可合并或删减的碎片容器",
            ))

        _validate_elements(page, errors, w_limit, h_limit)

    return errors


def _validate_elements(
    page: Dict[str, Any],
    errors: List[CompileError],
    slide_w_in: float = SLIDE_W_IN_16_9,
    slide_h_in: float = SLIDE_H_IN_16_9,
) -> None:
    pid = page.get("id", "?")
    elements = page.get("elements") or []
    rects: List[Tuple[str, float, float, float, float, bool]] = []

    for ei, el in enumerate(elements):
        loc = f"page[{pid}].element[{ei}]"
        if not isinstance(el, dict):
            errors.append(CompileError("E001", ErrorLevel.HARD, loc, "元素为对象", f"实际类型：{type(el).__name__}", "元素条目必须是对象"))
            continue

        eid = el.get("id")
        if not eid:
            errors.append(CompileError("E001", ErrorLevel.HARD, loc, "元素 id 非空", "缺失 id", "元素必须有唯一标识"))

        for field in REQUIRED_ELEMENT_FIELDS:
            if field not in el or el[field] in (None, ""):
                errors.append(CompileError(
                    "E001", ErrorLevel.HARD, f"{loc}.{field}",
                    "元素必填字段非空",
                    f"字段 {field} 缺失或为空",
                    f"元素 {eid or loc} 缺少必要字段 {field}",
                ))

        etype = el.get("type")
        if etype and etype not in VALID_ELEMENT_TYPES:
            errors.append(CompileError(
                "E001", ErrorLevel.SOFT, f"{loc}.type",
                f"type ∈ {sorted(VALID_ELEMENT_TYPES)}",
                f"实际：{etype!r}",
                "元素类型未定义，将退回文本框渲染",
            ))

        if etype in TEXT_ELEMENT_TYPES:
            if "content" not in el or el.get("content") is None or not isinstance(el.get("content"), str):
                errors.append(CompileError(
                    "E001", ErrorLevel.HARD, f"{loc}.{eid}.content",
                    "文本类元素 content 为字符串",
                    f"实际：{el.get('content')!r}",
                    "文本元素缺少 content 字段",
                ))

        x = _num(el.get("x"), loc)
        y = _num(el.get("y"), loc)
        w = _num(el.get("w"), loc)
        h = _num(el.get("h"), loc)

        if None not in (x, y, w, h):
            if w <= 0 or h <= 0:
                errors.append(CompileError(
                    "E002", ErrorLevel.HARD, f"{loc}.{eid}",
                    "元素宽高均大于 0",
                    f"({x:.2f},{y:.2f},{w:.2f}x{h:.2f})",
                    "元素尺寸非法（宽或高 ≤ 0），阻断编译",
                ))
            elif (
                x < -BOUND_TOL_IN
                or y < -BOUND_TOL_IN
                or x + w > slide_w_in + BOUND_TOL_IN
                or y + h > slide_h_in + BOUND_TOL_IN
            ):
                errors.append(CompileError(
                    "E002", ErrorLevel.HARD, f"{loc}.{eid}",
                    "元素在幻灯片边界内",
                    f"({x:.2f},{y:.2f},{w:.2f}x{h:.2f}) 超出边界 {slide_w_in}x{slide_h_in}",
                    "元素超出幻灯片可绘制区域，阻断编译",
                ))
            else:
                overlap_allowed = bool(el.get("overlap_allowed", False))
                rects.append((str(eid), x, y, x + w, y + h, overlap_allowed))

        if etype in TEXT_ELEMENT_TYPES:
            fs = _num(el.get("font_size"), loc)
            if fs is not None:
                if etype == "title" and fs < MIN_TITLE_PT:
                    errors.append(CompileError(
                        "E005", ErrorLevel.SOFT, f"{loc}.{eid}.font_size",
                        f"标题字号 ≥ {MIN_TITLE_PT}pt",
                        f"实际：{fs:.0f}pt",
                        "标题过小，影响可读性与视觉层级",
                    ))
                elif etype in {"body", "text", "number"} and fs < MIN_BODY_PT:
                    errors.append(CompileError(
                        "E005", ErrorLevel.SOFT, f"{loc}.{eid}.font_size",
                        f"正文/数字字号 ≥ {MIN_BODY_PT}pt",
                        f"实际：{fs:.0f}pt",
                        "正文或数字过小，中文在投影下难以辨认",
                    ))
                elif etype in {"caption", "footer"} and fs < MIN_CAPTION_PT:
                    errors.append(CompileError(
                        "E005", ErrorLevel.SOFT, f"{loc}.{eid}.font_size",
                        f"辅助注释/页脚字号 ≥ {MIN_CAPTION_PT}pt",
                        f"实际：{fs:.0f}pt",
                        "注释或页脚字号过小，无法辨识",
                    ))

        if etype == "chart":
            chart = el.get("chart")
            if not isinstance(chart, dict) or not chart.get("data"):
                errors.append(CompileError(
                    "E004", ErrorLevel.HARD, f"{loc}.{eid}.chart",
                    "图表 chart.data 非空",
                    "chart.data 缺失或为空",
                    "图表元素必须携带数据；语义一致性需人工或 AI 复核",
                ))
            else:
                claim = page.get("claim", "") or ""
                claim_nums = set(_extract_numbers(claim))
                raw_chart_nums = set(_chart_data_numbers(chart))
                verifiable_nums = _chart_verifiable_numbers(chart)
                if claim_nums and verifiable_nums:
                    if not any(abs(c - ch) <= 1e-2 for c in claim_nums for ch in verifiable_nums):
                        errors.append(CompileError(
                            "E008", ErrorLevel.SOFT, f"{loc}.{eid}.chart",
                            "页面主张中的定量数值可由图表数据验证或推导",
                            f"claim 数字：{sorted(claim_nums)}；chart 原始数值：{sorted(raw_chart_nums)}",
                            "图表数据与页面主张的数值未建立可验证对应关系，请复核数字、单位或口径是否冲突",
                        ))

        if etype == "image" and not el.get("image_id"):
            errors.append(CompileError(
                "E003", ErrorLevel.HARD, f"{loc}.{eid}.image_id",
                "图片元素 image_id 非空",
                "image_id 缺失",
                "图片元素未指定资产 ID",
            ))

    # 重叠检查：仅全屏背景自动承载前景；其余元素未显式授权 overlap_allowed=true 时拦截实质重叠
    OVERLAP_AREA_THRESHOLD = 0.5  # sq in
    OVERLAP_MIN_EDGE = 0.05       # in，忽略纯边缘微接触
    OVERLAP_RATIO_THRESHOLD = 0.25

    foreground_rects = []
    for (eid, x, y, x2, y2, overlap_allowed) in rects:
        is_full_bg = (
            x <= 0.02
            and y <= 0.02
            and x2 >= slide_w_in - 0.02
            and y2 >= slide_h_in - 0.02
        )
        if is_full_bg:
            continue
        foreground_rects.append((eid, x, y, x2, y2, overlap_allowed))

    for i in range(len(foreground_rects)):
        for j in range(i + 1, len(foreground_rects)):
            id_a, ax1, ay1, ax2, ay2, a_allowed = foreground_rects[i]
            id_b, bx1, by1, bx2, by2, b_allowed = foreground_rects[j]
            if a_allowed or b_allowed:
                continue
            ox = max(0.0, min(ax2, bx2) - max(ax1, bx1))
            oy = max(0.0, min(ay2, by2) - max(ay1, by1))
            if ox <= OVERLAP_MIN_EDGE or oy <= OVERLAP_MIN_EDGE:
                continue
            area = ox * oy
            min_area = min((ax2 - ax1) * (ay2 - ay1), (bx2 - bx1) * (by2 - by1))
            if area > OVERLAP_AREA_THRESHOLD or (min_area > 0 and area / min_area > OVERLAP_RATIO_THRESHOLD):
                errors.append(CompileError(
                    "E002", ErrorLevel.HARD, f"page[{pid}].overlap",
                    "页面内元素不得实质重叠（未设置 overlap_allowed）",
                    f"{id_a} 与 {id_b} 重叠面积约 {area:.2f} sq in",
                    "相邻元素实质重叠会破坏阅读顺序；如属设计授权图层融合，请在承载层设置 overlap_allowed=true",
                ))


# ---------------------------------------------------------------------------
# PPTX 渲染
# ---------------------------------------------------------------------------

def _hex_to_rgb(hex_color: Optional[str], fallback: Tuple[int, int, int] = (0x26, 0x26, 0x26)) -> RGBColor:
    if not hex_color or not isinstance(hex_color, str):
        return RGBColor(*fallback)
    h = hex_color.strip().lstrip("#")
    if len(h) != 6:
        return RGBColor(*fallback)
    try:
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return RGBColor(*fallback)


def _rel_luminance(hex_color: Optional[str], default: float = 1.0) -> float:
    """计算 HEX 颜色的 WCAG 相对亮度（0.0=纯黑，1.0=纯白）。"""
    if not hex_color or not isinstance(hex_color, str):
        return default
    h = hex_color.strip().lstrip("#")
    if len(h) != 6:
        return default
    try:
        chans = []
        for idx in (0, 2, 4):
            s = int(h[idx:idx + 2], 16) / 255.0
            chans.append(s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4)
        return 0.2126 * chans[0] + 0.7152 * chans[1] + 0.0722 * chans[2]
    except ValueError:
        return default


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


def _apply_run_font(run, font_name: str, size_pt: float, color_hex: Optional[str], bold: bool = False) -> None:
    """按 ECMA-376 规范同时写入西文 <a:latin> 与东亚 <a:ea> 字体定义。"""
    run.font.size = Pt(size_pt)
    run.font.color.rgb = _hex_to_rgb(color_hex)
    if bold:
        run.font.bold = True
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    existing_ea = rPr.find(qn("a:ea"))
    if existing_ea is not None:
        rPr.remove(existing_ea)
    ea = rPr.makeelement(qn("a:ea"), {"typeface": font_name})
    latin = rPr.find(qn("a:latin"))
    if latin is not None:
        latin.addnext(ea)
    else:
        rPr.append(ea)


def _add_textbox(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    content: str,
    font_size: Optional[float],
    color: Optional[str],
    align: Optional[str],
    heading_font: str,
    body_font: str,
    etype: str = "body",
    bold: Optional[bool] = None,
) -> None:
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = MSO_ANCHOR.TOP

    is_display = etype in {"title", "number"}
    default_pt = (
        MIN_TITLE_PT if etype == "title"
        else DEFAULT_NUMBER_PT if etype == "number"
        else DEFAULT_CAPTION_PT if etype in {"caption", "footer"}
        else MIN_BODY_PT
    )
    size = font_size or default_pt
    use_bold = is_display if bold is None else bool(bold)
    font_name = heading_font if is_display else body_font
    line_spacing = 1.15 if is_display else 1.3

    lines = str(content).split("\n")
    for li, line in enumerate(lines):
        p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
        p.line_spacing = line_spacing
        if align:
            p.alignment = {
                "left": PP_ALIGN.LEFT,
                "center": PP_ALIGN.CENTER,
                "right": PP_ALIGN.RIGHT,
                "justify": PP_ALIGN.JUSTIFY,
            }.get(align, PP_ALIGN.LEFT)
        run = p.add_run()
        run.text = line
        _apply_run_font(run, font_name, size, color, bold=use_bold)


def _add_shape(slide, x: float, y: float, w: float, h: float,
               color: Optional[str], shape_kind: str) -> None:
    shape_map = {
        "rect": MSO_SHAPE.RECTANGLE,
        "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
    }
    shape_type = shape_map.get(shape_kind, MSO_SHAPE.RECTANGLE)
    sp = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if color:
        sp.fill.solid()
        sp.fill.fore_color.rgb = _hex_to_rgb(color)
    else:
        sp.fill.background()
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


def _add_image(slide, x: float, y: float, w: float, h: float,
               image_id: Optional[str], asset_root: str) -> bool:
    if not image_id:
        return False
    img_path = _resolve_image_path(image_id, asset_root)
    if img_path is None:
        return False
    try:
        pic = slide.shapes.add_picture(str(img_path), Inches(x), Inches(y), Inches(w), Inches(h))
        img_w, img_h = pic.image.size
        if img_w > 0 and img_h > 0 and w > 0 and h > 0:
            src_ratio = img_w / img_h
            dst_ratio = w / h
            if src_ratio > dst_ratio * 1.005:
                crop = (1.0 - dst_ratio / src_ratio) / 2.0
                pic.crop_left = crop
                pic.crop_right = crop
            elif dst_ratio > src_ratio * 1.005:
                crop = (1.0 - src_ratio / dst_ratio) / 2.0
                pic.crop_top = crop
                pic.crop_bottom = crop
        return True
    except Exception:
        return False


def _compose_chart_caption(chart: Dict[str, Any]) -> str:
    title = str(chart.get("title") or "").strip()
    meta: List[str] = []
    unit = str(chart.get("unit") or "").strip()
    period = str(chart.get("period") or "").strip()
    if unit and unit not in title:
        meta.append(unit)
    if period and period not in title:
        meta.append(period)
    suffix = " · ".join(meta)
    if title and suffix:
        return f"{title}（{suffix}）"
    return title or suffix


def _add_table(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    chart: Dict[str, Any],
    heading_font: str,
    body_font: str,
    palette: Optional[List[str]] = None,
    bg_hex: Optional[str] = None,
) -> bool:
    data = chart.get("data")
    if not isinstance(data, dict):
        return False

    headers: List[str] = []
    rows: List[List[str]] = []
    if isinstance(data.get("headers"), list) and isinstance(data.get("rows"), list) and data["headers"] and data["rows"]:
        headers = [str(c) for c in data["headers"]]
        for r in data["rows"]:
            if not isinstance(r, list) or len(r) != len(headers):
                return False
            rows.append([str(c) for c in r])
    elif isinstance(data.get("categories"), list) and isinstance(data.get("series"), list) and data["categories"] and data["series"]:
        cats = [str(c) for c in data["categories"]]
        series_list = data["series"]
        headers = ["指标 / 类别"] + [str(s.get("name", f"系列{i+1}")) for i, s in enumerate(series_list) if isinstance(s, dict)]
        if len(headers) < 2:
            return False
        for s in series_list:
            if not isinstance(s, dict) or not isinstance(s.get("values"), list) or len(s["values"]) != len(cats):
                return False
        for ci, cat in enumerate(cats):
            rows.append([cat] + [str(s["values"][ci]) for s in series_list])
    else:
        return False

    dark = _rel_luminance(bg_hex, default=1.0) < 0.18
    ink_primary = "#F3F4F6" if dark else "#111827"
    ink_body = "#D1D5DB" if dark else "#374151"
    rule_strong = "#6B7280" if dark else "#4B5563"
    rule_subtle = "#2E3440" if dark else "#E5E7EB"

    caption = _compose_chart_caption(chart)
    table_y, table_h = y, h
    if caption and h > 1.0:
        cap_h = min(0.45, h * 0.16)
        _add_textbox(slide, x, y, w, cap_h, caption, 15, ink_primary, "left",
                     heading_font, body_font, etype="body", bold=True)
        table_y = y + cap_h + 0.06
        table_h = max(0.5, h - cap_h - 0.06)

    n_rows = 1 + len(rows)
    n_cols = len(headers)
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(table_y), Inches(w), Inches(table_h))
    tbl = tbl_shape.table

    # 清除内置带状/外框样式，回归“场域 > 对齐 > 发丝线 > 字体层级”无竖框三线秩序
    tblPr = tbl._tbl.find(qn("a:tblPr"))
    if tblPr is not None:
        for attr in ("bandRow", "bandCol", "firstRow", "lastRow", "firstCol", "lastCol"):
            tblPr.set(attr, "0")
        style_id = tblPr.find(qn("a:tableStyleId"))
        if style_id is not None:
            style_id.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"

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
        _apply_run_font(run, heading_font, 13.5, ink_primary, bold=True)

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
            _apply_run_font(run, body_font, 14, ink_primary if ci == 0 else ink_body, bold=(ci == 0))
    return True


def _add_chart(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    chart: Optional[Dict[str, Any]],
    heading_font: str = DEFAULT_FONTS["heading"],
    body_font: str = DEFAULT_FONTS["body"],
    palette: Optional[List[str]] = None,
    bg_hex: Optional[str] = None,
) -> bool:
    if not isinstance(chart, dict):
        return False
    kind = chart.get("kind")
    if kind not in VALID_CHART_KINDS:
        return False
    if kind == "table":
        return _add_table(slide, x, y, w, h, chart, heading_font, body_font, palette=palette, bg_hex=bg_hex)

    data = chart.get("data")
    if not isinstance(data, dict):
        return False
    categories = data.get("categories")
    series_list = data.get("series")
    if not isinstance(series_list, list) or len(series_list) == 0:
        return False

    dark = _rel_luminance(bg_hex, default=1.0) < 0.18
    ink_title = "#F3F4F6" if dark else "#111827"
    ink_tick = "#D1D5DB" if dark else "#4B5563"
    rule_grid = "#2E3440" if dark else "#E5E7EB"

    try:
        chart_data = CategoryChartData()
        cat_list = list(categories) if isinstance(categories, list) and len(categories) > 0 else None
        if cat_list is not None:
            chart_data.categories = cat_list

        for s in series_list:
            if not isinstance(s, dict):
                return False
            raw_vals = s.get("values")
            if not isinstance(raw_vals, list) or len(raw_vals) == 0:
                return False
            if cat_list is not None and len(raw_vals) != len(cat_list):
                return False
            parsed_vals: List[float] = []
            for v in raw_vals:
                fv = _num(v)
                if fv is None:
                    return False
                parsed_vals.append(fv)
            chart_data.add_series(str(s.get("name", "系列")), parsed_vals)

        xl_type = {
            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE,
            "donut": XL_CHART_TYPE.DOUGHNUT,
        }[kind]
        gframe = slide.shapes.add_chart(xl_type, Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
        chart_obj = gframe.chart

        # 去除图表外框与绘图区底色容器，融入页面连续场域
        cs = chart_obj._chartSpace
        cs_spPr = cs.find(qn("c:spPr"))
        if cs_spPr is None:
            cs_spPr = cs.makeelement(qn("c:spPr"), {})
            cs.append(cs_spPr)
        cs_spPr.append(cs_spPr.makeelement(qn("a:noFill"), {}))
        cs_ln = cs_spPr.makeelement(qn("a:ln"), {})
        cs_ln.append(cs_ln.makeelement(qn("a:noFill"), {}))
        cs_spPr.append(cs_ln)
        try:
            chart_obj.plot_area.format.fill.background()
            chart_obj.plot_area.format.line.fill.background()
        except Exception:
            pass

        caption = _compose_chart_caption(chart)
        chart_obj.has_title = bool(caption)
        if caption:
            tf = chart_obj.chart_title.text_frame
            tf.clear()
            run = tf.paragraphs[0].add_run()
            run.text = caption
            _apply_run_font(run, heading_font, 15, ink_title, bold=True)

        if chart_obj.has_legend:
            chart_obj.legend.font.name = body_font
            chart_obj.legend.font.size = Pt(11.5)
            chart_obj.legend.font.color.rgb = _hex_to_rgb(ink_tick)

        for axis_attr in ("category_axis", "value_axis"):
            try:
                axis = getattr(chart_obj, axis_attr)
                axis.tick_labels.font.name = body_font
                axis.tick_labels.font.size = Pt(12)
                axis.tick_labels.font.color.rgb = _hex_to_rgb(ink_tick)
                axis.format.line.color.rgb = _hex_to_rgb(rule_grid)
                axis.format.line.width = Pt(0.75)
                if axis_attr == "value_axis":
                    axis.has_major_gridlines = True
                    axis.major_gridlines.format.line.color.rgb = _hex_to_rgb(rule_grid)
                    axis.major_gridlines.format.line.width = Pt(0.5)
            except Exception:
                pass

        if palette and kind in {"bar", "line"}:
            series_colors = [
                c for c in palette
                if isinstance(c, str)
                and (_rel_luminance(c, 0.5) > 0.08 if dark else _rel_luminance(c, 0.5) < 0.85)
            ]
            for si, ser in enumerate(chart_obj.series):
                if not series_colors:
                    break
                hex_c = series_colors[si % len(series_colors)]
                if kind == "bar":
                    ser.format.fill.solid()
                    ser.format.fill.fore_color.rgb = _hex_to_rgb(hex_c)
                elif kind == "line":
                    ser.format.line.color.rgb = _hex_to_rgb(hex_c)
        return True
    except Exception:
        return False


def compile_plan(
    plan: Dict[str, Any],
    out_path: str,
    asset_root: str = ".",
    strict: bool = False,
) -> Tuple[bool, List[CompileError]]:
    """返回 (是否成功, 错误列表)。strict=True 时 SOFT 也视为阻断。"""
    ar = plan.get("aspect_ratio", "16:9")
    slide_w_in, slide_h_in = _slide_dimensions(ar)

    errors = validate_plan(plan, slide_w_in, slide_h_in)
    hard_errors = [e for e in errors if e.level == ErrorLevel.HARD]
    if hard_errors:
        return False, errors

    fonts = plan.get("fonts") or {}
    heading_font = fonts.get("heading", DEFAULT_FONTS["heading"])
    body_font = fonts.get("body", DEFAULT_FONTS["body"])
    palette = plan.get("palette") if isinstance(plan.get("palette"), list) else None

    prs = Presentation()
    prs.slide_width = Inches(slide_w_in)
    prs.slide_height = Inches(slide_h_in)
    blank_layout = prs.slide_layouts[6]

    for page in plan.get("pages", []):
        slide = prs.slides.add_slide(blank_layout)
        bg_hex = page.get("background")
        if bg_hex:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = _hex_to_rgb(bg_hex, fallback=(0xFF, 0xFF, 0xFF))

        pid = page.get("id", "?")
        elements = page.get("elements") or []
        ordered = sorted(elements, key=lambda e: e.get("z_index", 0))

        for el in ordered:
            etype = el.get("type", "text")
            eid = el.get("id", "?")
            raw_x = _num(el.get("x"), pid) or 0.0
            raw_y = _num(el.get("y"), pid) or 0.0
            raw_w = _num(el.get("w"), pid) or 1.0
            raw_h = _num(el.get("h"), pid) or 0.4
            x = max(0.0, raw_x)
            y = max(0.0, raw_y)
            w = min(raw_w, slide_w_in - x) if x < slide_w_in else raw_w
            h = min(raw_h, slide_h_in - y) if y < slide_h_in else raw_h

            content = el.get("content", "")
            fs = _num(el.get("font_size"), None)
            color = el.get("color")
            align = el.get("align")
            bold = el.get("bold") if isinstance(el.get("bold"), bool) else None

            if etype in TEXT_ELEMENT_TYPES:
                _add_textbox(
                    slide, x, y, w, h, content, fs, color, align,
                    heading_font, body_font, etype=etype, bold=bold,
                )
            elif etype == "image":
                ok = _add_image(slide, x, y, w, h, el.get("image_id"), asset_root)
                if not ok:
                    _add_shape(slide, x, y, w, h, "#DDDDDD", "rect")
                    _add_textbox(
                        slide, x + 0.05, y + 0.05, max(0.1, w - 0.1), max(0.1, h - 0.1),
                        f"[缺失资产：{el.get('image_id')}]", 12, "#666666", "left",
                        heading_font, body_font, etype="caption",
                    )
                    errors.append(CompileError(
                        "E003", ErrorLevel.HARD, f"page[{pid}].element[{eid}].image",
                        f"资产 {el.get('image_id')} 存在",
                        "文件未找到",
                        "图片资源缺失，已绘制占位框，需作者替换",
                    ))
            elif etype == "chart":
                ok = _add_chart(
                    slide, x, y, w, h, el.get("chart"),
                    heading_font=heading_font, body_font=body_font, palette=palette, bg_hex=bg_hex,
                )
                if not ok:
                    _add_shape(slide, x, y, w, h, "#EEEEEE", "rect")
                    _add_textbox(
                        slide, x + 0.05, y + 0.05, max(0.1, w - 0.1), max(0.1, h - 0.1),
                        "[图表渲染失败：数据或类型非法]", 12, "#666666", "left",
                        heading_font, body_font, etype="caption",
                    )
                    errors.append(CompileError(
                        "E007", ErrorLevel.HARD, f"page[{pid}].element[{eid}].chart",
                        f"图表 {eid} 类型合法且数据结构完整",
                        f"kind={el.get('chart', {}).get('kind')!r} 或数据非法，渲染失败",
                        "图表/表格渲染失败，已绘制占位框，请核查 kind 与 categories/series/rows",
                    ))
            elif etype in {"shape", "divider"}:
                _add_shape(slide, x, y, w, h, color, el.get("shape_kind", "rect"))
            else:
                _add_textbox(
                    slide, x, y, w, h, content, fs, color, align,
                    heading_font, body_font, etype="text", bold=bold,
                )

    out_file = Path(out_path)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_file))

    success = not any(e.level == ErrorLevel.HARD for e in errors)
    if strict and any(e.level == ErrorLevel.SOFT for e in errors):
        success = False
    return success, errors


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="编译 slide plan 为 PPTX 并做硬错误检查")
    parser.add_argument("plan", help="slide plan JSON 路径")
    parser.add_argument("output", help="输出 PPTX 路径")
    parser.add_argument("--assets", default=".", help="素材根目录（默认当前目录）")
    parser.add_argument("--strict", action="store_true",
                        help="严格模式：SOFT 级别问题也视为失败")
    args = parser.parse_args(argv)

    plan, load_errors = load_plan(args.plan)
    errors: List[CompileError] = list(load_errors)
    if plan is None:
        for e in errors:
            print(str(e), file=sys.stderr)
        return 1

    success, comp_errors = compile_plan(plan, args.output, asset_root=args.assets, strict=args.strict)
    errors.extend(comp_errors)

    if errors:
        print(f"\n共发现 {len(errors)} 个问题：", file=sys.stderr)
        for e in errors:
            print(str(e), file=sys.stderr)

    if not success:
        print("\n编译失败：存在未解决的阻断问题。", file=sys.stderr)
        return 1
    print(f"编译成功：{args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
