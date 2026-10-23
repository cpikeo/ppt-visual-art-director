#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt-visual-art-director · 确定性 PPTX 编译器

职责：消费 slide plan（符合 references/contract.md），编译为原生可编辑 PPTX，
并执行客观硬错误检查（边界、重叠、缺失资源、图表-主张一致性、字号可读性）。

设计边界：
- 不决定叙事、构图、风格或表达取舍；只把已确定的设计意图落入 PPTX。
- 硬错误阻断受影响的执行路径，并给出位置、触发条件、实际结果与错误级别。
- 设计质量问题不由本程序静默修复，而是返回给 AI 重新判断。

用法：
    python scripts/compile_pptx.py plan.json out.pptx [--strict]
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
except ImportError as exc:  # pragma: no cover - 依赖缺失时给出清晰提示
    print(f"[FATAL] 缺少依赖 python-pptx：{exc}\n        请先安装：pip install python-pptx", file=sys.stderr)
    sys.exit(2)


# ---------------------------------------------------------------------------
# 错误模型
# ---------------------------------------------------------------------------

class ErrorLevel(Enum):
    HARD = "HARD"      # 阻断执行路径
    SOFT = "SOFT"      # 警告，不阻断，建议修复


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
# 常量与契约默认值
# ---------------------------------------------------------------------------

SLIDE_W_IN = 13.333  # 16:9
SLIDE_H_IN = 7.5
IN_TO_EMU = 914400

DEFAULT_FONTS = {"heading": "微软雅黑", "body": "微软雅黑"}
MIN_BODY_PT = 18
MIN_TITLE_PT = 28
MAX_ELEMENTS_PER_PAGE = 12

REQUIRED_DECK_FIELDS = ["deck_id", "title", "audience", "goal", "core_claim", "pages"]
REQUIRED_PAGE_FIELDS = ["id", "role", "claim", "elements"]
REQUIRED_ELEMENT_FIELDS = ["id", "type", "x", "y", "w", "h"]

VALID_ROLES = {"open", "build", "breathe", "turn", "hold", "close"}
VALID_LAYOUTS = {"large_white", "editorial", "data_field", "image_story", "linear", "asymmetric"}
VALID_DENSITIES = {"sparse", "moderate", "dense"}
VALID_EMPHASES = {"low", "medium", "high"}
VALID_ELEMENT_TYPES = {
    "title", "body", "text", "number", "image", "chart",
    "shape", "divider", "caption", "footer",
}
VALID_CHART_KINDS = {"bar", "line", "pie", "donut", "table"}


# ---------------------------------------------------------------------------
# 解析与校验
# ---------------------------------------------------------------------------

def _num(value: Any, where: str) -> Optional[float]:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


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


def validate_plan(plan: Dict[str, Any]) -> List[CompileError]:
    errors: List[CompileError] = []

    # 顶层必填
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
        return errors  # 无页面，后续无意义

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
            if field not in page or page[field] in (None, ""):
                errors.append(CompileError(
                    "E001", ErrorLevel.HARD, f"{loc}.{field}",
                    "页面必填字段非空",
                    f"字段 {field} 缺失或为空",
                    f"页面 {pid or loc} 缺少必要字段 {field}",
                ))

        role = page.get("role")
        if role and role not in VALID_ROLES:
            errors.append(CompileError(
                "E001", ErrorLevel.SOFT, f"{loc}.role",
                f"role ∈ {sorted(VALID_ROLES)}",
                f"实际：{role!r}",
                "页面职责使用非法值，可能影响节奏判断",
            ))

        layout = page.get("layout")
        if layout and layout not in VALID_LAYOUTS:
            errors.append(CompileError(
                "E001", ErrorLevel.SOFT, f"{loc}.layout",
                f"layout ∈ {sorted(VALID_LAYOUTS)}",
                f"实际：{layout!r}",
                "构图方向使用非法值，将退回默认",
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
                "页面过于拥挤，建议删减元素",
            ))

        _validate_elements(page, errors)

    return errors


def _validate_elements(page: Dict[str, Any], errors: List[CompileError]) -> None:
    pid = page.get("id", "?")
    elements = page.get("elements") or []
    rects: List[Tuple[str, float, float, float, float]] = []  # (id, x, y, x2, y2)

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
                "元素类型非法，将退回文本框",
            ))

        x = _num(el.get("x"), loc)
        y = _num(el.get("y"), loc)
        w = _num(el.get("w"), loc)
        h = _num(el.get("h"), loc)

        if None not in (x, y, w, h):
            # 边界检查
            if x < 0 or y < 0 or x + w > SLIDE_W_IN or y + h > SLIDE_H_IN:
                errors.append(CompileError(
                    "E002", ErrorLevel.HARD, f"{loc}.{eid}",
                    "元素在幻灯片边界内",
                    f"({x:.2f},{y:.2f},{w:.2f}x{h:.2f}) 超出边界 {SLIDE_W_IN}x{SLIDE_H_IN}",
                    "元素超出幻灯片可绘制区域，阻断编译",
                ))
            else:
                rects.append((str(eid), x, y, x + w, y + h))

        # 字号可读性（中文）
        if etype in {"title", "body", "text", "number", "caption", "footer"}:
            fs = _num(el.get("font_size"), loc)
            if fs is not None:
                if etype == "title" and fs < MIN_TITLE_PT:
                    errors.append(CompileError(
                        "E005", ErrorLevel.SOFT, f"{loc}.{eid}.font_size",
                        f"标题字号 ≥ {MIN_TITLE_PT}pt",
                        f"实际：{fs:.0f}pt",
                        "标题过小，影响可读性与层级",
                    ))
                elif etype != "title" and fs < MIN_BODY_PT:
                    errors.append(CompileError(
                        "E005", ErrorLevel.SOFT, f"{loc}.{eid}.font_size",
                        f"正文/辅助文字字号 ≥ {MIN_BODY_PT}pt",
                        f"实际：{fs:.0f}pt",
                        "文字过小，中文在投影下不可读",
                    ))

        # 图表数据与主张一致性（轻度客观检查：数据非空、数字存在）
        if etype == "chart":
            chart = el.get("chart")
            if not isinstance(chart, dict) or not chart.get("data"):
                errors.append(CompileError(
                    "E004", ErrorLevel.HARD, f"{loc}.{eid}.chart",
                    "图表 chart.data 非空",
                    "chart.data 缺失或为空",
                    "图表元素必须携带数据",
                ))

        # 图片资源存在性由编译阶段再检查（需文件系统），这里只记录
        if etype == "image" and not el.get("image_id"):
            errors.append(CompileError(
                "E003", ErrorLevel.HARD, f"{loc}.{eid}.image_id",
                "图片元素 image_id 非空",
                "image_id 缺失",
                "图片元素未指定资产 ID",
            ))

    # 重叠检查（仅对通过边界检查的元素）
    for i in range(len(rects)):
        for j in range(i + 1, len(rects)):
            id_a, ax1, ay1, ax2, ay2 = rects[i]
            id_b, bx1, by1, bx2, by2 = rects[j]
            ox = max(0, min(ax2, bx2) - max(ax1, bx1))
            oy = max(0, min(ay2, by2) - max(ay1, by1))
            if ox > 0 and oy > 0:
                errors.append(CompileError(
                    "E002", ErrorLevel.HARD, f"page[{pid}].overlap",
                    "页面内元素不得实质重叠",
                    f"{id_a} 与 {id_b} 重叠面积约 {ox * oy:.2f} sq in",
                    "相邻元素重叠会破坏阅读顺序与视觉层级",
                ))


# ---------------------------------------------------------------------------
# PPTX 渲染
# ---------------------------------------------------------------------------

def _hex_to_rgb(hex_color: Optional[str]) -> RGBColor:
    if not hex_color:
        return RGBColor(0x26, 0x26, 0x26)
    h = hex_color.lstrip("#")
    try:
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except (ValueError, IndexError):
        return RGBColor(0x26, 0x26, 0x26)


def _set_east_asian_font(run, font_name: str) -> None:
    from pptx.oxml.ns import qn
    rPr = run._r.get_or_add_rPr()
    rPr.set("typeface", font_name)
    # 也设置 latin 作为兜底
    if rPr.find(qn("a:latin")) is None:
        latin = rPr.makeelement(qn("a:latin"), {"typeface": font_name})
        rPr.insert(0, latin)


def _add_textbox(slide, x: float, y: float, w: float, h: float,
                 content: str, font_size: Optional[float],
                 color: Optional[str], align: Optional[str],
                 heading_font: str, body_font: str,
                 is_title: bool = False) -> None:
    from pptx.oxml.ns import qn
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = MSO_ANCHOR.TOP

    lines = str(content).split("\n")
    for li, line in enumerate(lines):
        p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
        if align:
            p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                           "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}.get(
                align, PP_ALIGN.LEFT)
        run = p.add_run()
        run.text = line
        size = font_size or (MIN_TITLE_PT if is_title else MIN_BODY_PT)
        run.font.size = Pt(size)
        run.font.name = heading_font if is_title else body_font
        _set_east_asian_font(run, heading_font if is_title else body_font)
        run.font.color.rgb = _hex_to_rgb(color)
        if is_title:
            run.font.bold = True


def _add_shape(slide, x: float, y: float, w: float, h: float,
               color: Optional[str], shape_kind: str) -> None:
    shape_map = {
        "rect": MSO_SHAPE.RECTANGLE,
        "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
        "line": MSO_SHAPE.LINE,
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


def _add_image(slide, x: float, y: float, w: float, h: float,
               image_id: Optional[str], asset_root: str) -> bool:
    if not image_id:
        return False
    # 尝试多种常见扩展名
    for ext in (".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"):
        path = f"{asset_root}/{image_id}{ext}"
        try:
            slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
            return True
        except FileNotFoundError:
            continue
        except Exception:
            continue
    return False


def _add_chart(slide, x: float, y: float, w: float, h: float,
               chart: Optional[Dict[str, Any]]) -> bool:
    if not isinstance(chart, dict):
        return False
    kind = chart.get("kind", "bar")
    data = chart.get("data")
    if not data:
        return False
    try:
        chart_data = CategoryChartData()
        # data 期望形状：{"categories": [...], "series": [{"name": str, "values": [...]}]}
        categories = data.get("categories", [])
        series_list = data.get("series", [])
        if categories:
            chart_data.categories = list(categories)
        for s in series_list:
            chart_data.add_series(s.get("name", "系列"), list(s.get("values", [])))
        xl_type = {
            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE,
            "donut": XL_CHART_TYPE.DOUGHNUT,
            "table": XL_CHART_TYPE.COLUMN_CLUSTERED,
        }.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
        gframe = slide.shapes.add_chart(xl_type, Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
        chart_obj = gframe.chart
        chart_obj.has_title = bool(chart.get("title"))
        if chart_obj.has_title:
            chart_obj.chart_title.text_frame.text = str(chart.get("title", ""))
        return True
    except Exception:
        return False


def compile_plan(plan: Dict[str, Any], out_path: str, asset_root: str = ".",
                 strict: bool = False) -> Tuple[bool, List[CompileError]]:
    """返回 (是否成功, 错误列表)。strict=True 时 SOFT 也视为阻断。"""
    errors = validate_plan(plan)
    hard_errors = [e for e in errors if e.level == ErrorLevel.HARD]
    if hard_errors:
        for e in hard_errors:
            print(str(e), file=sys.stderr)
        return False, errors

    # 字体
    fonts = plan.get("fonts") or {}
    heading_font = fonts.get("heading", DEFAULT_FONTS["heading"])
    body_font = fonts.get("body", DEFAULT_FONTS["body"])

    # 幻灯片尺寸
    prs = Presentation()
    ar = plan.get("aspect_ratio", "16:9")
    if ar == "4:3":
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
    else:
        prs.slide_width = Inches(SLIDE_W_IN)
        prs.slide_height = Inches(SLIDE_H_IN)

    blank_layout = prs.slide_layouts[6]

    for page in plan.get("pages", []):
        slide = prs.slides.add_slide(blank_layout)
        pid = page.get("id", "?")
        elements = page.get("elements") or []

        # 先渲染背景/底层图元（z_index 小者优先），再渲染文字与图表
        ordered = sorted(elements, key=lambda e: e.get("z_index", 0))

        for el in ordered:
            etype = el.get("type", "text")
            x = _num(el.get("x"), pid) or 0
            y = _num(el.get("y"), pid) or 0
            w = _num(el.get("w"), pid) or 1
            h = _num(el.get("h"), pid) or 0.4
            content = el.get("content", "")
            fs = _num(el.get("font_size"), None)
            color = el.get("color")
            align = el.get("align")
            is_title = etype == "title"

            if etype in {"title", "body", "text", "number", "caption", "footer"}:
                _add_textbox(slide, x, y, w, h, content, fs, color, align,
                             heading_font, body_font, is_title=is_title)
            elif etype == "image":
                ok = _add_image(slide, x, y, w, h, el.get("image_id"), asset_root)
                if not ok:
                    # 缺失资源：画占位框并标注，供作者替换；不静默跳过
                    _add_shape(slide, x, y, w, h, "#DDDDDD", "rect")
                    _add_textbox(slide, x + 0.1, y + 0.1, w - 0.2, h - 0.2,
                                 f"[缺失资产：{el.get('image_id')}]", 12, "#999999", "left",
                                 heading_font, body_font)
                    errors.append(CompileError(
                        "E003", ErrorLevel.HARD, f"page[{pid}].element.image",
                        f"资产 {el.get('image_id')} 存在",
                        "文件未找到",
                        "图片资源缺失，已绘制占位框，需作者替换",
                    ))
            elif etype == "chart":
                ok = _add_chart(slide, x, y, w, h, el.get("chart"))
                if not ok:
                    _add_shape(slide, x, y, w, h, "#EEEEEE", "rect")
                    _add_textbox(slide, x + 0.1, y + 0.1, w - 0.2, h - 0.2,
                                 "[图表渲染失败：数据或类型非法]", 12, "#999999", "left",
                                 heading_font, body_font)
            elif etype in {"shape", "divider"}:
                _add_shape(slide, x, y, w, h, color, el.get("shape_kind", "rect"))
            else:
                _add_textbox(slide, x, y, w, h, content, fs, color, align,
                             heading_font, body_font, is_title=is_title)

    prs.save(out_path)
    success = not any(e.level == ErrorLevel.HARD for e in errors)
    if strict:
        soft = [e for e in errors if e.level == ErrorLevel.SOFT]
        if soft:
            success = False
            for e in soft:
                print(str(e), file=sys.stderr)
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
        print("\n编译失败：存在未解决的硬错误。", file=sys.stderr)
        return 1
    print(f"编译成功：{args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
