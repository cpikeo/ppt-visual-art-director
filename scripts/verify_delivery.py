#!/usr/bin/env python3
"""对照 plan、实际 PPTX 与真实渲染，输出六道质量门的可审查交付报告。

三方核对：plan（意图）↔ PPTX 对象（执行）↔ LibreOffice 渲染（当前环境预览）。
报告区分三类结论：事实（对象核对、渲染测量）、自动诊断（估算与筛查，可按 policy 具名复核）、
设计判断（G-DECK / G-TGT，由人基于渲染证据给出）。任何一类都不冒充另一类。
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import io
import json
import math
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))  # 与 compile_pptx 同目录分发，直接导入

from compile_pptx import (  # noqa: E402  先导入：它对缺失的 python-pptx / jsonschema 给出清晰提示，而不是原始 traceback
    CompileError,
    chart_deemphasis_ink,
    DEFAULT_FONTS,
    ErrorLevel,
    POLICY,
    PLAN_VALIDATOR,
    PLAN_SCHEMA,
    TEXT_ELEMENT_TYPES,
    _compose_chart_caption,
    _contrast_ratio,
    _expected_image_crops,
    _normalize_table_data,
    _NUM_TOKEN_RE,
    _read_unit,
    _rel_luminance,
    _slide_dimensions,
    _table_caption_height,
    _table_geometry,
    PAGE_NUMBER_ID,
    load_plan,
    page_background,
    page_number_config,
    page_number_element,
    page_number_specs,
    resolved_palette,
    run_styles,
    text_style,
    unresolved_reviews,
    validate_plan,
)
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches

# 质量门拓扑、诊断码归属与人工/机器身份全部只从 policy.json 读取。
# 契约文档解释「为什么」，这里是唯一的技术路由表；改契约不改 policy 会被装载检查直接拦住。
_GATE_META = POLICY["quality_gates"]
_VALID_VERIFIERS = {"machine", "human", "declaration"}
if not all(isinstance(meta, dict) and meta.get("verifier") in _VALID_VERIFIERS
           for meta in _GATE_META.values()):
    raise RuntimeError("policy.json 版本不匹配：quality_gates 的值必须是 {label, verifier} 对象"
                       "（verifier ∈ machine/human/declaration）；旧版字符串形状请按 policy 2.0 升级")
if not isinstance(POLICY["diagnostics"].get("gate_of_code"), dict):
    raise RuntimeError("policy.json 版本不匹配：缺少 diagnostics.gate_of_code 路由表")
GATE_IDS = tuple(_GATE_META)
GATE_VERIFIER = {gid: str(meta["verifier"]) for gid, meta in _GATE_META.items()}
HUMAN_GATES = tuple(gid for gid in GATE_IDS if GATE_VERIFIER[gid] == "human")
GATE_OF_CODE = dict(POLICY["diagnostics"]["gate_of_code"])
if set(GATE_OF_CODE.values()) - set(GATE_IDS):
    raise RuntimeError(f"policy.json 诊断码指向未定义的质量门：{sorted(set(GATE_OF_CODE.values()) - set(GATE_IDS))}")
_unrouted = set(POLICY["diagnostics"]["reviewable_codes"]) - set(GATE_OF_CODE)
if _unrouted:
    raise RuntimeError(f"policy.json 可复核诊断码未登记归属门：{sorted(_unrouted)}")
# 每个诊断码只归属一道门：HARD 必然阻断该门；SOFT 未经具名复核时阻断（合规待定为 NOT VERIFIED）；
# INFO 只披露。未登记的诊断码按实现漂移阻断 G-STR——任何 HARD 都不能绕过全部质量门。
EMU_PER_IN = 914400
REPORT_VERSION = "1.6"


def _run_text(command: List[str], timeout: int = 180) -> subprocess.CompletedProcess:
    """子进程文本输出统一按 UTF-8 解码：pdftotext/poppler/LibreOffice 不按 Windows 区域编码出字。

    无法解码的字节以替换字符保留并在证据中显示，绝不让 reader 线程崩溃后把错误落到无关行号。
    """
    try:
        return subprocess.run(command, capture_output=True, encoding="utf-8", errors="replace", timeout=timeout)
    except subprocess.TimeoutExpired:
        # 渲染器卡死时返回非零结果，由调用方转成 G-PRE BLOCK，而不是抛 traceback 中断整个验证。
        return subprocess.CompletedProcess(command, 124, stdout="", stderr=f"命令超时（>{timeout}s）：{command[0]}")
    except OSError as exc:
        return subprocess.CompletedProcess(command, 127, stdout="", stderr=f"无法执行 {command[0]}：{exc}")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _atomic_write_text(path: Path, text: str) -> None:
    """同目录临时文件 + 原子替换；报告写入失败不会截断既有证据。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=path.suffix, dir=str(path.parent))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        temp_path.write_text(text, encoding="utf-8")
        os.replace(temp_path, path)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass


def _error_dict(error: CompileError) -> Dict[str, str]:
    return {
        "code": error.code,
        "level": error.level.value,
        "gate": GATE_OF_CODE.get(error.code, "G-STR"),
        "location": error.location,
        "condition": error.condition,
        "actual": error.actual,
        "message": error.message,
    }


def _gate(status: str, evidence: Optional[List[str]] = None, issues: Optional[List[str]] = None, **extra) -> Dict[str, Any]:
    result = {"status": status, "evidence": evidence or [], "issues": issues or []}
    result.update(extra)
    return result


def _default_config() -> Dict[str, Any]:
    return {
        "required_gates": list(POLICY["release"]["default_required_gates"]),
        "key_pages": [], "high_risk_pages": [], "historical_issue_pages": [],
        "target_powerpoint": {"configured": False},
        "reviews": {}, "risk_approvals": [],
        "sampling": {
            "low_risk_rate": POLICY["sampling"]["default_low_risk_rate"],
            "seed": POLICY["sampling"]["default_seed"],
        },
    }


def _load_config(path: Optional[str]) -> Tuple[Dict[str, Any], Optional[Path], List[str]]:
    if not path:
        return _default_config(), None, []
    config_path = Path(path).resolve()
    try:
        config = json.loads(config_path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        return _default_config(), config_path, [f"验收配置不可读：{exc}"]
    validator = PLAN_VALIDATOR.evolve(schema={"$ref": "#/$defs/acceptanceConfig", "$defs": PLAN_SCHEMA["$defs"]})
    issues = [f"验收配置 {list(e.absolute_path)}：{e.message}" for e in validator.iter_errors(config)]
    if not issues:
        issues.extend(f"未知质量门：{gate}" for gate in sorted(set(config.get("required_gates", [])) - set(GATE_IDS)))
        issues.extend(f"未知复核门：{gate}" for gate in sorted(set(config.get("reviews", {})) - set(GATE_IDS)))
    if issues:
        return _default_config(), config_path, issues
    merged = _default_config()
    sampling = dict(merged["sampling"], **config.get("sampling", {}))
    merged.update(config)
    merged["sampling"] = sampling
    return merged, config_path, []


# ---------------------------------------------------------------------------
# plan ↔ PPTX 对象核对（事实）
# ---------------------------------------------------------------------------

def _shape_map(slide) -> Tuple[Dict[str, Any], List[str]]:
    result: Dict[str, Any] = {}
    issues: List[str] = []
    for shape in slide.shapes:
        if shape.name in result:
            issues.append(f"PPTX 对象名重复：{shape.name}")
        result[shape.name] = shape
    return result, issues


def _geometry_issues(shape, expected: Tuple[float, float, float, float], name: str) -> List[str]:
    actual = (shape.left, shape.top, shape.width, shape.height)
    target = tuple(int(Inches(v)) for v in expected)
    if any(abs(a - b) > 1 for a, b in zip(actual, target)):
        return [f"{name} 几何漂移：plan={target} EMU，pptx={actual} EMU"]
    return []


def _audit_text(shape, el: Dict[str, Any], plan: Dict[str, Any], name: str) -> List[str]:
    """逐段逐片段核对文本对象：期望值全部来自 compile_pptx.text_style/run_styles 的同一解析。"""
    issues: List[str] = []
    if not getattr(shape, "has_text_frame", False):
        return [f"{name} 不是可编辑文本对象"]
    if shape.text_frame.text != el["content"]:
        issues.append(f"{name} 文本漂移：plan={el['content']!r}，pptx={shape.text_frame.text!r}")
    fonts = plan.get("fonts", {})
    style = text_style(el, fonts)
    expected_anchor = {"middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}.get(style["valign"], MSO_ANCHOR.TOP)
    if shape.text_frame.vertical_anchor != expected_anchor:
        issues.append(f"{name} 竖直锚定漂移：plan={style['valign']}，pptx={shape.text_frame.vertical_anchor}")
    run_specs = run_styles(el, fonts) if el.get("runs") else None
    for pi, paragraph in enumerate(shape.text_frame.paragraphs):
        try:
            actual_spacing = float(paragraph.line_spacing)
        except (TypeError, ValueError):
            actual_spacing = None
        if actual_spacing is None or abs(actual_spacing - style["line_spacing"]) > 1e-6:
            issues.append(f"{name} 行距漂移：第 {pi + 1} 段 plan={style['line_spacing']}，pptx={paragraph.line_spacing}")
        if el.get("space_after_pt") is not None:
            if paragraph.space_after is None or abs(paragraph.space_after.pt - float(el["space_after_pt"])) > 1e-6:
                issues.append(f"{name} 段后距漂移：第 {pi + 1} 段")
        if el.get("align"):
            expected_align = {
                "left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY,
            }[el["align"]]
            if paragraph.alignment != expected_align:
                issues.append(f"{name} 第 {pi + 1} 段对齐方式漂移")
        runs = [r for r in paragraph.runs if r.text]
        specs = run_specs if (run_specs and pi == 0) else [dict(style, text=r.text) for r in runs]
        if len(runs) != len(specs):
            issues.append(f"{name} 第 {pi + 1} 段片段数量漂移：plan={len(specs)}，pptx={len(runs)}")
        for run, spec in zip(runs, specs):
            if run.text != spec["text"]:
                issues.append(f"{name} 内联 run 切分漂移：plan={spec['text']!r}，pptx={run.text!r}")
                continue
            label = run.text[:20]
            if run.font.size is None or abs(run.font.size.pt - spec["size"]) > 1e-6:
                issues.append(f"{name} 字号漂移：{label!r}")
            if run.font.name != spec["font"]:
                issues.append(f"{name} 字体漂移：plan={spec['font']!r}，pptx={run.font.name!r}")
            if run.font.color.type != 1 or str(run.font.color.rgb) != spec["color"].lstrip("#"):
                issues.append(f"{name} 文本颜色漂移：{label!r}")
            if bool(run.font.bold) != bool(spec["bold"]):
                issues.append(f"{name} 字重漂移：{label!r}")
            rpr = run._r.rPr
            ea = rpr.find(qn("a:ea")) if rpr is not None else None
            if ea is None or ea.get("typeface") != spec["font"]:
                issues.append(f"{name} 缺少匹配的 <a:ea>：{label!r}")
            actual_spc = rpr.get("spc") if rpr is not None else None
            if spec["letter_spacing_pt"] is None:
                if actual_spc is not None:
                    issues.append(f"{name} 字距漂移（plan 未声明）：{label!r}")
            elif actual_spc is None or int(actual_spc) != int(round(float(spec["letter_spacing_pt"]) * 100)):
                issues.append(f"{name} 字距漂移：{label!r}")
    return issues


def _audit_chart(shape, el: Dict[str, Any], name: str, palette: List[str], bg_hex: str) -> List[str]:
    expected = el["chart"]
    kind = expected["kind"]
    issues: List[str] = []
    if kind == "table":
        if not getattr(shape, "has_table", False):
            return [f"{name} 不是原生可编辑表格"]
        headers, rows = _normalize_table_data(expected)
        table = shape.table
        if len(table.rows) != len(rows) + 1 or len(table.columns) != len(headers):
            return [f"{name} 表格行列数量漂移"]
        actual = [[table.cell(r, c).text for c in range(len(headers))] for r in range(len(rows) + 1)]
        if actual != [headers] + rows:
            issues.append(f"{name} 表格载荷漂移：plan={([headers] + rows)!r}，pptx={actual!r}")
        weights = expected.get("column_weights")
        if weights:
            total = float(sum(weights))
            actual_total = sum(int(col.width) for col in table.columns) or 1
            for ci, (col, weight) in enumerate(zip(table.columns, weights)):
                want, got = float(weight) / total, int(col.width) / actual_total
                if abs(got - want) > 0.02:
                    issues.append(f"{name} 列宽权重漂移：第 {ci + 1} 列 plan={want:.2%}，pptx={got:.2%}")
        return issues
    if not getattr(shape, "has_chart", False):
        return [f"{name} 不是原生可编辑图表"]
    chart = shape.chart
    expected_types = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
        "donut": XL_CHART_TYPE.DOUGHNUT,
    }
    if chart.chart_type != expected_types[kind]:
        issues.append(f"{name} 图表方向/类型漂移：plan={kind}，pptx={chart.chart_type}")
    data = expected["data"]
    actual_categories = [str(c) for c in chart.plots[0].categories]
    expected_categories = [str(c) for c in data["categories"]]
    if actual_categories != expected_categories:
        issues.append(f"{name} 类别顺序漂移：plan={expected_categories}，pptx={actual_categories}")
    if len(chart.series) != len(data["series"]):
        issues.append(f"{name} 系列数量漂移：plan={len(data['series'])}，pptx={len(chart.series)}")
    for index, (expected_series, actual_series) in enumerate(zip(data["series"], chart.series)):
        if actual_series.name != expected_series["name"]:
            issues.append(f"{name} 系列 {index} 名称漂移")
        actual_values = [float(v) for v in actual_series.values]
        expected_values = [float(v) for v in expected_series["values"]]
        if actual_values != expected_values:
            issues.append(f"{name} 系列 {index} 数值/顺序漂移：plan={expected_values}，pptx={actual_values}")
    caption = _compose_chart_caption(expected)
    actual_caption = chart.chart_title.text_frame.text if chart.has_title else ""
    if actual_caption != caption:
        issues.append(f"{name} 标题/单位/期间漂移：plan={caption!r}，pptx={actual_caption!r}")
    highlight = expected.get("highlight_category")
    if kind == "bar" and highlight is not None:
        deemphasis = chart_deemphasis_ink(bg_hex, kind)[0].lstrip("#").upper()
        for index, point in enumerate(chart.plots[0].series[0].points):
            want = palette[0].lstrip("#").upper() if index == highlight else deemphasis
            color = point.format.fill.fore_color
            if color.type != 1 or str(color.rgb) != want:
                issues.append(f"{name} 高亮/退色身份漂移：类别 {index + 1}")
    else:
        targets = list(chart.plots[0].series[0].points) if kind in {"pie", "donut"} else list(chart.series)
        for index, target in enumerate(targets):
            color = target.format.line.color if kind == "line" else target.format.fill.fore_color
            if color.type != 1 or str(color.rgb) != palette[index].lstrip("#").upper():
                issues.append(f"{name} 颜色身份漂移：位置 {index + 1}")
    if kind in {"pie", "donut"}:
        labels = chart.plots[0].data_labels
        if not labels.show_percentage or labels.show_value:
            issues.append(f"{name} 饼/环标签契约漂移")
    else:
        plot = chart.plots[0]
        if bool(expected.get("data_labels")) != bool(plot.has_data_labels):
            issues.append(f"{name} 直标数值声明与实际对象不一致")
        elif expected.get("data_labels") and not plot.data_labels.show_value:
            issues.append(f"{name} 直标未显示数值")
        if expected.get("axis_labels") is False and (chart.value_axis.visible
                                                     or chart.value_axis.tick_label_position != XL_TICK_LABEL_POSITION.NONE):
            issues.append(f"{name} 已声明隐藏数值轴，但轴线或刻度仍会绘制")
    return issues


def _shape_fill_hex(shape) -> Optional[str]:
    spPr = shape._element.find(qn("p:spPr"))
    srgb = spPr.find(qn("a:solidFill") + "/" + qn("a:srgbClr")) if spPr is not None else None
    return ("#" + srgb.get("val").upper()) if srgb is not None else None


def _slide_background_hex(slide) -> Optional[str]:
    """只读解析幻灯片自身的纯色底；继承母版/主题或非纯色时返回 None（不猜）。"""
    srgb = slide._element.find(qn("p:cSld") + "/" + qn("p:bg") + "/" + qn("p:bgPr") + "/" + qn("a:solidFill") + "/" + qn("a:srgbClr"))
    return ("#" + srgb.get("val").upper()) if srgb is not None else None


def audit_pptx(plan: Dict[str, Any], pptx_path: Path, presentation=None) -> Tuple[List[str], Dict[str, Any]]:
    issues: List[str] = []
    evidence: Dict[str, Any] = {"pages_checked": 0, "elements_checked": 0, "derived_objects_checked": 0}
    try:
        prs = presentation if presentation is not None else Presentation(str(pptx_path))
    except Exception as exc:
        return [f"PPTX 无法打开：{type(exc).__name__}: {exc}"], evidence
    if len(prs.slides) != len(plan["pages"]):
        issues.append(f"页数漂移：plan={len(plan['pages'])}，pptx={len(prs.slides)}")
    expected_w, expected_h = _slide_dimensions(plan.get("aspect_ratio"))
    if abs(prs.slide_width - int(Inches(expected_w))) > 1 or abs(prs.slide_height - int(Inches(expected_h))) > 1:
        issues.append("PPTX 画布尺寸漂移")
    palette = resolved_palette(plan)
    for page, slide in zip(plan["pages"], prs.slides):
        pid = page["id"]
        bg_hex = page_background(page)
        if _slide_background_hex(slide) != bg_hex:
            issues.append(f"{pid} 背景色漂移：plan={bg_hex}，pptx={_slide_background_hex(slide)}")
        shapes, shape_issues = _shape_map(slide)
        issues.extend(f"{pid}: {issue}" for issue in shape_issues)
        caption_heights: Dict[str, float] = {}
        expected_order: List[str] = []
        default_z = float(POLICY["defaults"]["element"]["z_index"])
        for el in sorted(page["elements"], key=lambda e: float(e.get("z_index", default_z))):
            name = f"{pid}.{el['id']}"
            if el["type"] == "chart" and el["chart"]["kind"] == "table":
                cap_h = _table_caption_height(el["chart"], el["h"])
                if cap_h:
                    caption_heights[name] = cap_h
                    expected_order.append(f"{name}.caption")
            expected_order.append(name)
        expected_names = set(expected_order)
        # 只有**计划外**且匹配 `<page>.<element>.error[.message]` 的对象才是编译失败占位。
        # 合法元素 id="error" 的对象名是 `<page>.error`，已在 expected_names 中，不能误判。
        diagnostic = [shape.name for shape in slide.shapes
                      if shape.name not in expected_names and re.search(r"\.error(?:\.|$)", shape.name)]
        if diagnostic:
            issues.append(f"{pid} 含编译失败诊断占位对象 {diagnostic}：此 PPTX 仅供诊断，不得作为交付物")
        extras = [shape.name for shape in slide.shapes if shape.name not in expected_names
                  and shape.name not in diagnostic]
        if extras:
            issues.append(f"{pid} 存在计划外对象：{extras}")
        actual_order = [shape.name for shape in slide.shapes if shape.name in expected_names]
        if actual_order != expected_order:
            issues.append(f"{pid} 图层顺序漂移：plan={expected_order}，pptx={actual_order}")
        for el in page["elements"]:
            name = f"{pid}.{el['id']}"
            if name in caption_heights:
                caption_name = f"{name}.caption"
                caption_shape = shapes.get(caption_name)
                if caption_shape is None:
                    issues.append(f"{caption_name} 派生表格标题缺失")
                else:
                    evidence["derived_objects_checked"] += 1
                    issues.extend(_geometry_issues(caption_shape, (el["x"], el["y"], el["w"], caption_heights[name]), caption_name))
                    if not getattr(caption_shape, "has_text_frame", False):
                        issues.append(f"{caption_name} 不是可编辑文本对象")
                    elif caption_shape.text_frame.text != _compose_chart_caption(el["chart"]):
                        issues.append(f"{caption_name} 表格标题/单位/期间漂移")
            shape = shapes.get(name)
            if shape is None:
                issues.append(f"{name} 在实际 PPTX 中缺失")
                continue
            evidence["elements_checked"] += 1
            expected_geometry = (el["x"], el["y"], el["w"], el["h"])
            if el["type"] == "chart" and el["chart"]["kind"] == "table":
                expected_geometry = _table_geometry(el["chart"], *expected_geometry)
            issues.extend(_geometry_issues(shape, expected_geometry, name))
            if el["type"] in TEXT_ELEMENT_TYPES:
                issues.extend(_audit_text(shape, el, plan, name))
            elif el["type"] == "chart":
                issues.extend(_audit_chart(shape, el, name, palette, bg_hex))
            elif el["type"] in {"shape", "divider"}:
                expected_kind = {"rect": MSO_SHAPE.RECTANGLE, "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
                                 "oval": MSO_SHAPE.OVAL}[el.get("shape_kind",
                                                              POLICY["defaults"]["element"]["shape_kind"])]
                if shape.shape_type != 1 or shape.auto_shape_type != expected_kind:
                    issues.append(f"{name} 原生形状类型漂移")
                if _shape_fill_hex(shape) != el["color"].upper():
                    issues.append(f"{name} 形状颜色漂移：plan={el['color']}，pptx={_shape_fill_hex(shape)}")
                srgb = shape._element.find(qn("p:spPr") + "/" + qn("a:solidFill") + "/" + qn("a:srgbClr"))
                alpha_el = srgb.find(qn("a:alpha")) if srgb is not None else None
                want_alpha = el.get("fill_alpha")
                if want_alpha is None:
                    if alpha_el is not None:
                        issues.append(f"{name} 未声明的填充透明度漂移")
                elif alpha_el is None or int(alpha_el.get("val", "0")) != int(round(float(want_alpha) * 100000)):
                    issues.append(f"{name} 填充透明度漂移：plan={want_alpha}")
            elif el["type"] == "image":
                if shape.shape_type != 13:
                    issues.append(f"{name} 不是可编辑图片对象")
                else:
                    src_w, src_h = shape.image.size
                    expected_crop = list(_expected_image_crops(src_w, src_h, el["w"], el["h"], el.get("crop_focus")))
                    actual_crop = [shape.crop_left, shape.crop_right, shape.crop_top, shape.crop_bottom]
                    if any(abs(a - e) > 0.00002 for a, e in zip(actual_crop, expected_crop)):
                        issues.append(f"{name} 图片裁切漂移（应与 crop_focus={el.get('crop_focus') or '居中'} 一致）")
                    # 真实生成素材闭环：登记里的文件哈希必须等于 PPTX 内嵌图片的字节哈希。
                    gen = (plan.get("assets", {}).get(el["image_id"]) or {}).get("image_generation") or {}
                    if gen.get("status") == "generated" and gen.get("sha256"):
                        actual_hash = hashlib.sha256(shape.image.blob).hexdigest()
                        evidence.setdefault("generated_images_checked", []).append(
                            {"object": name, "image_id": el["image_id"], "file": gen.get("file"), "tool": gen.get("tool"),
                             "src_px": [src_w, src_h], "sha256": actual_hash,
                             "expected_sha256": str(gen["sha256"]),
                             "matches": actual_hash.lower() == str(gen["sha256"]).lower()})
        evidence["pages_checked"] += 1
    return issues, evidence


# 注意：两色对比度直接复用 compile_pptx 导入的 _contrast_ratio（同一 WCAG 公式），
# 不再在验证器内另定义一份实现，避免编译与核验各自维护一份亮度/对比逻辑。


def _solid_hex(node, wrap: str = "") -> Optional[str]:
    """读 solidFill 的 srgbClr；wrap="ln" 时读描边填充（折线用线色而非填充色）。"""
    if node is None:
        return None
    parts = (wrap, "solidFill", "srgbClr") if wrap else ("solidFill", "srgbClr")
    found = node.find("/".join(qn(f"a:{part}") for part in parts))
    if found is None:
        return None
    value = found.get("val")
    return f"#{value.upper()}" if value else None


def _series_label(ser) -> str:
    text = ser.find(qn("c:tx") + "/" + qn("c:strRef") + "/" + qn("c:strCache") + "/" + qn("c:pt") + "/" + qn("c:v"))
    if text is None:
        text = ser.find(qn("c:tx") + "/" + qn("c:v"))
    return (text.text or "").strip() if text is not None else ""


def _chart_artifact_inks(chart_obj) -> Tuple[List[Tuple[str, str]], List[str]]:
    """从成品图表 XML 读取实际着色，不重跑编译期的取色逻辑。

    这是核验与生产解耦的关键：如果拿 plan 的 palette 再算一遍编译函数，编译器一旦
    着色出错，核验器会用同一份实现算出同样的结果，永远发现不了。

    点级填充优先于系列级：单系列高亮、饼/环的分类色都落在 dPt；折线取描边色。
    返回 (inks, gaps)，gaps 是成品里读不到显式着色的位置，按「未能核对」披露。
    """
    space = getattr(chart_obj, "_chartSpace", None)
    inks: List[Tuple[str, str]] = []
    gaps: List[str] = []
    if space is None:
        return inks, ["图表 XML 不可读，无法核对实际着色"]
    for index, ser in enumerate(space.findall(".//" + qn("c:ser"))):
        label = _series_label(ser) or f"系列 {index + 1}"
        points: List[Tuple[str, str]] = []
        for point in ser.findall(qn("c:dPt")):
            hex_value = _solid_hex(point.find(qn("c:spPr")), "")
            if hex_value:
                idx = point.find(qn("c:idx"))
                points.append((f"{label}·第 {int(idx.get('val')) + 1 if idx is not None else '?'} 项", hex_value))
        if points:
            inks.extend(points)
            continue
        hex_value = _solid_hex(ser.find(qn("c:spPr")), "") or _solid_hex(ser.find(qn("c:spPr")), "ln")
        if hex_value:
            inks.append((label, hex_value))
        else:
            gaps.append(f"{label}：成品中未找到显式着色（可能回退主题色）")
    return inks, gaps


def _technical_diagnostics(plan: Dict[str, Any], presentation) -> List[CompileError]:
    """在 VERIFY 阶段从实际 PPTX 重算编译期技术诊断（E009/E012）。

    这些指标依赖图片字节与最终落位，不能只存在于某一次编译的内存里：否则
    `compile --strict` 因低 PPI/裁切/图表对比失败，最终报告却看不到同一问题。
    这里以实际内嵌图片尺寸、实际裁切和声明几何重算，不读取编译日志。
    """
    errors: List[CompileError] = []
    image_cfg = POLICY["image"]
    uses_default_palette = "palette" not in plan
    for page, slide in zip(plan["pages"], presentation.slides):
        pid = page["id"]
        shapes, _issues = _shape_map(slide)
        for el in page["elements"]:
            eid = el["id"]
            shape = shapes.get(f"{pid}.{eid}")
            if shape is None:
                continue
            if el["type"] == "image" and getattr(shape, "shape_type", None) == 13:
                src_w, src_h = shape.image.size
                if src_w <= 0 or src_h <= 0:
                    continue
                visible_w = src_w * (1.0 - float(shape.crop_left) - float(shape.crop_right))
                visible_h = src_h * (1.0 - float(shape.crop_top) - float(shape.crop_bottom))
                ppi = min(visible_w / float(el["w"]), visible_h / float(el["h"]))
                loss = 1.0 - (visible_w * visible_h) / (src_w * src_h)
                loc = f"page[{pid}].element[{eid}].image"
                if ppi < float(image_cfg["min_effective_ppi"]):
                    errors.append(CompileError(
                        "E012", ErrorLevel.SOFT, loc,
                        f"有效分辨率 ≥ {image_cfg['min_effective_ppi']} ppi（技术下限）",
                        f"实际约 {ppi:.0f} ppi 于 {float(el['w']):.2f}×{float(el['h']):.2f} in 落位",
                        "投影下将明显发虚：换更高分辨率源图、缩小落位或按最终画幅重出；不评价构图优劣",
                    ))
                if loss > float(image_cfg["max_crop_loss"]):
                    errors.append(CompileError(
                        "E012", ErrorLevel.SOFT, loc,
                        f"裁切损失 ≤ {image_cfg['max_crop_loss']:.0%}（技术下限）",
                        f"实际丢失源图约 {loss:.0%} 面积",
                        "主体大概率被切伤：用 crop_focus 保主体，或直接按最终画幅比例生成/选材；有意取局部细节可具名复核",
                    ))
            elif el["type"] == "chart" and el["chart"]["kind"] != "table":
                loc = f"page[{pid}].element[{eid}].chart"
                if uses_default_palette:
                    errors.append(CompileError(
                        "E009", ErrorLevel.INFO, loc,
                        "缺省色来自唯一技术策略", "未提供 palette；使用 policy.json 默认图表色",
                        "仅披露执行默认，不证明品牌/色彩语义适配；仍须整稿复核",
                    ))
                kind = el["chart"]["kind"]
                floor = float(POLICY["chart"]["min_contrast"].get(kind, 0.0))
                bg_hex = page_background(page)
                if kind == "bar" and el["chart"].get("highlight_category") is not None:
                    # 退隐色由同一函数按当页底色兜底；调整过就写进报告，别只在编译终端里说一次。
                    _deemph_ink, _deemph_note = chart_deemphasis_ink(bg_hex, kind)
                    if _deemph_note:
                        errors.append(CompileError(
                            "E009", ErrorLevel.INFO, loc,
                            "反强调色须在页面底色上可辨", _deemph_note,
                            "系统按 min_contrast 兜底保证可辨；若需完全自定义请显式声明系列色",
                        ))
                inks, gaps = _chart_artifact_inks(getattr(shape, "chart", None))
                for role, hex_value in inks:
                    ratio = _contrast_ratio(hex_value, bg_hex)
                    if ratio < floor:
                        errors.append(CompileError(
                            "E009", ErrorLevel.SOFT, loc,
                            "图表着色须在页面底色上可辨且语义一致",
                            f"成品「{role}」为 {hex_value}：对比度 {ratio:.2f}:1 < 技术复核阈值 "
                            f"{floor:.1f}:1（底色 {bg_hex}）",
                            "已按成品实测，未换色：请核对渲染结果或调整配色；不代表 WCAG 合规",
                        ))
                for gap in gaps:
                    errors.append(CompileError(
                        "E009", ErrorLevel.SOFT, loc,
                        "图表着色须能从成品核对", gap,
                        "核验只报实测结果：着色未落盘或回退主题色时不能推定配色合规",
                    ))
    return errors








# ---------------------------------------------------------------------------
# 文字对比（E013）：在实际 PPTX 的图层上合成文字区域背景，再逐像素计算墨色对比
#
# 回答的唯一问题：每段文字（逐墨色、逐片段字号）与它真正压着的背景之间，对比分布是否达到
# policy 的可读要求。背景按 PPTX 的真实绘制顺序合成：页面底色 → 文字之下的纯色/半透明形状
# （按几何蒙版）→ 图片（嵌入像素 + 实际裁切窗口 + Alpha 通道）。采样区域优先取真实渲染里
# 归属该文字的字形框；没有渲染时退回声明文本框并在记录中注明。
# 还原不了的图层（图表/表格内部、非纯色填充、带效果的图片、主题填充）一律记 UNCERTAIN 交人工，
# 不给确定结论；不透明对象压在文字之上记为遮挡。它不评价配色好坏，也不等于完整无障碍审计。
# ---------------------------------------------------------------------------

_FILL_GEOMS = {"rect", "roundRect", "ellipse"}


def _emu_box(shape) -> Tuple[float, float, float, float]:
    return (shape.left / EMU_PER_IN, shape.top / EMU_PER_IN, shape.width / EMU_PER_IN, shape.height / EMU_PER_IN)


def _layer_of(shape) -> Dict[str, Any]:
    """把一个 PPTX 对象归类为背景合成中的图层：fill | picture | transparent | unknown。"""
    el = shape._element
    tag = el.tag.rsplit("}", 1)[-1]
    layer: Dict[str, Any] = {"name": shape.name, "box": _emu_box(shape)}
    if tag == "pic":
        blip = el.find(".//" + qn("a:blip"))
        effects = [c.tag.rsplit("}", 1)[-1] for c in (blip if blip is not None else []) if c.tag != qn("a:extLst")]
        if effects:
            return dict(layer, kind="unknown", reason=f"图片带效果 {effects}")
        return dict(layer, kind="picture", shape=shape)
    if tag != "sp":
        return dict(layer, kind="unknown", reason="图表/表格/组合对象的内部绘制不可还原" if tag in {"graphicFrame", "grpSp"} else f"对象类型 {tag}")
    spPr = el.find(qn("p:spPr"))
    if spPr is None:
        return dict(layer, kind="unknown", reason="缺少 spPr")
    if spPr.find(qn("a:noFill")) is not None:
        return dict(layer, kind="transparent")
    solid = spPr.find(qn("a:solidFill"))
    if solid is None:
        if any(spPr.find(qn(t)) is not None for t in ("a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill")):
            return dict(layer, kind="unknown", reason="非纯色填充")
        style_fill = el.find(qn("p:style") + "/" + qn("a:fillRef"))
        if style_fill is not None and style_fill.get("idx", "0") != "0":
            return dict(layer, kind="unknown", reason="主题样式填充")
        return dict(layer, kind="transparent")
    srgb = solid.find(qn("a:srgbClr"))
    geom = spPr.find(qn("a:prstGeom"))
    prst = geom.get("prst") if geom is not None else None
    if srgb is None or prst not in _FILL_GEOMS or any(c.tag != qn("a:alpha") for c in srgb):
        return dict(layer, kind="unknown", reason=f"填充色/几何不可还原（{prst}）")
    alpha_el = srgb.find(qn("a:alpha"))
    adj = 16667  # OOXML roundRect 缺省圆角
    gd = geom.find(qn("a:avLst") + "/" + qn("a:gd"))
    if gd is not None and gd.get("fmla", "").startswith("val "):
        adj = int(gd.get("fmla")[4:])
    return dict(layer, kind="fill", rgb="#" + srgb.get("val").upper(), geom=prst, adj=adj,
                alpha=(int(alpha_el.get("val")) / 100000.0) if alpha_el is not None else 1.0)


def _intersects(a: Tuple[float, float, float, float], b: Tuple[float, float, float, float]) -> bool:
    return min(a[2], b[2]) > max(a[0], b[0]) and min(a[3], b[3]) > max(a[1], b[1])


def _ltrb(box: Tuple[float, float, float, float]) -> Tuple[float, float, float, float]:
    x, y, w, h = box
    return (x, y, x + w, y + h)


def _hex_rgb(hex_color: str) -> Tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


class _Compositor:
    """按绘制顺序把文字下方的图层合成为 RGB 位图（单位：英寸 → 像素，密度 = render_dpi）。"""

    def __init__(self, image_module, image_draw, density: float):
        self.Image, self.ImageDraw, self.S = image_module, image_draw, density
        self._pictures: Dict[str, Any] = {}

    def _picture(self, layer: Dict[str, Any]):
        """图片的可见窗口（已执行裁切）缩放到落位尺寸，保留 Alpha；同一对象只算一次。"""
        key = layer["name"]
        if key not in self._pictures:
            pic = layer["shape"]
            with self.Image.open(io.BytesIO(pic.image.blob)) as source:
                src = source.convert("RGBA")
            w, h = src.size
            window = src.crop((round(pic.crop_left * w), round(pic.crop_top * h),
                               round((1 - pic.crop_right) * w), round((1 - pic.crop_bottom) * h)))
            _x, _y, lw, lh = layer["box"]
            self._pictures[key] = window.resize((max(1, round(lw * self.S)), max(1, round(lh * self.S))), self.Image.BILINEAR)
        return self._pictures[key]

    def _mask(self, layer: Dict[str, Any], origin: Tuple[float, float], size: Tuple[int, int], value: int):
        mask = self.Image.new("L", size, 0)
        x, y, w, h = layer["box"]
        rect = [(x - origin[0]) * self.S, (y - origin[1]) * self.S,
                (x + w - origin[0]) * self.S - 1, (y + h - origin[1]) * self.S - 1]
        draw = self.ImageDraw.Draw(mask)
        if layer["geom"] == "ellipse":
            draw.ellipse(rect, fill=value)
        elif layer["geom"] == "roundRect":
            draw.rounded_rectangle(rect, radius=layer["adj"] / 100000.0 * min(w, h) * self.S, fill=value)
        else:
            draw.rectangle(rect, fill=value)
        return mask

    def compose(self, bbox: Tuple[float, float, float, float], page_bg: str, layers: List[Dict[str, Any]]):
        x0, y0, x1, y1 = bbox
        size = (max(1, round((x1 - x0) * self.S)), max(1, round((y1 - y0) * self.S)))
        canvas = self.Image.new("RGBA", size, _hex_rgb(page_bg) + (255,))
        for layer in layers:
            if not _intersects(_ltrb(layer["box"]), bbox):
                continue
            if layer["kind"] == "fill":
                tint = self.Image.new("RGBA", size, _hex_rgb(layer["rgb"]) + (255,))
                canvas = self.Image.composite(tint, canvas, self._mask(layer, (x0, y0), size, round(255 * layer["alpha"])))
            elif layer["kind"] == "picture":
                overlay = self.Image.new("RGBA", size, (0, 0, 0, 0))
                lx, ly, _w, _h = layer["box"]
                overlay.paste(self._picture(layer), (round((lx - x0) * self.S), round((ly - y0) * self.S)))
                canvas = self.Image.alpha_composite(canvas, overlay)
        return canvas.convert("RGB")

    def coverage(self, bbox: Tuple[float, float, float, float], rects: List[Tuple[float, float, float, float]],
                 layer: Dict[str, Any]) -> float:
        """上层对象实际盖住采样区域的比例（图片按 Alpha，形状按几何蒙版与不透明度）。"""
        x0, y0, x1, y1 = bbox
        size = (max(1, round((x1 - x0) * self.S)), max(1, round((y1 - y0) * self.S)))
        if layer["kind"] == "picture":
            cover = self.Image.new("L", size, 0)
            lx, ly, _w, _h = layer["box"]
            cover.paste(self._picture(layer).getchannel("A"), (round((lx - x0) * self.S), round((ly - y0) * self.S)))
        else:
            cover = self._mask(layer, (x0, y0), size, 255)
        covered = total = 0
        for rx0, ry0, rx1, ry1 in rects:
            crop = cover.crop((round((rx0 - x0) * self.S), round((ry0 - y0) * self.S),
                               max(round((rx0 - x0) * self.S) + 1, round((rx1 - x0) * self.S)),
                               max(round((ry0 - y0) * self.S) + 1, round((ry1 - y0) * self.S))))
            hist = crop.histogram()
            total += sum(hist)
            covered += sum(hist[128:])
        return covered / total if total else 0.0


def _contrast_quantiles(counts: Dict[Tuple[int, int, int], int], ink_hex: str,
                        quantiles: Tuple[float, ...]) -> Tuple[float, ...]:
    """按像素计数加权的对比分布：返回各分位（0 = 最低、1 = 最高）。对每个像素算对比，深墨/浅墨方向一视同仁。"""
    ink = _rel_luminance(ink_hex)
    pairs = []
    for (r, g, b), n in counts.items():
        lum = _rel_luminance("#%02X%02X%02X" % (r, g, b))
        hi, lo = max(ink, lum), min(ink, lum)
        pairs.append(((hi + 0.05) / (lo + 0.05), n))
    pairs.sort()
    total = sum(n for _c, n in pairs)

    def _q(q: float) -> float:
        target, acc = q * total, 0
        for contrast, n in pairs:
            acc += n
            if acc >= target:
                return contrast
        return pairs[-1][0]

    return tuple(_q(q) for q in quantiles)


def _contrast_verdict(median: float, worst: float, need_median: float, need_worst: float,
                      floor: Dict[str, Any]) -> str:
    """BLOCK（跌破硬下限）< REVIEW（低于要求）< PASS；对中位与最差分位都单调。"""
    if median < float(floor["median"]) or worst < float(floor["worst"]):
        return "BLOCK"
    return "REVIEW" if median < need_median or worst < need_worst else "PASS"


def _run_word_rects(content: str, specs: List[Dict[str, Any]],
                    words: List[Dict[str, Any]]) -> Optional[Tuple[Dict[str, List[Tuple[float, float, float, float]]],
                                                                    Dict[str, float]]]:
    """把 PDF 词框按字符位置映射到具体 run，返回 {ink: rects} 与 {ink: 最小字号}。

    pdftotext 只给词框，不给 run 身份。按 content 的字符顺序把词框对回 run 范围；
    一个词跨两个 run 时同时计入两种墨色（保守）。任一词无法对齐时返回 None，
    调用方退回整框估算并记 NOT VERIFIED，而不是拿混合背景冒充某个 run 的实际背景。
    """
    if not words:
        return None
    if len(specs) == 1:
        rects = [(w["x0"] / 72.0, w["y0"] / 72.0, w["x1"] / 72.0, w["y1"] / 72.0) for w in words]
        return ({specs[0]["color"]: rects}, {specs[0]["color"]: float(specs[0]["size"])})
    if "".join(str(s["text"]) for s in specs) != content:
        return None
    compact = [(ch, i) for i, ch in enumerate(content) if not ch.isspace()]
    compact_text = "".join(ch for ch, _i in compact)
    cursor = 0
    mapped: List[Tuple[int, int, Dict[str, Any]]] = []
    for word in words:
        token = "".join(str(word.get("text") or "").split())
        if not token:
            continue
        if compact_text[cursor:cursor + len(token)] != token:
            found = compact_text.find(token, cursor, cursor + max(64, len(token) * 4))
            if found < 0:
                return None
            cursor = found
        start = compact[cursor][1]
        end = compact[cursor + len(token) - 1][1] + 1
        mapped.append((start, end, word))
        cursor += len(token)
    if not mapped:
        return None
    ranges: List[Tuple[int, int, str, float]] = []
    offset = 0
    for spec in specs:
        text = str(spec["text"])
        ranges.append((offset, offset + len(text), str(spec["color"]), float(spec["size"])))
        offset += len(text)
    ink_rects: Dict[str, List[Tuple[float, float, float, float]]] = {}
    ink_sizes: Dict[str, float] = {}
    for run_start, run_end, ink, size in ranges:
        rects = ink_rects.setdefault(ink, [])
        for start, end, word in mapped:
            if end <= run_start or start >= run_end:
                continue
            rects.append((word["x0"] / 72.0, word["y0"] / 72.0,
                          word["x1"] / 72.0, word["y1"] / 72.0))
        if rects and (ink not in ink_sizes or size < ink_sizes[ink]):
            ink_sizes[ink] = size
    ink_rects = {ink: rects for ink, rects in ink_rects.items() if rects}
    return (ink_rects, ink_sizes) if ink_rects else None


def _text_contrast(plan: Dict[str, Any], pptx_path: Path,
                   rendered_pages: Optional[List[Dict[str, Any]]] = None,
                   confirmed: Optional[Set[int]] = None, presentation=None) -> Tuple[List[CompileError], List[Dict[str, Any]]]:
    cfg = POLICY["text_contrast"]
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return [CompileError("E013", ErrorLevel.SOFT, "deck.text_contrast", "文字对比可计算",
                             "缺少 Pillow，未计算", "NOT VERIFIED：安装 Pillow 后重跑，不得推定可读")], []
    prs = presentation if presentation is not None else Presentation(str(pptx_path))
    compositor = _Compositor(Image, ImageDraw, float(POLICY["preview"]["render_dpi"]))
    floor, worst_q = cfg["hard_floor"], float(cfg["worst_quantile"])
    errors: List[CompileError] = []
    records: List[Dict[str, Any]] = []
    fonts = plan.get("fonts", {})
    for pi, (page, slide) in enumerate(zip(plan["pages"], prs.slides)):
        pid = page["id"]
        stack = [_layer_of(shape) for shape in slide.shapes]
        index_of = {layer["name"]: i for i, layer in enumerate(stack)}
        page_bg = _slide_background_hex(slide)
        owned = (_page_word_owners(page, rendered_pages[pi]["words"])[0]
                 if rendered_pages and pi < len(rendered_pages) and (confirmed is None or (pi + 1) in confirmed) else {})
        for el in page["elements"]:
            if el["type"] not in TEXT_ELEMENT_TYPES or not el["content"].strip():
                continue
            name = f"{pid}.{el['id']}"
            position = index_of.get(name)
            if position is None:
                continue  # 对象缺失已由 audit_pptx 记为 G-SEM 事实
            words = owned.get(el["id"]) or []
            specs = run_styles(el, fonts)
            mapped = _run_word_rects(el["content"], specs, words) if words else None
            if mapped:
                ink_rects, ink_sizes = mapped
                all_rects = [(w["x0"] / 72.0, w["y0"] / 72.0, w["x1"] / 72.0, w["y1"] / 72.0) for w in words]
                region = "rendered_glyphs"
            else:
                all_rects = [_ltrb((el["x"], el["y"], el["w"], el["h"]))]
                region = "declared_box"
                ink_sizes = {}
                for style in specs:
                    if style["text"].strip():
                        ink_sizes[style["color"]] = min(
                            ink_sizes.get(style["color"], float("inf")), float(style["size"]))
                ink_rects = {ink: all_rects for ink in ink_sizes}
            bbox = (min(r[0] for r in all_rects), min(r[1] for r in all_rects),
                    max(r[2] for r in all_rects), max(r[3] for r in all_rects))
            below = [layer for layer in stack[:position] if _intersects(_ltrb(layer["box"]), bbox)]
            above = [layer for layer in stack[position + 1:] if _intersects(_ltrb(layer["box"]), bbox)]
            uncertain = [f"{layer['name']}：{layer['reason']}" for layer in below if layer["kind"] == "unknown"]
            uncertain += [f"{layer['name']}（上层）：{layer['reason']}" for layer in above if layer["kind"] == "unknown"]
            uncertain += [f"{layer['name']}（上层）：半透明填充改变了文字本身的颜色"
                          for layer in above if layer["kind"] == "fill" and layer["alpha"] < 0.999]
            if page_bg is None:
                uncertain.append("页面底色继承母版/主题或非纯色，未解析")
            occluders = [layer for layer in above if layer["kind"] == "picture"
                         or (layer["kind"] == "fill" and layer["alpha"] >= 0.999)]
            covered = {layer["name"]: round(compositor.coverage(bbox, all_rects, layer), 3) for layer in occluders}
            covered = {k: v for k, v in covered.items() if v > 0}

            location = f"page[{pid}].element({el['id']}).contrast"
            layer_names = [layer["name"] for layer in below if layer["kind"] in {"fill", "picture"}]
            for ink, size in sorted(ink_sizes.items()):
                rects = ink_rects.get(ink) or all_rects
                ink_bbox = (min(r[0] for r in rects), min(r[1] for r in rects),
                            max(r[2] for r in rects), max(r[3] for r in rects))
                composite = compositor.compose(ink_bbox, page_bg or POLICY["defaults"]["background"], below)
                counts: Dict[Tuple[int, int, int], int] = {}
                for rx0, ry0, rx1, ry1 in rects:
                    box = (round((rx0 - ink_bbox[0]) * compositor.S), round((ry0 - ink_bbox[1]) * compositor.S),
                           max(round((rx0 - ink_bbox[0]) * compositor.S) + 1, round((rx1 - ink_bbox[0]) * compositor.S)),
                           max(round((ry0 - ink_bbox[1]) * compositor.S) + 1, round((ry1 - ink_bbox[1]) * compositor.S)))
                    crop = composite.crop(box)
                    for n, rgb in crop.getcolors(crop.width * crop.height):
                        counts[rgb] = counts.get(rgb, 0) + n
                median, worst, lowest, highest = _contrast_quantiles(counts, ink, (0.5, worst_q, 0.0, 1.0))
                need_median = float(cfg["large_text"] if size >= float(cfg["large_text_pt"]) else cfg["normal_text"])
                need_worst = float(cfg["large_text"])
                outcome = _contrast_verdict(median, worst, need_median, need_worst, floor)
                # 结论成立的前提：采样的就是字形处；只有声明框时，框内最低与最高对比必须给出同一结论，
                # 即无论字形落在框内哪里结论都不变（纯色底恒成立）。否则框内统计只是估算，不代表字形处。
                placement_free = region == "rendered_glyphs" or (
                    _contrast_verdict(lowest, lowest, need_median, need_worst, floor)
                    == _contrast_verdict(highest, highest, need_median, need_worst, floor))
                record = {
                    "page": pid, "text": el["id"], "ink": ink, "size_pt": size, "region": region,
                    "background_layers": layer_names or ["page_background"],
                    "contrast_median": round(median, 2), f"contrast_p{round(worst_q * 100)}": round(worst, 2),
                    "required_median": need_median, "required_worst": need_worst,
                    "sample_px": sum(counts.values()),
                }
                if region == "declared_box":
                    record.update(contrast_min=round(lowest, 2), contrast_max=round(highest, 2))
                detail = (f"[{pid}.{el['id']}] 墨色 {ink}（{size:g}pt）对比：中位 {median:.2f}:1（需 ≥{need_median:.1f}）、"
                          f"最差 {worst_q:.0%} 分位 {worst:.2f}:1（需 ≥{need_worst:.1f}）；背景={record['background_layers']}；"
                          f"采样={'渲染字形框' if region == 'rendered_glyphs' else '声明文本框（未定位字形）'}")
                if not placement_free:
                    detail += f"；框内对比 {lowest:.2f}–{highest:.2f}:1 跨越判定线，结论取决于字形落点"
                if covered:
                    certain = region == "rendered_glyphs" or max(covered.values()) >= 0.999
                    record.update(status="BLOCK" if certain else "REVIEW", occluded_by=covered)
                    errors.append(CompileError(
                        "E013", ErrorLevel.HARD if certain else ErrorLevel.SOFT, location,
                        "文字不得被上层不透明对象遮挡",
                        f"[{pid}.{el['id']}] 被上层对象遮挡（覆盖比例）：{covered}"
                        + ("" if certain else "；按声明框估算，未经渲染无法确认字形是否被盖住"),
                        "调整叠放顺序或几何；遮挡下的文字无对比可言",
                    ))
                elif uncertain:
                    record.update(status="UNCERTAIN", uncertain_layers=uncertain)
                    errors.append(CompileError(
                        "E013", ErrorLevel.SOFT, location, "文字背景的图层须可还原后才能判定对比",
                        detail + f"；无法还原：{uncertain}（上述数值只基于可还原图层）",
                        "看渲染图人工确认可读性并具名复核；或改用可还原的纯色/图片承载文字",
                    ))
                elif outcome == "BLOCK":
                    record["status"] = "BLOCK" if placement_free else "REVIEW"
                    errors.append(CompileError(
                        "E013", ErrorLevel.HARD if placement_free else ErrorLevel.SOFT, location,
                        f"文字对比不得跌破硬下限（中位 {floor['median']}:1 / 最差分位 {floor['worst']}:1）",
                        detail + ("" if placement_free else "；需真实渲染确认字形是否落在低对比区"),
                        "不可读，不可豁免：换墨色、移字到安静区、加有依据的承托层或换图",
                    ))
                elif outcome == "REVIEW":
                    record["status"] = "REVIEW"
                    errors.append(CompileError(
                        "E013", ErrorLevel.SOFT, location, "文字对比应达到可读要求", detail,
                        "加深/提亮墨色、移字或调整承托层；确需保留须具名复核并写明目标投影环境实测",
                    ))
                elif placement_free:
                    record["status"] = "PASS"
                else:
                    # 框内统计达标，但字形可能恰好落在框内低对比处：不能据此认定可读，交给真实渲染。
                    record.update(status="NOT VERIFIED",
                                  reason="未定位真实字形且框内背景不均匀；用 --render-dir 从本次 PPTX 渲染后按字形框重判")
                records.append(record)
    return errors, records


# ---------------------------------------------------------------------------
# 真实渲染（G-PRE）：LibreOffice → PDF → PNG / 文本层 / 字体表
# ---------------------------------------------------------------------------

def _pdftotext_bbox(pdf: Path, executable: Optional[str] = None, raw_path: Optional[Path] = None) -> List[Dict[str, Any]]:
    """读取真实 PDF 文本层；解析失败必须向上失败，不能当作没有溢出。"""
    executable = executable or shutil.which("pdftotext")
    if not executable:
        raise RuntimeError("缺少 pdftotext")
    proc = _run_text([executable, "-bbox", str(pdf), "-"], timeout=180)
    if proc.returncode != 0:
        raise RuntimeError(f"pdftotext -bbox 失败：{proc.stderr or proc.stdout}")
    raw = proc.stdout
    if raw_path:
        raw_path.write_text(raw, encoding="utf-8")
    root = ET.fromstring(re.sub(r"<!DOCTYPE[^>]*>", "", raw).strip())
    pages: List[Dict[str, Any]] = []
    for page in root.iter("{http://www.w3.org/1999/xhtml}page"):
        words = [{
            "x0": float(word.get("xMin")), "y0": float(word.get("yMin")),
            "x1": float(word.get("xMax")), "y1": float(word.get("yMax")),
            "text": word.text or "",
        } for word in page.iter("{http://www.w3.org/1999/xhtml}word")]
        pages.append({"width": float(page.get("width")), "height": float(page.get("height")), "words": words})
    return pages


# 一个 PDF 词框只能归属到唯一一个声明文本框：按"词框面积落在该框内的比例"取最大且过半者。
# 否则大字会同时落进相邻小框，伪造小框的断行或溢出。
_WORD_OWNER_MIN_RATIO = 0.5


def _chart_regions(page: Dict[str, Any]) -> List[Tuple[float, float, float, float]]:
    """图表/表格的声明几何（pt）。其内部文字由 LibreOffice 自行绘制，词框不属于任何声明文本框。"""
    return [(float(el["x"]) * 72.0, float(el["y"]) * 72.0,
             float(el["x"] + el["w"]) * 72.0, float(el["y"] + el["h"]) * 72.0)
            for el in page.get("elements", []) if el.get("type") == "chart"]


# 注：矩形相交复用下方 _intersects（与 _boxes_intersect 原实现一致），不重复定义。


def _inside_regions(word: Dict[str, Any], regions: List[Tuple[float, float, float, float]]) -> bool:
    cx, cy = (word["x0"] + word["x1"]) / 2.0, (word["y0"] + word["y1"]) / 2.0
    return any(x0 <= cx <= x1 and y0 <= cy <= y1 for x0, y0, x1, y1 in regions)


def _ambiguous_text(plan: Dict[str, Any]) -> Dict[int, List[Tuple[str, List[str]]]]:
    """与图表/表格叠放的声明文本框：词框归属不可靠，断行与溢出检查应降级为人工复核。

    返回 {页下标: [(元素 id, 相交的图表元素 id 列表)]}。这里只按声明几何判断，
    不做像素猜测；被跳过的对象会在报告的 G-PRE.limitations 里逐项点名。
    """
    result: Dict[int, List[Tuple[str, List[str]]]] = {}
    for pi, page in enumerate(plan["pages"]):
        charts = [(el["id"], _chart_regions({"elements": [el]})[0])
                  for el in page["elements"] if el["type"] == "chart"]
        if not charts:
            continue
        for el in page["elements"]:
            if el["type"] not in TEXT_ELEMENT_TYPES:
                continue
            box = (float(el["x"]) * 72.0, float(el["y"]) * 72.0,
                   float(el["x"] + el["w"]) * 72.0, float(el["y"] + el["h"]) * 72.0)
            hits = [cid for cid, region in charts if _intersects(box, region)]
            if hits:
                result.setdefault(pi, []).append((el["id"], hits))
    return result


def _page_word_owners(plan_page: Dict[str, Any], words: List[Dict[str, Any]]):
    """把本页 PDF 词框分配给至多一个声明文本框；返回 (owned, unowned, boxes)。

    owned[eid] 是归属该元素的词框；unowned 是没被任何文本框认领的词（多为图表/
    表格内部文字等，它们不该被算进任何文字元素）。boxes[eid] 为元素框的 pt 坐标。
    归属用"词框面积落在该框内的比例"，而非简单的框相交：后者会让一个词同时属于
    多个框。比例相同时归更小的框，避免宽题框吞掉坐在其中的小标签。
    中心落在图表/表格声明几何内的词是图表内部文字（坐标轴、直标、单元格），
    直接丢弃：它们既不该被认领成正文，也不该被 _spill_words 当作溢出回收。
    """
    boxes = {el["id"]: (el["x"] * 72.0, el["y"] * 72.0,
                        (el["x"] + el["w"]) * 72.0, (el["y"] + el["h"]) * 72.0)
             for el in plan_page["elements"] if el["type"] in TEXT_ELEMENT_TYPES}
    chart_regions = _chart_regions(plan_page)
    owned: Dict[str, List[Dict[str, Any]]] = {eid: [] for eid in boxes}
    unowned: List[Dict[str, Any]] = []
    for w in words:
        if _inside_regions(w, chart_regions):
            continue
        warea = max(1e-6, (w["x1"] - w["x0"]) * (w["y1"] - w["y0"]))
        best_id, best_ratio, best_box_area = None, 0.0, float("inf")
        for eid, (bx0, by0, bx1, by1) in boxes.items():
            ox = max(0.0, min(w["x1"], bx1) - max(w["x0"], bx0))
            oy = max(0.0, min(w["y1"], by1) - max(w["y0"], by0))
            ratio = (ox * oy) / warea
            if ratio <= 0.0:
                continue
            box_area = (bx1 - bx0) * (by1 - by0)
            if ratio > best_ratio + 1e-9 or (abs(ratio - best_ratio) <= 1e-9 and box_area < best_box_area):
                best_id, best_ratio, best_box_area = eid, ratio, box_area
        if best_id is not None and best_ratio >= _WORD_OWNER_MIN_RATIO:
            owned[best_id].append(w)
        else:
            unowned.append(w)
    return owned, unowned, boxes


def _spill_words(box, owned_words, unowned_words, font_pt, tolerance):
    """回收"从该框溢出到框外、因中心出框而无人认领"的后续行。

    判据刻意收紧：与框内最后一行行距连续（间隙 < 0.75 行高）、方向向下、且水平
    起点落在本框跨度内。这样图表/表格内部文字、远处标题和相邻元素都不会被误算
    成本框溢出，而真正"首行还在框内、后续行落到框下"的溢出仍被抓住（避免漏报）。
    """
    bx0, by0, bx1, _by1 = box
    line_h = (max((w["y1"] - w["y0"] for w in owned_words), default=0.0)
              or max(12.0, font_pt * 1.2))
    max_gap = 0.75 * line_h
    frontier = max((w["y1"] for w in owned_words), default=by0)
    remaining = sorted(unowned_words, key=lambda w: (w["y0"], w["x0"]))
    taken = set()
    spill: List[Dict[str, Any]] = []
    changed = True
    while changed:
        changed = False
        for i, w in enumerate(remaining):
            if i in taken or w["y0"] < by0 - tolerance or w["y0"] > frontier + max_gap:
                continue
            if w["x0"] < bx0 - tolerance or w["x0"] > bx1:
                continue
            taken.add(i)
            spill.append(w)
            frontier = max(frontier, w["y1"])
            changed = True
    # 同一未归属词框只能被一个声明框回收：就地从本页 unowned 中移除，
    # 避免相邻两个框都把同一行算成自己的溢出（重复报错/墨迹归属不清）。
    if spill:
        claimed = {id(w) for w in spill}
        unowned_words[:] = [w for w in unowned_words if id(w) not in claimed]
    return spill


def _visual_lines(words: List[Dict[str, Any]]) -> List[List[Dict[str, Any]]]:
    """把词框聚成视觉行，行内按 x 排序。

    按"词框垂直中心落入已有行带"聚行，而不是比较顶边：中西混排时不同字体的字形框顶边
    可相差数 pt（如 Noto CJK 与 DejaVu 的数字），按顶边聚行会把同一行拆成两行。
    """
    rows: List[Dict[str, Any]] = []
    for w in sorted(words, key=lambda x: ((x["y0"] + x["y1"]) / 2, x["x0"])):
        center = (w["y0"] + w["y1"]) / 2
        row = next((r for r in rows if r["y0"] <= center <= r["y1"]), None)
        if row is None:
            rows.append({"y0": w["y0"], "y1": w["y1"], "words": [w]})
        else:
            row["words"].append(w)
            row["y0"], row["y1"] = min(row["y0"], w["y0"]), max(row["y1"], w["y1"])
    return [sorted(r["words"], key=lambda x: x["x0"]) for r in sorted(rows, key=lambda r: r["y0"])]


# 中文避头尾（GB/T 15834、W3C clreq 的常见禁则子集；客观事实，单源常量）：
# 收尾类标点与附着于数字的符号不得落在行首；开引号/左括号不得落在行尾。
CJK_FORBIDDEN_LINE_START = set("，。、；：？！）》」』】〕〗〙〛〉”’…‥·%％‰℃,.;:!?)]}")
CJK_FORBIDDEN_LINE_END = set("（《「『【〔〖〘〚〈“‘([{")
_NUMBER_SUFFIX_CHARS = set("年月日号时分秒")


def _explicit_breaks(content: str, lines: List[str]) -> List[bool]:
    """第 i 个视觉换行是否正好落在 content 的显式换行处（作者有意分行），否则为自动折行。"""
    ends, acc = set(), 0
    for paragraph in content.split("\n"):
        acc += len(re.sub(r"\s+", "", paragraph))
        ends.add(acc)
    flags, acc = [], 0
    for line in lines[:-1]:
        acc += len(re.sub(r"\s+", "", line))
        flags.append(acc in ends)
    return flags


def _number_split(prev: str, nxt: str) -> Optional[str]:
    """数字、数字+单位（含单位内部，如"万人｜次"）、符号/货币+数字被折到两行时返回说明。

    把相邻两行拼起来，找跨越折点的数值记号及紧随的单位；单位识别复用 compile_pptx 的单一单位表。
    """
    if not prev or not nxt:
        return None
    joined, cut = prev + nxt, len(prev)
    for m in _NUM_TOKEN_RE.finditer(joined):
        if m.start() >= cut:
            break
        if cut < m.end():
            return "数字被拆行"
        unit, end = _read_unit(joined, m.end())
        if unit and cut < end:
            return f"数字与单位「{unit}」被拆开"
        if cut == m.end() and nxt[0] in _NUMBER_SUFFIX_CHARS:
            return f"数字与「{nxt[0]}」被拆开"
    if prev[-1] in "+-−±¥$€£" and nxt[0].isdigit():
        return "符号/货币与数字被拆开"
    return None


def _line_break_issues(plan: Dict[str, Any], pages: List[Dict[str, Any]],
                       confirmed: Optional[Set[int]] = None,
                       skip: Optional[Dict[int, Set[str]]] = None) -> List[str]:
    """逐个声明文本框核对真实渲染的分行：避头尾、孤字，以及自动折行造成的数字/单位拆分。

    按元素而不是按整页聚行：分栏同基线的两段文字不会被拼成一行而互相掩盖。作者用显式
    换行刻意把巨数与单位分成两行时，不算折行错误。机器只报事实，不评判排版审美。
    页序未经文字核对确认的页（confirmed 不含该页号）不参与——词框可能属于别页。
    """
    issues: List[str] = []
    for pi, (page, page_box) in enumerate(zip(plan["pages"], pages)):
        if confirmed is not None and (pi + 1) not in confirmed:
            continue
        owned, _unowned, _boxes = _page_word_owners(page, page_box["words"])
        skipped = (skip or {}).get(pi, set())
        for el in page["elements"]:
            if el["type"] not in TEXT_ELEMENT_TYPES or el["id"] in skipped:
                continue
            lines = ["".join(w["text"] for w in line) for line in _visual_lines(owned[el["id"]])]
            explicit = _explicit_breaks(el["content"], lines)
            where = f"[{page['id']}.{el['id']}]"
            paragraph_starts = [0] + [i + 1 for i, flag in enumerate(explicit) if flag]
            for start, end in zip(paragraph_starts, paragraph_starts[1:] + [len(lines)]):
                last = re.sub(r"[\s" + re.escape("".join(CJK_FORBIDDEN_LINE_START | CJK_FORBIDDEN_LINE_END)) + r"]", "", lines[end - 1]) if end > start else ""
                if end - start >= 2 and len(last) == 1:
                    issues.append(f"{where} 第 {end} 行为孤字「{lines[end - 1]}」：段落末行只剩一个字")
            for i, text in enumerate(lines):
                if not text:
                    continue
                if text[0] in CJK_FORBIDDEN_LINE_START:
                    issues.append(f"{where} 第 {i + 1} 行行首为禁则字符「{text[0]}」：{text[:24]}")
                if text[-1] in CJK_FORBIDDEN_LINE_END:
                    issues.append(f"{where} 第 {i + 1} 行行尾为禁则字符「{text[-1]}」：{text[-24:]}")
                if i + 1 < len(lines) and not explicit[i]:
                    split = _number_split(text, lines[i + 1])
                    if split:
                        issues.append(f"{where} 第 {i + 1}/{i + 2} 行之间{split}：…{text[-8:]}｜{lines[i + 1][:8]}…")
    return issues


def _overflow_violations(plan: Dict[str, Any], pages: List[Dict[str, Any]],
                         confirmed: Optional[Set[int]] = None,
                         skip: Optional[Dict[int, Set[str]]] = None) -> Tuple[List[Dict[str, Any]], Dict[Tuple[str, str], Optional[float]]]:
    """逐声明文本框比较渲染词框与框的边界；返回 (越界候选, 词框越界量 pt，None=未找到渲染文本)。

    词框先按 _page_word_owners 归属到唯一文本框；溢出到框外的后续行由 _spill_words 回收。
    词框是行框与全角字面而非笔画，候选是否真有墨迹出框由 _ink_outside 用渲染像素确认。
    页序未经文字核对确认的页不参与核验（其词框可能属于别页）。
    """
    candidates: List[Dict[str, Any]] = []
    measured: Dict[Tuple[str, str], Optional[float]] = {}
    tolerance = float(POLICY["preview"]["overflow_tolerance_pt"])
    for pi, (page, page_box) in enumerate(zip(plan["pages"], pages)):
        if confirmed is not None and (pi + 1) not in confirmed:
            continue
        owned, unowned, boxes = _page_word_owners(page, page_box["words"])
        skipped = (skip or {}).get(pi, set())
        for el in page["elements"]:
            if el["type"] not in TEXT_ELEMENT_TYPES or not el["content"].strip() or el["id"] in skipped:
                continue
            box = boxes[el["id"]]
            hits = list(owned[el["id"]])
            hits.extend(_spill_words(box, hits, unowned, text_style(el)["size"], tolerance))
            key, where = (page["id"], el["id"]), f"[{page['id']}.{el['id']}]"
            if not hits:
                measured[key] = None
                candidates.append({"page_index": pi, "key": key, "box": box, "hits": [],
                                   "detail": f"{where} 未在声明框附近找到渲染文本，可能完全溢出或丢失"})
                continue
            sides = {"下": max(w["y1"] for w in hits) - box[3], "右": max(w["x1"] for w in hits) - box[2],
                     "上": box[1] - min(w["y0"] for w in hits), "左": box[0] - min(w["x0"] for w in hits)}
            worst = max(sides.values())
            measured[key] = worst
            if worst > tolerance:
                candidates.append({"page_index": pi, "key": key, "box": box, "hits": hits,
                                   "detail": f"{where} 词框越界 {worst:.1f}pt（" + " / ".join(f"{k} {v:.1f}" for k, v in sides.items()) + "）"})
    return candidates, measured


def _ink_outside(slide, png, box: Tuple[float, float, float, float], hits: List[Dict[str, Any]],
                 compositor: "_Compositor") -> Optional[int]:
    """越界词框中、声明框（含容差）之外真正有笔画的像素数；底板含不可还原图层时返回 None（保守）。

    行末悬挂标点的右半空白、巨号数字的降部空间都会让词框越出文本框却没有任何可见笔画。
    这里把渲染 PNG 与"去掉全部文字的图层合成底板"逐像素比较，只认真实笔画。
    """
    tol = float(POLICY["preview"]["overflow_tolerance_pt"])
    bx0, by0, bx1, by1 = box[0] - tol, box[1] - tol, box[2] + tol, box[3] + tol
    rects = []
    for w in hits:
        for r in ((max(w["x0"], bx1), w["y0"], w["x1"], w["y1"]), (w["x0"], w["y0"], min(w["x1"], bx0), w["y1"]),
                  (w["x0"], max(w["y0"], by1), w["x1"], w["y1"]), (w["x0"], w["y0"], w["x1"], min(w["y1"], by0))):
            if r[2] - r[0] > 0.5 and r[3] - r[1] > 0.5:
                rects.append(tuple(v / 72.0 for v in r))
    if not rects:
        return 0
    layers = [_layer_of(shape) for shape in slide.shapes]
    page_bg = _slide_background_hex(slide)
    if page_bg is None or any(layer["kind"] == "unknown" and _intersects(_ltrb(layer["box"]), r) for layer in layers for r in rects):
        return None
    plate_layers = [layer for layer in layers if layer["kind"] in {"fill", "picture"}]
    delta, density, ink = int(POLICY["preview"]["overflow_ink_delta"]), compositor.S, 0
    for r in rects:
        plate = compositor.compose(r, page_bg, plate_layers)
        left, top = round(r[0] * density), round(r[1] * density)
        real = png.crop((left, top, left + plate.width, top + plate.height)).tobytes()
        base = plate.tobytes()
        ink += sum(1 for i in range(0, len(base), 3)
                   if max(abs(base[i] - real[i]), abs(base[i + 1] - real[i + 1]), abs(base[i + 2] - real[i + 2])) > delta)
    return ink


def _focus_coverage(plan: Dict[str, Any]) -> Dict[str, Any]:
    """主角是设计判断：面积、明暗、语义、位置与阅读路径不能由矩形面积排名代替。"""
    human = []
    for page in plan["pages"]:
        unmapped = page.get("focus_unmapped")
        if unmapped:
            reason = f"声明 {unmapped['kind']}：{unmapped['reason']}"
        else:
            reason = f"声明主角 {page.get('focus_element')}；请对照真实渲染判断层级与阅读路径"
        human.append({"page": page["id"], "reason": reason})
    return {"machine_checked": [], "human_required": human}


def _font_embedding_issues(output: str) -> List[str]:
    """严格解析 pdffonts 的 emb 列，不依赖字体名所在列或模糊子串。"""
    issues: List[str] = []
    for line in output.splitlines()[2:]:
        if not line.strip():
            continue
        columns = re.search(r"\s+(yes|no)\s+(yes|no)\s+(yes|no)\s+\d+\s+\d+", line)
        if not columns:
            issues.append(f"无法解析 pdffonts 行：{line}")
        elif columns.group(1) != "yes":
            issues.append(f"字体未嵌入：{line}")
    return issues


def _requested_fonts(plan: Dict[str, Any]) -> List[str]:
    fonts = plan.get("fonts", {})
    names = {fonts.get("heading", DEFAULT_FONTS["heading"]), fonts.get("body", DEFAULT_FONTS["body"])}
    for page in plan["pages"]:
        for el in page["elements"]:
            if el["type"] in TEXT_ELEMENT_TYPES:
                names.update(style["font"] for style in run_styles(el, fonts))
    return sorted(names)


def _font_substitution(plan: Dict[str, Any], pdffonts_output: str) -> Dict[str, Any]:
    """请求字体在本机是否安装、渲染实际嵌入了哪些字体。替换发生时，排版测量只代表替代字体。"""
    embedded = sorted({re.sub(r"^[A-Z]{6}\+", "", line.split()[0])
                       for line in pdffonts_output.splitlines()[2:] if line.strip()})
    fc_list, fc_match = shutil.which("fc-list"), shutil.which("fc-match")
    if not fc_list or not fc_match:
        return {"status": "NOT VERIFIED", "reason": "缺少 fontconfig（fc-list/fc-match），无法判定请求字体是否安装",
                "embedded_pdf_fonts": embedded}
    listing = _run_text([fc_list, ":", "family"], timeout=60).stdout
    installed = {family.strip().lower() for line in listing.splitlines() for family in line.split(",") if family.strip()}
    rows = []
    for name in _requested_fonts(plan):
        present = name.lower() in installed
        substitute = None if present else _run_text(
            [fc_match, "-f", "%{family[0]}", name], timeout=60).stdout.strip()
        rows.append({"requested": name, "installed": present, "fontconfig_substitute": substitute})
    return {"status": "SUBSTITUTED" if any(not r["installed"] for r in rows) else "MATCHED",
            "fonts": rows, "embedded_pdf_fonts": embedded}


def _is_blank_render(image) -> bool:
    """与页面众数灰阶相差明显的像素占比不超过 policy 下限即视为空白页（深底、浅底一视同仁）。"""
    hist = image.convert("L").histogram()
    mode = max(range(256), key=hist.__getitem__)
    distinct = sum(n for level, n in enumerate(hist) if abs(level - mode) > 8)
    return distinct / max(1, sum(hist)) <= float(POLICY["preview"]["blank_ink_min_ratio"])


_PAGE_PNG = re.compile(r"page-(\d+)\.png")


def _page_images(render_dir: Path, page_count: int) -> Tuple[List[Path], List[str]]:
    """按文件名里的整数页码把 pdftoppm 输出排成 PDF 页序：返回列表的下标 i 就是第 i + 1 页。

    pdftoppm 的约定只有「前缀-页码.png」；补零位数是实现细节（poppler 按总页数的位数补零），
    所以不按字符串排序。页码必须恰为 1..page_count、无重复、无无法解析的文件，否则返回空列表与问题：
    宁可不用像素证据，也不拿错页的图做判断。
    """
    found: Dict[int, List[str]] = {}
    issues: List[str] = []
    for path in render_dir.glob("page-*.png"):
        match = _PAGE_PNG.fullmatch(path.name)
        if match:
            found.setdefault(int(match.group(1)), []).append(path.name)
        else:
            issues.append(f"无法从文件名解析页码：{path.name}")
    issues += [f"第 {n} 页有多张 PNG：{sorted(names)}" for n, names in sorted(found.items()) if len(names) > 1]
    missing = [n for n in range(1, page_count + 1) if n not in found]
    if missing:
        issues.append(f"PNG 页码不连续：缺第 {missing} 页（PDF 共 {page_count} 页）")
    extra = sorted(n for n in found if not 1 <= n <= page_count)
    if extra:
        issues.append(f"PNG 页码超出 PDF 页数 {page_count}：{extra}")
    if issues:
        return [], issues
    return [render_dir / found[n][0] for n in range(1, page_count + 1)], []


def _slide_text(slide) -> str:
    parts: List[str] = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            parts.append(shape.text_frame.text)
        elif getattr(shape, "has_table", False) and shape.has_table:
            parts.extend(cell.text for row in shape.table.rows for cell in row.cells)
    return "".join(parts)


def _page_signature(page: Dict[str, Any]) -> Tuple[Any, ...]:
    """页的完整渲染指纹：文字相同但数据、颜色、背景或图层不同的页不能互相顶替。

    页序核对只看文字；两页文字完全相同时，任何错位都测不出来。只有两页**可见结构完全
    相同**（含图表数据、形状颜色、背景、裁切与 z_index）时，互换才既不可分辨也无害。
    这里直接用规范化后的整页元素 JSON，避免白名单遗漏未来新增字段。
    """
    visible = {
        "background": page.get("background"),
        "page_number_tone": page.get("page_number_tone"),
        "elements": page.get("elements", []),
    }
    return (json.dumps(visible, ensure_ascii=False, sort_keys=True, separators=(",", ":")),)


def _interchangeable(index: int, scores: List[float], best: float, rendered: List[Counter], want: Counter,
                     plan_pages: List[Dict[str, Any]]) -> bool:
    """本页是否与并列最优页可互换：文字精确一致，且结构完全相同（互换既测不出也无害）。"""
    if rendered[index] != want:
        return False
    tied = [j for j, score in enumerate(scores) if score >= best - 1e-9]
    mine = _page_signature(plan_pages[index]) if index < len(plan_pages) else None
    for j in tied:
        other = _page_signature(plan_pages[j]) if j < len(plan_pages) else None
        if mine != other:
            return False
    return True


def _page_text_order(slides, pages: List[Dict[str, Any]],
                     plan_pages: Optional[List[Dict[str, Any]]] = None) -> Tuple[List[str], Dict[str, Any]]:
    """PPTX 第 i 页的文字必须在 PDF 第 i 页命中最多：按字符多重集比较，与断行、词框切分无关。

    只用 PPTX 自身的文本框与表格文字（图表内部文字不计）。没有文字的页无法这样核对；文字与别页
    并列命中最高的页无法区分——两类都如实列出。无文字页的位置可以**由排除法确认**：当全稿只有
    一页无文字、且其余每一页都在自己的页码上唯一命中最优时，任何错位都会让某个文字页失去唯一
    性，因此该无文字页必然在位；两张以上无文字页之间可能互换，保留为未确认。
    """
    def chars(text: str) -> Counter:
        return Counter(ch for ch in text if not ch.isspace())

    rendered = [chars("".join(w["text"] for w in page["words"])) for page in pages]
    plan_pages = plan_pages or []
    issues: List[str] = []
    no_text: List[int] = []
    indistinct: List[int] = []
    checked = 0
    for i, slide in enumerate(slides):
        want = chars(_slide_text(slide))
        total = sum(want.values())
        if not total:
            no_text.append(i + 1)
            continue
        if i >= len(rendered):
            continue  # 页数不一致另行报告
        scores = [sum((want & got).values()) / total for got in rendered]
        best = max(scores)
        checked += 1
        if best < 1.0 - 1e-9:
            # 即使本页是“相对最高”，只要 PDF 文本层没有完整命中，页序就没有确认；
            # 不拿低命中率的页去支撑溢出、断行或页码证据。
            indistinct.append(i + 1)
        elif scores[i] < best - 1e-9:
            j = scores.index(best)
            issues.append(f"页序不对应：PPTX 第 {i + 1} 页的文字在 PDF 第 {i + 1} 页命中 {scores[i]:.0%}，"
                          f"在 PDF 第 {j + 1} 页命中 {best:.0%}")
        elif scores.count(best) > 1 and not _interchangeable(i, scores, best, rendered, want, plan_pages):
            # 并列最优且本页无法与并列页区分：文字不完全一致，或文字相同而版式/配图不同
            indistinct.append(i + 1)
    stats: Dict[str, Any] = {"text_checked_pages": checked, "no_text_pages": no_text, "indistinct_pages": indistinct}
    if not issues and not indistinct and len(no_text) == 1 and checked == len(slides) - 1:
        stats["confirmed_by_elimination"] = list(no_text)
    return issues, stats


def _contact_sheet(pngs: List[Path], plan: Dict[str, Any], out_path: Path,
                   unconfirmed: Optional[Set[int]] = None) -> Path:
    """整稿缩略总览：G-DECK 判断节奏、重复、失衡与跨页一致性的材料；它本身不评分。

    pngs 必须是 _page_images 校验过的页序列表（下标 i 即第 i + 1 页），标签才与 plan 页对应。
    页序未经文字核对确认的页在标签上标注「页序未证实」，避免把错页的观感当成该页的事实。
    """
    from PIL import Image, ImageDraw
    thumb_w, cols, gap, label_h = 320, 4, 16, 18
    thumbs = []
    for png in pngs:
        with Image.open(png) as source:
            image = source.convert("RGB")
        thumbs.append(image.resize((thumb_w, max(1, round(image.height * thumb_w / image.width)))))
    th = max(t.height for t in thumbs)
    rows = math.ceil(len(thumbs) / cols)
    sheet = Image.new("RGB", (cols * (thumb_w + gap) + gap, rows * (th + label_h + gap) + gap), (232, 232, 232))
    draw = ImageDraw.Draw(sheet)
    pages = plan["pages"] if plan else []
    unconfirmed = unconfirmed or set()
    for i, thumb in enumerate(thumbs):
        x, y = gap + (i % cols) * (thumb_w + gap), gap + (i // cols) * (th + label_h + gap)
        role = pages[i]["role"] if i < len(pages) else "?"
        label = f"{i + 1:02d}  {role}"
        if (i + 1) in unconfirmed:
            label += "（页序未证实）"
        draw.text((x, y + 2), label, fill=(64, 64, 64))
        sheet.paste(thumb, (x, y + label_h))
    sheet.save(out_path)
    return out_path


def _with_page_numbers(plan: Dict[str, Any]) -> Dict[str, Any]:
    """给**核对用**的 plan 副本注入编译器派生的页码元素。

    只在 audit_pptx / _text_contrast 使用：让同一套对象核对与对比度计算直接覆盖页码，
    而不是另写一套几何/样式解析。作者 plan、validate_plan 和断行检查不使用这个副本。
    """
    out = dict(plan)
    pages: List[Dict[str, Any]] = []
    for i, page in enumerate(plan["pages"]):
        derived = page_number_element(plan, i)
        copy = dict(page)
        copy["elements"] = list(page["elements"]) + ([derived] if derived is not None else [])
        pages.append(copy)
    out["pages"] = pages
    return out


def _page_number_audit(plan: Dict[str, Any], pptx_path: Path, presentation=None) -> Tuple[List[CompileError], List[Dict[str, Any]]]:
    """E017：在**实际 PPTX 对象**上核对页码存在性、编号、角区、左对齐、样式、图层与统一参考线。

    plan 里写了什么不算数；这里读的是 python-pptx 打开后的真实形状。
    """
    errors: List[CompileError] = []
    records: List[Dict[str, Any]] = []
    cfg = page_number_config(plan)
    specs = page_number_specs(plan)
    fonts = plan.get("fonts", {})
    try:
        prs = presentation if presentation is not None else Presentation(str(pptx_path))
    except Exception as exc:
        return [CompileError("E017", ErrorLevel.HARD, "deck.page_numbers", "页码对象可在 PPTX 中核对",
                             f"PPTX 无法打开：{type(exc).__name__}: {exc}", "不得推定页码存在")], records
    left_edges: List[float] = []
    for i, (label, spec) in enumerate(specs):
        pid = plan["pages"][i]["id"]
        loc = f"page[{pid}].page_number"
        expected_name = f"{pid}.{PAGE_NUMBER_ID}"
        rec: Dict[str, Any] = {"page": pid, "expected": label, "corner": cfg["corner"], "object": expected_name}
        if i >= len(prs.slides):
            errors.append(CompileError("E017", ErrorLevel.HARD, loc, "每页必须有页码对象",
                                       f"PPTX 只有 {len(prs.slides)} 页，plan 第 {i+1} 页不存在", "页码缺失，禁止交付"))
            rec["status"] = "BLOCK"
            records.append(rec)
            continue
        slide = prs.slides[i]
        shapes = list(slide.shapes)
        # 精确核对派生对象名，不按最后一个点段猜测：合法元素 id="foo._page_number"
        # 的对象名是 `<pid>.foo._page_number`，不能因此被认成第二个页码。
        found = [s for s in shapes if s.name == expected_name]
        if label is None:
            if found:
                errors.append(CompileError("E017", ErrorLevel.HARD, loc, "声明省略页码的页面不得出现页码对象",
                                           f"发现 {[s.name for s in found]}", "例外必须显式且真实"))
                rec["status"] = "BLOCK"
            else:
                rec.update(status="OMITTED_DECLARED", reason=cfg["omit_reason"])
            records.append(rec)
            continue
        if len(found) != 1 or found[0].name != expected_name:
            errors.append(CompileError(
                "E017", ErrorLevel.HARD, loc, "每页有且仅有一个确定性派生的页码对象",
                f"期望 {expected_name!r}，实际 {[s.name for s in found]}",
                "页码缺失、重复或对象身份被改写都会阻断交付；不得用手写 footer 顶替"))
            rec["status"] = "BLOCK"
            records.append(rec)
            continue
        shape = found[0]
        rec["status"] = "PASS"
        rec["text"] = shape.text_frame.text if getattr(shape, "has_text_frame", False) else None
        rec["geometry_in"] = [_emu_box(shape)[j] for j in range(4)]
        if not getattr(shape, "has_text_frame", False):
            errors.append(CompileError("E017", ErrorLevel.HARD, loc, "页码必须是原生可编辑文本对象",
                                       "对象无文本框架", "页码不得图片化"))
            rec["status"] = "BLOCK"
        else:
            if shape.text_frame.text != label:
                errors.append(CompileError("E017", ErrorLevel.HARD, loc, "页码文字必须与实际页序一致",
                                           f"期望 {label!r}，实际 {shape.text_frame.text!r}",
                                           "编号由 pages 顺序派生；乱序/跳号不得交付"))
                rec["status"] = "BLOCK"
            para = shape.text_frame.paragraphs[0]
            if para.alignment != PP_ALIGN.LEFT:
                errors.append(CompileError("E017", ErrorLevel.HARD, loc, "页码文字必须框内左对齐",
                                           f"实际 alignment={para.alignment}",
                                           "页码框位于右侧角区，但数字本身不得右对齐"))
                rec["status"] = "BLOCK"
            runs = [r for r in para.runs if r.text]
            expected_style = text_style(spec, fonts)
            style_ok = bool(runs)
            for run in runs:
                if (run.font.size is None or abs(run.font.size.pt - expected_style["size"]) > 1e-6
                        or run.font.name != expected_style["font"]
                        or run.font.color.type != 1
                        or str(run.font.color.rgb) != expected_style["color"].lstrip("#")
                        or bool(run.font.bold) != bool(expected_style["bold"])):
                    style_ok = False
            if not style_ok:
                errors.append(CompileError("E017", ErrorLevel.HARD, loc, "页码字体/字号/字重/墨色全稿一致",
                                           f"期望 {expected_style['size']:g}pt / {expected_style['font']} / {expected_style['color']} / 非粗体",
                                           "样式只在 policy.page_number 定义一次；不得逐页变体"))
                rec["status"] = "BLOCK"
            rec["align"] = str(para.alignment)
        issues = _geometry_issues(shape, (spec["x"], spec["y"], spec["w"], spec["h"]), expected_name)
        if issues:
            errors.append(CompileError("E017", ErrorLevel.HARD, loc, "页码框必须位于指定角区的统一安全边距内",
                                       "；".join(issues), "位置由 policy 决定；不得随手挪动或贴边"))
            rec["status"] = "BLOCK"
        left_edges.append(float(shape.left) / EMU_PER_IN)
        if shapes[-1].name != expected_name:
            errors.append(CompileError("E017", ErrorLevel.HARD, loc, "页码对象必须位于最上层，不被其他对象遮挡",
                                       f"最上层对象是 {shapes[-1].name!r}", "调整图层顺序，不得删除页码"))
            rec["status"] = "BLOCK"
        rec["style_pt"] = spec["font_size"]
        rec["ink"] = spec["color"]
        records.append(rec)
    if len(left_edges) > 1 and max(left_edges) - min(left_edges) > float(POLICY["canvas"]["bound_tolerance_in"]):
        errors.append(CompileError(
            "E017", ErrorLevel.HARD, "deck.page_numbers", "整稿页码文本框左边缘必须与统一参考线对齐",
            f"左边缘范围 {min(left_edges):.3f}–{max(left_edges):.3f} in",
            "角区与样式在整稿级配置里只定义一次；不得逐页偏移"))
        for rec in records:
            if rec.get("status") == "PASS":
                rec["status"] = "BLOCK"
    return errors, records


def _page_number_render(plan: Dict[str, Any], pdf_pages: List[Dict[str, Any]],
                        confirmed: Optional[Set[int]], pngs_ok: bool) -> Tuple[List[CompileError], List[Dict[str, Any]]]:
    """E018：核对真实 PDF 的页码词框、位置与文字重叠。

    词框存在并不独立证明像素可见；E017 检查 PPTX 图层，E013 检查页码与背景对比，
    仍需看 PNG。此处不把框外的年份、数据或标签误当成页码。
    """
    errors: List[CompileError] = []
    records: List[Dict[str, Any]] = []
    cfg = page_number_config(plan)
    sw, sh = _slide_dimensions(plan.get("aspect_ratio"))
    zone = POLICY["page_number"]["corner_zone"]
    zone_x0_pt = sw * (1.0 - float(zone["x_from_right_ratio"])) * 72.0
    zone_h_pt = sh * float(zone["y_from_edge_ratio"]) * 72.0
    if cfg["corner"] == "top_right":
        zone_y0_pt, zone_y1_pt = 0.0, zone_h_pt
    else:
        zone_y0_pt, zone_y1_pt = sh * 72.0 - zone_h_pt, sh * 72.0
    tol = float(POLICY["preview"]["overflow_tolerance_pt"])
    text_margin_pt = float(POLICY["typography"]["margin_in"]["x"]) * 72.0

    for i, (label, spec) in enumerate(page_number_specs(plan)):
        pid = plan["pages"][i]["id"]
        loc = f"page[{pid}].page_number.render"
        rec: Dict[str, Any] = {"page": pid, "expected": label, "corner": cfg["corner"], "status": "NOT VERIFIED"}
        if label is None:
            rec.update(status="OMITTED_DECLARED", reason=cfg["omit_reason"])
            records.append(rec)
            continue
        if not pngs_ok or i >= len(pdf_pages) or (confirmed is not None and (i + 1) not in confirmed):
            rec["reason"] = ("无可用真实渲染证据" if not pngs_ok or i >= len(pdf_pages)
                             else "该页页序未经文字核对确认，未使用词框证据")
            records.append(rec)
            continue
        page_box = pdf_pages[i]
        words = page_box["words"]

        def in_zone(w: Dict[str, Any]) -> bool:
            cx, cy = (w["x0"] + w["x1"]) / 2.0, (w["y0"] + w["y1"]) / 2.0
            return zone_x0_pt <= cx <= sw * 72.0 and zone_y0_pt <= cy <= zone_y1_pt

        # 页码归属以派生对象的实际声明框为准，不能把角区内的年份、数据标签等都当成页码。
        bx0, by0 = spec["x"] * 72.0 - tol, spec["y"] * 72.0 - tol
        bx1, by1 = (spec["x"] + spec["w"]) * 72.0 + tol, (spec["y"] + spec["h"]) * 72.0 + tol
        matching = [w for w in words if w["text"].strip() == label]
        box_hits = [w for w in matching if in_zone(w)
                    and bx0 <= (w["x0"] + w["x1"]) / 2 <= bx1
                    and by0 <= (w["y0"] + w["y1"]) / 2 <= by1]
        if len(box_hits) != 1:
            detail = (f"期望 {label!r}，页码框内匹配 {len(box_hits)} 个；"
                      f"整页匹配位置 {[round(w['x0']) for w in matching][:5]}")
            errors.append(CompileError(
                "E018", ErrorLevel.HARD, loc, "真实 PDF 的页码框内必须出现且仅出现一个页码词框",
                detail, "可能缺失、重复、乱序、错位、被遮挡或被裁切；逐项检查渲染图"))
            rec["status"] = "BLOCK"
            records.append(rec)
            continue
        glyph = box_hits[0]
        want_x = (spec["x"] * 72.0) + text_margin_pt
        want_y = spec["y"] * 72.0
        problems: List[str] = []
        if abs(glyph["x0"] - want_x) > tol:
            problems.append(f"字形左边缘 {glyph['x0']:.1f}pt 偏离统一参考线 {want_x:.1f}pt（容差 {tol:.1f}pt）")
        if abs(glyph["y0"] - want_y) > 2 * tol:
            problems.append(f"字形顶部 {glyph['y0']:.1f}pt 偏离角区基线 {want_y:.1f}pt")
        page_w, page_h = page_box["width"], page_box["height"]
        if (glyph["x0"] < 2 or glyph["y0"] < 2 or glyph["x1"] > page_w - 2 or glyph["y1"] > page_h - 2):
            problems.append("字形超出页面安全范围，疑似被裁切")
        overlapping = [w for w in words if w is not glyph
                       and min(w["x1"], glyph["x1"]) - max(w["x0"], glyph["x0"]) > 0.5
                       and min(w["y1"], glyph["y1"]) - max(w["y0"], glyph["y0"]) > 0.5]
        if overlapping:
            problems.append(f"与 {len(overlapping)} 个其他词框重叠（{''.join(w['text'] for w in overlapping[:3])[:24]}…）")
        rec.update(status="BLOCK" if problems else "PASS",
                   glyph_bbox_pt=[round(glyph["x0"], 1), round(glyph["y0"], 1),
                                  round(glyph["x1"], 1), round(glyph["y1"], 1)],
                   left_edge_pt=round(want_x, 1), problems=problems)
        if problems:
            errors.append(CompileError("E018", ErrorLevel.HARD, loc,
                                       "页码词框须左对齐、位于统一角区、不越界且不与其他词框重叠",
                                       "；".join(problems), "页码是导航信息，不能用风格或空间不足为由取消"))
        records.append(rec)
    return errors, records


def _render(pptx_path: Path, render_dir: Path, plan: Optional[Dict[str, Any]], presentation=None) -> Tuple[Dict[str, Any], List[str], Dict[str, Any]]:
    """从本次 PPTX 重新生成 PDF/PNG 证据并执行渲染侧事实检查；返回 (G-PRE, 日志, 渲染事实)。"""
    tools = {
        "soffice": shutil.which("soffice") or shutil.which("libreoffice"),
        "pdftoppm": shutil.which("pdftoppm"),
        "pdftotext": shutil.which("pdftotext"), "pdffonts": shutil.which("pdffonts"),
    }
    facts: Dict[str, Any] = {}
    if not all(tools.values()):
        missing = sorted(k for k, v in tools.items() if not v)
        return _gate("NOT VERIFIED", issues=[f"缺少真实渲染工具：{missing}"]), [], facts
    if render_dir.exists() and any(render_dir.iterdir()):
        return _gate("BLOCK", issues=[f"渲染目录必须为空，拒绝覆盖旧证据：{render_dir}"]), [], facts
    render_dir.mkdir(parents=True, exist_ok=True)
    # LibreOffice 的用户配置放进系统临时目录，用完即删：不污染证据目录，也不让重跑因\"目录非空\"失败。
    profile = Path(tempfile.mkdtemp(prefix="lo-profile-"))
    logs: List[str] = []
    try:
        command = [tools["soffice"], f"-env:UserInstallation={profile.as_uri()}", "--headless",
                   "--convert-to", "pdf", "--outdir", str(render_dir), str(pptx_path)]
        proc = _run_text(command, timeout=180)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    logs.append((proc.stdout + proc.stderr).strip())
    pdf = render_dir / f"{pptx_path.stem}.pdf"
    if proc.returncode != 0 or not pdf.is_file():
        return _gate("BLOCK", issues=["PPTX→PDF 转换失败或未生成本次 PPTX 对应 PDF"], environment=logs), logs, facts

    proc = _run_text([tools["pdftoppm"], "-r", str(POLICY["preview"]["render_dpi"]), "-png", str(pdf),
                      str(render_dir / "page")], timeout=180)
    logs.append((proc.stdout + proc.stderr).strip())
    if proc.returncode != 0:
        return _gate("BLOCK", issues=["PDF→PNG 转换失败"], environment=logs), logs, facts
    font_proc = _run_text([tools["pdffonts"], str(pdf)], timeout=180)
    (render_dir / "pdffonts.txt").write_text(font_proc.stdout, encoding="utf-8")
    (render_dir / "render.log").write_text("\n".join(logs), encoding="utf-8")
    issues: List[str] = []
    limitations: List[str] = []
    not_verified: List[str] = []
    extra: Dict[str, Any] = {"render_dpi": POLICY["preview"]["render_dpi"]}
    if font_proc.returncode != 0:
        issues.append("pdffonts 执行失败")

    from PIL import Image
    prs = presentation if presentation is not None else Presentation(str(pptx_path))
    slides = prs.slides
    expected_pages = len(slides)
    bbox_path = render_dir / "bbox.html"
    pages = _pdftotext_bbox(pdf, tools["pdftotext"], bbox_path)
    pdf_pages = len(pages)
    # bbox 已包含每页实际尺寸和页数：复用一次解析，且逐页核对而非只看首页。
    want_w, want_h = prs.slide_width / EMU_PER_IN * 72.0, prs.slide_height / EMU_PER_IN * 72.0
    tol = float(POLICY["canvas"]["bound_tolerance_in"]) * 72.0
    extra["pdf_page_sizes_pt"] = [[page["width"], page["height"]] for page in pages]
    if pages:
        extra["pdf_page_size_pt"] = extra["pdf_page_sizes_pt"][0]  # 保留原报告字段
    extra["expected_page_size_pt"] = [round(want_w, 3), round(want_h, 3)]
    for index, page in enumerate(pages, 1):
        got_w, got_h = page["width"], page["height"]
        if abs(got_w - want_w) > tol or abs(got_h - want_h) > tol:
            issues.append(f"PDF 第 {index} 页尺寸与 PPTX 画布不一致：{got_w}×{got_h}pt，"
                          f"期望 {want_w:.3f}×{want_h:.3f}pt（容差 {tol:.2f}pt）")
    # 页序三方对应：PNG→PDF 靠文件名整数页码（pdftoppm 约定），PDF→PPTX 靠逐页文字命中。
    # 哪一环不成立，就不使用依赖它的证据：错页的图和错页的词框都会把别页的事实安到本页头上。
    count_issues: List[str] = []
    if pdf_pages != expected_pages:
        count_issues.append(f"PPTX/PDF 页数不一致：pptx={expected_pages}, PDF bbox={pdf_pages}")
    pngs, png_issues = _page_images(render_dir, pdf_pages)
    text_issues, text_order = _page_text_order(slides, pages, plan["pages"] if plan else [])
    pages_ok = not count_issues and not text_issues
    pngs_ok = pages_ok and not png_issues
    # 页序部分未证实：无文字页（排除法未覆盖）与无法区分的重复页。这些页的词框、像素与
    # 字形证据一律不采信——错页的图会把别页的事实安到本页头上；未核验不等于通过。
    unconfirmed = sorted((set(text_order["no_text_pages"]) - set(text_order.get("confirmed_by_elimination", [])))
                         | set(text_order["indistinct_pages"]))
    confirmed = {n for n in range(1, expected_pages + 1) if n not in unconfirmed}
    # PDF 页码词框只核对落点/文字重叠；对比与图层分别由 E013/E017 核对。
    if plan is not None:
        pn_render_errors, pn_render_records = _page_number_render(plan, pages, confirmed, pngs_ok)
        extra["page_numbers"] = pn_render_records
        facts["page_number_errors"] = pn_render_errors
        facts["page_number_records"] = pn_render_records
        for rec in pn_render_records:
            if rec["status"] == "NOT VERIFIED":
                item = f"[{rec['page']}.page_number] 页码真实渲染未核验：{rec['reason']}；该页证据不计入通过"
                limitations.append(item)
                not_verified.append(item)
    issues.extend(count_issues + png_issues + text_issues)
    if not pages_ok or png_issues:
        order_status = "BLOCK"
    elif unconfirmed:
        order_status = "PARTIAL"
    else:
        order_status = "PASS"
    extra["page_order"] = dict(
        status=order_status, pptx_slides=expected_pages, pdf_pages=pdf_pages, png_pages=len(pngs),
        basis="PNG 按文件名整数页码对应 PDF 页（校验 1..N 连续、无重复、无杂项）；PDF 第 i 页须是 PPTX 第 i 页文字命中最多的页",
        unconfirmed_pages=unconfirmed,
        **text_order,
    )
    evidence = [str(pdf)] + [str(p) for p in pngs] + [str(bbox_path), str(render_dir / "pdffonts.txt")]
    if not png_issues:
        for png in pngs:
            with Image.open(png) as source:
                if _is_blank_render(source):
                    issues.append(f"疑似空白渲染页：{png.name}")
    issues.extend(_font_embedding_issues(font_proc.stdout))
    if unconfirmed:
        item = (f"第 {unconfirmed} 页页序未经文字核对确认（无文字且排除法不适用，或与别页无法区分）："
                "未使用这些页的词框/像素证据执行溢出、断行与字形对比；缩略总览中对应标签已标注")
        limitations.append(item)
        not_verified.append(item)
    if pngs_ok:
        sheet = _contact_sheet(pngs, plan or {}, render_dir / "contact_sheet.png", set(unconfirmed))
        facts["contact_sheet"] = str(sheet)
        evidence.append(str(sheet))
    else:
        limitations.append("页序未通过校验：不生成缩略总览，不用像素确认溢出")
    observations: List[str] = []
    measured: Dict[Tuple[str, str], Optional[float]] = {}
    # 与图表/表格叠放的文本框，词框可能被图表内部文字污染；机器不下断行/溢出结论，
    # 改为逐项列进 limitations 让 G-DECK 看渲染图确认（误报成本高于自动判定的收益）。
    ambiguous_raw = _ambiguous_text(plan) if plan else {}
    skip = {pi: {eid for eid, _ in items} for pi, items in ambiguous_raw.items()}
    if plan:
        for pi, items in sorted(ambiguous_raw.items()):
            if (pi + 1) not in confirmed:
                continue
            for eid, chart_ids in items:
                item = (f"[{plan['pages'][pi]['id']}.{eid}] 文本框与图表/表格 {chart_ids} 叠放："
                        "PDF 词框无法可靠区分图表内部文字与该文本框，断行/孤字/溢出未由机器判定，"
                        f"请人工对照 page-{pi + 1}.png 确认")
                limitations.append(item)
                not_verified.append(item)
    if plan is None:
        issues.append("无有效 plan，无法核对声明文本框的真实渲染")
    else:
        if pages_ok:
            candidates, measured = _overflow_violations(plan, pages, confirmed, skip)
            from PIL import ImageDraw
            compositor = _Compositor(Image, ImageDraw, float(POLICY["preview"]["render_dpi"]))
            for cand in candidates:
                if not cand["hits"]:
                    issues.append(cand["detail"])
                    continue
                if (cand["page_index"] + 1) not in confirmed:
                    limitations.append(cand["detail"] + "；该页页序未经文字核对确认，未用词框或像素下结论")
                    continue
                if not pngs_ok:
                    issues.append(cand["detail"] + "；PNG 页码未通过校验，不能用像素确认，按词框判为溢出")
                    continue
                with Image.open(pngs[cand["page_index"]]) as source:
                    ink = _ink_outside(slides[cand["page_index"]], source.convert("RGB"), cand["box"], cand["hits"], compositor)
                if ink is None:
                    issues.append(cand["detail"] + "；框外有不可还原图层，按词框保守判为溢出")
                elif ink > 0:
                    issues.append(cand["detail"] + f"；框外笔画 {ink} px，溢出")
                else:
                    observations.append(cand["detail"] + "；框外无笔画（悬挂标点或字体行框），不计溢出")
                    measured[cand["key"]] = 0.0
            issues.extend(_line_break_issues(plan, pages, confirmed, skip))
            slide_tolerance = float(POLICY["preview"]["slide_bounds_tolerance_pt"])
            for index, page in enumerate(pages):
                for word in page["words"]:
                    if (word["x0"] < -slide_tolerance or word["y0"] < -slide_tolerance
                            or word["x1"] > page["width"] + slide_tolerance or word["y1"] > page["height"] + slide_tolerance):
                        issues.append(f"第 {index + 1} 页文字越出页面：{word['text']!r}")
            facts.update(pages=pages, overflow_measured=measured)
            facts["confirmed_pages"] = confirmed
        else:
            limitations.append("PDF 页与 PPTX 页未能对应：逐文本框溢出、断行、页面边界及基于字形的对比未执行")
        has_text = any(el["type"] in TEXT_ELEMENT_TYPES and el["content"].strip()
                       for page in plan["pages"] for el in page["elements"])
        if has_text and len(font_proc.stdout.splitlines()) <= 2:
            issues.append("页面含文本但 pdffonts 未报告字体")
        substitution = _font_substitution(plan, font_proc.stdout)
        if substitution["status"] != "MATCHED":
            limitations.append("请求字体未全部安装或无法核验：本次溢出、断行与对比的字形区域基于替代字体，"
                               "目标字体下的排版须由 G-TGT 确认")
        extra.update(text_elements_checked=len(measured), font_substitution=substitution)
    extra.update(
        environment="LibreOffice + poppler：只证明本次 PPTX 在当前环境的预览，不替代 G-DECK/G-TGT",
        scope="PNG/PDF/PPTX 页序对应与 PDF 页面尺寸、非空白、声明文本框的笔画边界（词框越界经像素确认）、页面边界、"
              "逐文本框的避头尾、孤字与数字/单位拆行（图表/表格叠放的文本框按 limitations 交人工）、PDF 字体嵌入与"
              "请求字体替换；不证明审美、目标 PowerPoint 排版或无障碍合规",
        limitations=limitations,
        not_verified=not_verified,
        observations=observations,
    )
    status = "BLOCK" if issues else ("NOT VERIFIED" if not_verified else "PASS")
    return _gate(status, evidence, issues, **extra), logs, facts


def _review_gate(config: Dict[str, Any], gate_id: str, target: bool = False, pptx_sha256: Optional[str] = None, evidence_root: Optional[Path] = None) -> Dict[str, Any]:
    review = config.get("reviews", {}).get(gate_id, {})
    status = str(review.get("status", "NOT_VERIFIED")).replace("_", " ").upper()
    if target and not config.get("target_powerpoint", {}).get("configured", False):
        return _gate("NOT VERIFIED", issues=["目标 PowerPoint 版本/OS/字体环境待配置"])
    if status == "NOT VERIFIED":
        return _gate("NOT VERIFIED", issues=[review.get("note") or "未执行人工/目标环境验证"])
    if status == "BLOCK":
        return _gate("BLOCK", issues=[review.get("note") or "人工复核阻断"])
    if status == "PASS":
        missing = [field for field in ("reviewer", "evidence") if not review.get(field)]
        if target:
            target_config = config.get("target_powerpoint", {})
            missing += [field for field in ("version", "operating_system", "fonts") if not target_config.get(field)]
        if review.get("pptx_sha256") != pptx_sha256 or not pptx_sha256:
            missing.append("匹配本次 PPTX 的 pptx_sha256")
        evidence_path = Path(str(review.get("evidence") or ""))
        if not evidence_path.is_absolute():
            evidence_path = (evidence_root or Path.cwd()) / evidence_path
        if not evidence_path.is_file():
            missing.append("可读取的本地 evidence 文件")
        if missing:
            return _gate("BLOCK", issues=[f"PASS 记录缺少必要证据字段：{missing}"])
        return _gate("PASS", [str(evidence_path)], reviewer=review["reviewer"], evidence_sha256=_sha256(evidence_path), pptx_sha256=pptx_sha256)
    return _gate("NOT VERIFIED", issues=[f"复核状态不是 PASS：{status or '空'}"])


def _sampling(plan: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    pages = plan["pages"]
    population = [p["id"] for p in pages]
    configured = set(config.get("key_pages", []) + config.get("high_risk_pages", []) + config.get("historical_issue_pages", []))
    mandatory_roles = list(POLICY["sampling"]["mandatory_roles"])
    reasons: Dict[str, List[str]] = {}
    for page in pages:
        pid = page["id"]
        if pid in configured:
            reasons.setdefault(pid, []).append("配置指定")
        if any(el["type"] in {"chart", "image"} for el in page["elements"]):
            reasons.setdefault(pid, []).append("含图表/图像")
        if page["role"] in mandatory_roles:
            reasons.setdefault(pid, []).append(f"叙事角色={page['role']}")
    mandatory = [pid for pid in population if pid in reasons]
    low_risk = [pid for pid in population if pid not in reasons]
    rate, seed = float(config["sampling"]["low_risk_rate"]), int(config["sampling"]["seed"])
    # 比例的基数是低风险页面，不是全部页面；无低风险页或比例为 0 时样本为 0。
    count = 0 if not low_risk or rate == 0 else min(len(low_risk), math.ceil(len(low_risk) * rate))
    sampled = sorted(random.Random(seed).sample(low_risk, count)) if count else []
    return {
        "population": population, "mandatory_pages": mandatory, "mandatory_reasons": {pid: reasons[pid] for pid in mandatory},
        "low_risk_population": low_risk, "sampled_low_risk_pages": sampled,
        "selected_pages": [pid for pid in population if pid in reasons or pid in sampled],
        "rate": rate, "rate_base": "low_risk_population", "seed": seed, "method": "seeded random.sample",
        "purpose": "安排额外深查；不替代 G-DECK 逐页及整稿视觉验收",
        "role_mix": {role: sum(1 for p in pages if p["role"] == role) for role in sorted({p["role"] for p in pages})},
        "configuration_issues": [f"配置页面不在 plan 中：{pid}" for pid in sorted(configured - set(population))],
    }


_BLANKET_SCOPE_RE = re.compile(
    r"^(全部|所有|任何|一切|整体|整稿|所有页面|全部页面|所有项|全部项|all|any|everything|\*|等|等等|[。\.，,\s])*$",
    re.IGNORECASE)


def _parse_expiry(text: Any) -> Optional[date]:
    """从 expires_when 提取可核验的日期（ISO / 斜杠 / 中文年月日）；条件式表述返回 None。"""
    match = re.search(r"(20\d{2})\s*[-/年]\s*(\d{1,2})\s*[-/月]\s*(\d{1,2})", str(text or ""))
    if not match:
        return None
    try:
        return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return None


def _gate_issue_keys(gate: Optional[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """把一门 NOT VERIFIED 的未验证项拆成可绑定的问题键。

    键优先取方括号定位符（[页面.元素]）；没有定位符时以完整问题描述为键。
    风险批准的 scope 必须点名具体键；门未验证但没有任何问题描述时不能自动覆盖。
    """
    keys: List[Dict[str, Any]] = []
    for text in list((gate or {}).get("issues") or []) + list((gate or {}).get("not_verified") or []):
        locators = re.findall(r"\[([^\]]+)\]", text)
        if locators:
            keys.extend({"locator": loc, "anchors": [loc]} for loc in locators)
        elif text.strip():
            keys.append({"locator": text, "anchors": [text]})
    if not keys and (gate or {}).get("status") == "NOT VERIFIED":
        keys.append({"locator": "未提供未验证项，不能批准", "anchors": []})
    return keys


def _scope_covers(scope: str, key: Dict[str, Any]) -> bool:
    if not key["anchors"]:
        return False  # 缺少问题描述，无法证明 scope 对应具体未验证项
    return any(anchor and re.search(r"(?<![A-Za-z0-9_.-])" + re.escape(anchor) + r"(?![A-Za-z0-9_.-])", scope)
               for anchor in key["anchors"])


_ID_REF_RE = re.compile(r"(?<![A-Za-z0-9_.-])[A-Za-z]{1,6}\d+[A-Za-z0-9_-]*(?:\.[A-Za-z0-9_-]+)*(?![A-Za-z0-9_.-])")


def _plan_ids(plan: Optional[Dict[str, Any]]) -> Set[str]:
    """plan 中真实存在的页面 id 与元素 id：scope 引用它们时必须对得上。"""
    ids: Set[str] = set()
    for page in (plan or {}).get("pages") or []:
        if isinstance(page, dict):
            if page.get("id"):
                ids.add(str(page["id"]))
            for el in page.get("elements") or []:
                if isinstance(el, dict) and el.get("id"):
                    eid = str(el["id"])
                    ids.add(eid)
                    if page.get("id"):
                        ids.add(f"{page['id']}.{eid}")
    return ids


def _scope_id_violations(scope: str, known: Set[str]) -> List[str]:
    """对常见 p1 / p1.title / p1.title-note 形态辅助检查是否存在。

    这不是任意 Unicode 或自由文本 ID 的解析器；具体未验证项仍须由 _scope_covers 精确匹配。
    """
    if not known:
        return []
    return [f"scope 引用了本交付中不存在的页面/元素：{token}"
            for token in sorted(set(_ID_REF_RE.findall(scope))) if token not in known]


def _risk_records(config: Dict[str, Any], gates: Dict[str, Dict[str, Any]], required: Set[str],
                  pptx_sha256: Optional[str], plan: Optional[Dict[str, Any]] = None
                  ) -> Tuple[List[Dict[str, Any]], List[str], List[Dict[str, Any]]]:
    """逐条核验风险批准：门与类别权限、交付物哈希、有效期、scope 是否绑定本交付真实存在的未验证项。

    返回 (记录, 阻断级问题, 逐条核验结果)。越权、哈希不符、过期、笼统范围或范围不对应
    任何真实未验证项都是阻断级；仅有效期无法机检的记录不算覆盖，交给未覆盖判定。
    """
    records = config.get("risk_approvals", [])
    issues: List[str] = []
    review: List[Dict[str, Any]] = []
    today = datetime.now(timezone.utc).date()
    for i, record in enumerate(records):
        gate_id = record.get("gate")
        violations: List[str] = []
        caveats: List[str] = []
        if gate_id not in GATE_IDS:
            violations.append("引用未知质量门")
        elif gate_id not in HUMAN_GATES:
            violations.append("只能覆盖人工门（G-DECK / G-TGT）")
        elif gate_id in required:
            violations.append("强制门不得以风险批准放行")
        if record.get("category") in POLICY["release"]["non_approvable"]:
            violations.append(f"类别「{record.get('category')}」不可批准")
        if pptx_sha256 and str(record.get("pptx_sha256") or "").lower() != pptx_sha256:
            violations.append("pptx_sha256 与本次交付物不符")
        expiry = _parse_expiry(record.get("expires_when"))
        if expiry is None:
            caveats.append("有效期无法核验（未能从 expires_when 解析日期），该记录不计为覆盖")
        elif expiry < today:
            violations.append(f"已过期（{expiry.isoformat()}）")
        scope = str(record.get("scope") or "")
        if not scope.strip():
            violations.append("scope 为空")
        elif _BLANKET_SCOPE_RE.match(scope.strip()):
            violations.append("scope 为笼统范围，未绑定具体问题")
        else:
            violations.extend(_scope_id_violations(scope, _plan_ids(plan)))
        keys = _gate_issue_keys(gates.get(gate_id)) if gate_id in gates else []
        covered = [key for key in keys if _scope_covers(scope, key)]
        if keys and not covered and scope.strip() and not _BLANKET_SCOPE_RE.match(scope.strip()):
            violations.append("scope 未对应本交付中该门任何未验证项")
        review.append({
            "index": i, "gate": gate_id, "category": record.get("category"), "scope": scope,
            "covers": [key["locator"] for key in covered], "violations": violations, "caveats": caveats,
            "valid": not violations and not caveats,
        })
        if violations:
            issues.append(f"风险记录 {i}（{gate_id}）：" + "；".join(violations))
    return records, issues, review


def _route_diagnostics(errors: List[CompileError], plan: Optional[Dict[str, Any]]) -> Tuple[Dict[str, List[str]], List[str], List[str]]:
    """把诊断按 GATE_OF_CODE 归门：返回 (各门阻断项, 合规待定项, 已具名复核项)。"""
    blocking: Dict[str, List[str]] = {gate: [] for gate in GATE_IDS}
    pending_compliance: List[str] = []
    reviewed: List[str] = []
    unresolved = {id(e) for e in unresolved_reviews(plan or {}, errors)}
    for error in errors:
        if error.level == ErrorLevel.INFO:
            continue
        gate = GATE_OF_CODE.get(error.code)
        if gate is None:
            blocking["G-STR"].append(f"诊断码 {error.code} 未登记归属门（实现漂移）：{error}")
        elif error.level == ErrorLevel.SOFT and id(error) not in unresolved:
            reviewed.append(str(error))
        elif gate == "G-COMPLY" and error.level == ErrorLevel.SOFT:
            pending_compliance.append(str(error))
        else:
            blocking[gate].append(str(error))
    return blocking, pending_compliance, reviewed


def _supersede_estimates(errors: List[CompileError], measured: Dict[Tuple[str, str], Optional[float]]) -> List[CompileError]:
    """E010 是装载估算：真实渲染测量到同一文本框后以测量为准，估算降为 INFO 并写明测量结果。"""
    tolerance = float(POLICY["preview"]["overflow_tolerance_pt"])
    out: List[CompileError] = []
    for error in errors:
        match = re.fullmatch(r"page\[(.+?)\]\.element\[\d+\]\((.+?)\)\.content", error.location) if error.code == "E010" else None
        key = (match.group(1), match.group(2)) if match else None
        if key and key in measured and measured[key] is not None:
            over = measured[key]
            verdict = "未溢出" if over <= tolerance else f"溢出 {over:.1f}pt（由 G-PRE 阻断）"
            out.append(CompileError(error.code, ErrorLevel.INFO, error.location, error.condition,
                                    f"{error.actual}；真实渲染测量：{verdict}", "估算已被本次渲染测量取代"))
        else:
            out.append(error)
    return out


def _declared_intent(plan: Dict[str, Any]) -> Dict[str, Any]:
    """THINK 声明的整稿意图与逐页验收口径：交给 G-DECK 人工对照渲染总览，机器不评分。"""
    return {
        "visual_concept": plan.get("visual_concept"),
        "rhythm": plan.get("rhythm", []),
        "pages": [{"id": p["id"], "role": p["role"], "focus": p["focus"], "focus_element": p.get("focus_element"),
                   "focus_unmapped": p.get("focus_unmapped"),
                   "acceptance": p.get("acceptance", []), "constraints": p.get("constraints", [])} for p in plan["pages"]],
    }


_RENDER_TOOLS = ("pdftoppm", "pdftotext", "pdffonts")

_INSTALL_HINT = (
    "LibreOffice Impress（PPTX→PDF）：Windows `winget install TheDocumentFoundation.LibreOffice`，"
    "macOS `brew install --cask libreoffice`，Debian/Ubuntu `sudo apt install libreoffice-impress`。"
    "poppler（pdftoppm/pdftotext/pdffonts）：macOS `brew install poppler`，"
    "Debian/Ubuntu `sudo apt install poppler-utils`，Windows 用 `scoop install poppler` "
    "或 conda-forge 的 poppler，并把 bin 目录加入 PATH。"
)


def _soffice_path() -> Optional[str]:
    return shutil.which("soffice") or shutil.which("libreoffice")


def _missing_render_tools() -> List[str]:
    """真实预览所需工具的逐项探测；缺哪一项就单独点名，不笼统说「环境不可用」。"""
    missing = [] if _soffice_path() else ["LibreOffice（soffice）"]
    missing.extend(name for name in _RENDER_TOOLS if not shutil.which(name))
    return missing


def _render_toolchain_hint() -> str:
    """G-PRE 无法执行时给出能照做的下一步，而不是只留一个 NOT_FOUND。"""
    missing = _missing_render_tools()
    if not missing:
        return "渲染工具链齐备：给 --render-dir 一个空目录即可产出 G-PRE 证据"
    return f"缺少 {'、'.join(missing)}。{_INSTALL_HINT}装好后重跑，并给 --render-dir 一个空目录。"


def _environment_report(config: Dict[str, Any]) -> Dict[str, Any]:
    """把环境探测写进报告：缺什么、因此哪道门无法完成，读者不必再猜。"""
    missing = _missing_render_tools()
    return {
        "python": sys.version.split()[0],
        "python-pptx": importlib.metadata.version("python-pptx"),
        "libreoffice": _soffice_path() or "NOT_FOUND",
        "render_tools": {name: (shutil.which(name) or "NOT_FOUND") for name in _RENDER_TOOLS},
        "render_ready": not missing,
        "missing_render_tools": missing,
        "target_powerpoint": config.get("target_powerpoint", {}),
    }


@dataclass
class _VerifyState:
    """verify() 的中间状态。

    拆成几段之后，这十几项结果需要在段与段之间传递。用一个显式状态对象代替长参数表
    和隐式字典键：某一段忘了赋值会立刻变成 AttributeError，而不是悄悄生成一份错报告。
    """

    plan_path: str
    pptx: Path
    config: Dict[str, Any]
    config_file: Optional[Path]
    config_issues: List[str]
    plan: Optional[Dict[str, Any]]
    errors: List[CompileError]
    contract_blocked: bool
    presentation: Any = None
    parse_error: Optional[str] = None
    gates: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    render_facts: Dict[str, Any] = field(default_factory=dict)
    object_issues: List[str] = field(default_factory=list)
    object_evidence: Dict[str, Any] = field(default_factory=dict)
    sampling: Dict[str, Any] = field(default_factory=lambda: {"configuration_issues": []})
    blocking: Dict[str, List[str]] = field(default_factory=dict)
    pending_compliance: List[str] = field(default_factory=list)
    reviewed: List[Dict[str, Any]] = field(default_factory=list)
    review_binding: Dict[str, Any] = field(default_factory=dict)
    risk_records: List[Dict[str, Any]] = field(default_factory=list)
    risk_review: List[Dict[str, Any]] = field(default_factory=list)


def _prepare_verification(plan_path: str, pptx_path: str, config_path: Optional[str],
                          render_dir: Optional[str], asset_root: Optional[str]) -> _VerifyState:
    """读入 plan 与配置、跑契约检查、打开 PPTX、按需渲染。

    G-PRE 的每一种「没能渲染」都必须落成明确的门状态：未要求、契约未通过、文件打不开、
    工具缺失。渲染器自身的异常也只能变成 BLOCK，不能把整份报告吞掉。
    """
    pptx = Path(pptx_path).resolve()
    if asset_root is None:
        plan_dir = Path(plan_path).resolve().parent
        preferred = plan_dir / "assets"
        asset_root = str(preferred if preferred.is_dir() else plan_dir)
    config, config_file, config_issues = _load_config(config_path)
    plan, load_errors = load_plan(plan_path)
    errors: List[CompileError] = list(load_errors)
    if plan is not None:
        errors.extend(validate_plan(plan, asset_root))
    contract_blocked = plan is None or any(e.level == ErrorLevel.HARD for e in errors)
    state = _VerifyState(
        plan_path=plan_path, pptx=pptx, config=config, config_file=config_file,
        config_issues=config_issues, plan=plan, errors=errors, contract_blocked=contract_blocked,
    )
    if not contract_blocked and pptx.is_file():
        try:
            state.presentation = Presentation(str(pptx))
        except Exception as exc:
            state.parse_error = f"PPTX 无法打开：{type(exc).__name__}: {exc}"

    if not render_dir:
        state.gates["G-PRE"] = _gate("NOT VERIFIED", issues=[
            "未要求从本次 PPTX 生成渲染预览（未提供 --render-dir）",
            _render_toolchain_hint(),
        ])
    elif contract_blocked:
        state.gates["G-PRE"] = _gate("NOT VERIFIED", issues=["plan 契约未通过，未渲染可能过期的 PPTX"])
    elif state.presentation is None:
        state.gates["G-PRE"] = _gate("BLOCK", issues=[state.parse_error or f"PPTX 不存在：{pptx}"])
    else:
        try:
            state.gates["G-PRE"], render_logs, state.render_facts = _render(
                pptx, Path(render_dir).resolve(), plan, state.presentation)
        except Exception as exc:  # 渲染或解析异常不得让 CLI 丢失质量门报告
            state.gates["G-PRE"] = _gate("BLOCK", issues=[f"渲染执行失败：{type(exc).__name__}: {exc}"])
            render_logs = []
        if render_logs:
            state.gates["G-PRE"]["logs"] = render_logs
        if state.gates["G-PRE"]["status"] != "PASS":
            hint = _render_toolchain_hint()
            issues = state.gates["G-PRE"].setdefault("issues", [])
            if hint not in issues:
                issues.append(hint)
    state.errors.extend(state.render_facts.get("page_number_errors", []))
    return state


def _audit_objects(state: _VerifyState) -> None:
    """对象核对 → 技术诊断重算 → 页码 → 文字对比 → 覆盖度与抽样。

    每一段单独兜底：核验程序自己失败时必须记 HARD，绝不允许因为「没有结论」而被当成
    「没有问题」。这是整个报告里最需要 fail-closed 的地方。
    """
    if state.contract_blocked or state.presentation is None:
        return
    plan = state.plan
    audit_plan = _with_page_numbers(plan)  # 只给核对用：派生页码纳入同一套对象/对比度解析
    try:
        state.object_issues, state.object_evidence = audit_pptx(audit_plan, state.pptx, state.presentation)
    except Exception as exc:  # 核验程序失败：如实阻断，不推定通过
        state.object_issues = [f"实际对象核验失败（不得推定通过）：{type(exc).__name__}: {exc}"]
    try:
        state.errors.extend(_technical_diagnostics(plan, state.presentation))
    except Exception as exc:
        state.errors.append(CompileError(
            "E012", ErrorLevel.HARD, "deck.technical_diagnostics",
            "图片 ppi/裁切与图表配色诊断可由实际 PPTX 重算",
            f"{type(exc).__name__}: {exc}",
            "技术诊断无法重算，不能推定编译期没有同类问题；修复核验后重跑",
        ))
    for image in state.object_evidence.get("generated_images_checked", []):
        if not image["matches"]:
            state.errors.append(CompileError(
                "E011", ErrorLevel.HARD, image["object"], "生成素材登记哈希必须与 PPTX 内嵌图片一致",
                f"plan={image['expected_sha256']}，pptx={image['sha256']}",
                "登记文件与实际置入素材不一致：确认来源并重新登记，不得推定合规"))
    try:
        pn_errors, state.object_evidence["page_numbers"] = _page_number_audit(
            plan, state.pptx, state.presentation)
    except Exception as exc:  # 核验程序失败不能视为页码正确
        pn_errors = [CompileError("E017", ErrorLevel.HARD, "deck.page_numbers", "页码对象核验可执行",
                                  f"页码核验失败：{type(exc).__name__}: {exc}", "修复核验或输入后重跑")]
        state.object_evidence["page_numbers"] = []
    state.errors.extend(pn_errors)
    try:
        contrast_errors, state.object_evidence["text_contrast"] = _text_contrast(
            audit_plan, state.pptx, state.render_facts.get("pages"),
            state.render_facts.get("confirmed_pages"), state.presentation)
    except Exception as exc:  # 未完成对比核验不等于对比合格
        contrast_errors = [CompileError("E013", ErrorLevel.HARD, "deck.text_contrast", "文字对比核验可执行",
                                        f"对比核验失败：{type(exc).__name__}: {exc}", "修复核验或输入后重跑")]
        state.object_evidence["text_contrast"] = []
    state.errors.extend(contrast_errors)
    if state.render_facts.get("pages"):
        state.errors = _supersede_estimates(state.errors, state.render_facts.get("overflow_measured", {}))
    state.object_evidence["focus_coverage"] = _focus_coverage(plan)
    state.sampling = _sampling(plan, state.config)


def _assemble_machine_gates(state: _VerifyState) -> None:
    """把诊断按 policy 路由到各道门，并严格区分「确定缺陷」与「未能核验」。

    PASS 只覆盖真正执行并通过的范围：没有 PPTX 时 G-STR 是 NOT VERIFIED 而不是 PASS，
    G-SEM 在没有字形定位时把未核验项单独列进 not_verified。
    """
    gates = state.gates
    plan = state.plan
    blocking, state.pending_compliance, state.reviewed = _route_diagnostics(
        state.errors, None if state.contract_blocked else plan)
    state.blocking = blocking
    # E018 由 _render 在真实词框上产生，归到 G-PRE；不能因为 G-PRE 原本 PASS 就被忽略。
    if blocking["G-PRE"]:
        gates["G-PRE"]["issues"] = list(dict.fromkeys(
            list(gates["G-PRE"].get("issues", [])) + blocking["G-PRE"]))
        gates["G-PRE"]["status"] = "BLOCK"

    structural = blocking["G-STR"] + state.config_issues + state.sampling["configuration_issues"]
    str_status = "BLOCK" if structural else ("NOT VERIFIED" if state.presentation is None else "PASS")
    if str_status == "NOT VERIFIED":
        structural.append("实际 PPTX 不可用，页码对象契约 E017 未核对")
    gates["G-STR"] = _gate(str_status, [str(Path(state.plan_path).resolve())], structural,
                           scope="Schema 与契约硬约束（含页码对象契约 E017）",
                           page_numbers=state.object_evidence.get("page_numbers", []))

    if state.contract_blocked or state.presentation is None:
        reason = ("契约未通过，不核对可能过期的 PPTX" if state.contract_blocked
                  else (state.parse_error or f"PPTX 不存在：{state.pptx}"))
        gates["G-SEM"] = _gate("BLOCK", issues=[reason] + blocking["G-SEM"])
    else:
        sem_issues = state.object_issues + blocking["G-SEM"]
        # 无法定位字形的对比与确定缺陷分开，不得冒充 PASS。
        sem_unverified = [f"[{r['page']}.{r['text']}] 墨色 {r['ink']} 对比未核验：{r['reason']}"
                          f"（框内 {r['contrast_min']}–{r['contrast_max']}:1）"
                          for r in state.object_evidence.get("text_contrast", []) if r["status"] == "NOT VERIFIED"]
        gates["G-SEM"] = _gate("BLOCK" if sem_issues else ("NOT VERIFIED" if sem_unverified else "PASS"),
                               [str(state.pptx)], sem_issues or sem_unverified, not_verified=sem_unverified,
                               scope="plan↔PPTX 对象、数据语义筛查、文字对比；视觉主角及阅读路径均由 G-DECK 判断",
                               **state.object_evidence)

    state.review_binding = {
        "pptx_sha256": _sha256(state.pptx) if state.pptx.is_file() else None,
        "evidence_root": state.config_file.parent if state.config_file else Path.cwd(),
    }
    gates["G-DECK"] = _review_gate(state.config, "G-DECK", **state.review_binding)
    if blocking["G-DECK"]:
        gates["G-DECK"] = _gate("BLOCK", issues=blocking["G-DECK"])
    if not state.contract_blocked:
        gates["G-DECK"]["declared_intent"] = _declared_intent(plan)
        gates["G-DECK"]["contact_sheet"] = state.render_facts.get("contact_sheet")
        if "focus_coverage" in state.object_evidence:
            gates["G-DECK"]["focus_human_required"] = state.object_evidence["focus_coverage"]["human_required"]
    gates["G-TGT"] = _review_gate(state.config, "G-TGT", target=True, **state.review_binding)

    if blocking["G-COMPLY"]:
        gates["G-COMPLY"] = _gate("BLOCK", issues=blocking["G-COMPLY"])
    elif state.pending_compliance:
        gates["G-COMPLY"] = _gate("NOT VERIFIED", issues=state.pending_compliance)
    elif state.contract_blocked:
        gates["G-COMPLY"] = _gate("NOT VERIFIED", issues=["plan 无效，素材检查未执行"])
    else:
        assets = plan.get("assets", {})
        images = [(p["id"], el) for p in plan["pages"] for el in p["elements"] if el["type"] == "image"]
        used = sorted({el["image_id"] for _pid, el in images})
        required_images = {f"{pid}.{el['id']}" for pid, el in images
                           if (assets[el["image_id"]].get("image_generation") or {}).get("status") == "generated"}
        checked_images = {r["object"] for r in state.object_evidence.get("generated_images_checked", [])}
        missing_images = sorted(required_images - checked_images)
        gates["G-COMPLY"] = _gate(
            "NOT VERIFIED" if missing_images else "PASS", [str(Path(state.plan_path).resolve())],
            [f"生成素材未能核对 PPTX 内嵌哈希：{missing_images}"] if missing_images else [],
            references={image_id: assets[image_id].get("reference") for image_id in used},
            scope=f"声明、凭据位置、证据用途及已生成素材内嵌哈希；图片 {len(used)} 项，不冒充外部法律审计")


def _release_decision(state: _VerifyState) -> Dict[str, Any]:
    """发布判定：强制门未通过即阻断；可选门未验证时要求逐项具名覆盖。"""
    gates = state.gates
    required = set(state.config.get("required_gates", []))
    for gate_id, gate in gates.items():
        gate["mandatory"] = gate_id in required
        gate["verifier"] = GATE_VERIFIER[gate_id]
    state.risk_records, risk_issues, state.risk_review = _risk_records(
        state.config, gates, required, state.review_binding["pptx_sha256"],
        None if state.contract_blocked else state.plan)
    blocking_reasons = [f"{gid}={gate['status']}" for gid, gate in gates.items()
                        if gate["status"] == "BLOCK" or (gate["mandatory"] and gate["status"] == "NOT VERIFIED")]
    blocking_reasons.extend(risk_issues)
    uncovered: List[str] = []
    release_policy = POLICY["release"]
    if blocking_reasons:
        release_status = release_policy["mandatory_gate_not_verified"]
    else:
        # 每个未验证项都要被一条**有效**记录点名覆盖；同门的其他问题没有被点名就仍然未验证。
        optional_unverified = [gid for gid, gate in gates.items()
                               if not gate["mandatory"] and gate["status"] == "NOT VERIFIED"]
        for gid in optional_unverified:
            valid = [entry for entry in state.risk_review if entry["valid"] and entry["gate"] == gid]
            for key in _gate_issue_keys(gates[gid]):
                if not any(key["locator"] in entry["covers"] for entry in valid):
                    uncovered.append(f"{gid}: {key['locator']}")
        release_status = (release_policy["optional_gate_not_verified"] if uncovered
                          else ("PASS — Accepted Risk" if optional_unverified else "PASS — Delivery"))
    return {"status": release_status, "blocking_reasons": blocking_reasons, "uncovered_risks": uncovered}


def _build_report(state: _VerifyState, release: Dict[str, Any]) -> Dict[str, Any]:
    """组装报告：所有哈希、版本与门状态都取自本次实际产物，不读取编译期缓存。"""
    design_reviews = [] if state.contract_blocked else state.plan.get("design_reviews", [])
    matched = {(e.code, e.location) for e in state.errors}
    return {
        "report_version": REPORT_VERSION,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "project": state.config.get("project"),
        "plan": str(Path(state.plan_path).resolve()),
        "plan_sha256": _sha256(Path(state.plan_path)) if Path(state.plan_path).is_file() else None,
        "pptx": str(state.pptx),
        "pptx_sha256": state.review_binding["pptx_sha256"],
        "config": str(state.config_file) if state.config_file else None,
        "config_sha256": _sha256(state.config_file) if state.config_file and state.config_file.is_file() else None,
        "policy_version": POLICY.get("policy_version"),
        "policy_sha256": _sha256(Path(__file__).resolve().with_name("policy.json")),
        "environment": _environment_report(state.config),
        "diagnostics": [_error_dict(e) for e in state.errors],
        "reviewed_diagnostics": state.reviewed,
        "unmatched_design_reviews": [r for r in design_reviews if (r["code"], r["location"]) not in matched],
        "gates": state.gates,
        "sampling": state.sampling,
        "risk_approvals": state.risk_records,
        "risk_approval_review": state.risk_review,
        "release": release,
    }


def verify(plan_path: str, pptx_path: str, config_path: Optional[str] = None,
           render_dir: Optional[str] = None, asset_root: Optional[str] = None) -> Dict[str, Any]:
    """对照 plan、PPTX 与本次渲染，产出六门报告。四步各自可读，失败一律落成门状态。"""
    state = _prepare_verification(plan_path, pptx_path, config_path, render_dir, asset_root)
    _audit_objects(state)
    _assemble_machine_gates(state)
    return _build_report(state, _release_decision(state))

def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="核验 plan、实际 PPTX、真实渲染与交付质量门")
    parser.add_argument("plan")
    parser.add_argument("pptx")
    parser.add_argument("--config", help="项目验收配置 JSON")
    parser.add_argument("--render-dir", help="空目录；从本次 PPTX 重新生成 PDF/PNG 证据")
    parser.add_argument("--report", help="质量门 JSON 报告输出路径")
    parser.add_argument("--assets", default=None,
                        help="素材根目录；默认取 plan 所在目录或其 assets/，用于核对生成图登记文件与哈希")
    args = parser.parse_args(argv)
    if args.report:
        report_path = Path(args.report).resolve()
        protected = [Path(args.plan), Path(args.pptx)]
        if args.config:
            config_file = Path(args.config).resolve()
            protected.append(config_file)
            try:
                raw_config = json.loads(config_file.read_text(encoding="utf-8-sig"))
                for review in (raw_config.get("reviews") or {}).values():
                    evidence = (review or {}).get("evidence")
                    if evidence:
                        evidence_path = Path(str(evidence))
                        protected.append(evidence_path if evidence_path.is_absolute()
                                         else config_file.parent / evidence_path)
            except (OSError, UnicodeDecodeError, json.JSONDecodeError, AttributeError):
                pass
        for source in protected:
            try:
                same = report_path == source.resolve()
            except OSError:
                same = False
            if same:
                print(f"[拒绝] 报告输出不得覆盖输入文件：{source}", file=sys.stderr)
                return 1
    report = verify(args.plan, args.pptx, args.config, args.render_dir, args.assets)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        _atomic_write_text(Path(args.report), text + "\n")
    status = report["release"]["status"]
    print(status)
    for gate_id, gate in report["gates"].items():
        print(f"{gate_id}: {gate['status']} ({'mandatory' if gate['mandatory'] else 'optional'})")
    if args.report:
        print(f"报告：{args.report}")
    return 0 if status.startswith("PASS") else (2 if status == "NOT VERIFIED" else 1)


if __name__ == "__main__":
    sys.exit(main())
