#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""原生图表与表格的构建（python-pptx + OOXML 补丁）。"""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from lxml import etree

from model import (LIMITS, chart_flags, PT_EMU, chart_color_map, chart_items, chart_label_plan, chart_mode,
                   chart_style, chart_surface, chart_unit, contrast_ratio, explicit_box, mix_toward, qn)
from ooxml import _ensure_family, _set_run_font, _strip_style


def _deemphasize(bg: str, kind: str) -> str:
    """退隐色仍要读得出来：往远离承载面的方向走，够对比就停。往底色里混只会更低。"""
    dark = chart_mode(bg) == "dark"
    base = LIMITS["chart"]["deemphasis"]["dark" if dark else "light"]
    floor = LIMITS["chart"]["min_contrast"].get(kind, 2.0)
    target = "#FFFFFF" if dark else "#000000"
    candidate = base
    for step in range(0, 25):
        candidate = mix_toward(base, target, step / 24 * 0.45)
        if contrast_ratio(candidate, bg) >= floor:
            return candidate
    return candidate

def _left_align_chart_title(obj) -> None:
    """把图表标题框固定到图表区左上：默认 auto layout 会把它居中，与左对齐的轴脱节。"""
    title = obj._chartSpace.find(qn("c:chart")).find(qn("c:title"))
    layout = None if title is None else title.find(qn("c:layout"))
    if layout is None:
        return
    for child in list(layout):
        layout.remove(child)
    manual = etree.SubElement(layout, qn("c:manualLayout"))
    for tag, val in (("c:layoutTarget", "inner"), ("c:xMode", "edge"), ("c:yMode", "edge")):
        etree.SubElement(manual, qn(tag)).set("val", val)
    for tag, val in (("c:x", "0"), ("c:y", "0.02"), ("c:w", "1"), ("c:h", "0.14")):
        etree.SubElement(manual, qn(tag)).set("val", val)

def _flip_category_axis(obj) -> None:
    """横向条形图默认把第一个类别放在最下面；阅读顺序应从数据的第一项开始。"""
    cat = obj._chartSpace.find(".//" + qn("c:catAx"))
    if cat is None:
        return
    scaling = cat.find(qn("c:scaling"))
    if scaling is None:
        return
    for old in scaling.findall(qn("c:orientation")):
        scaling.remove(old)
    node = etree.Element(qn("c:orientation"))
    node.set("val", "maxMin")
    scaling.insert(0, node)

def _chart_type(chart: Dict[str, Any]):
    from pptx.enum.chart import XL_CHART_TYPE as XL
    kind, stacked, horiz = chart["kind"], chart.get("stacked"), chart.get("horizontal")
    if kind == "bar":
        if stacked == "percent":
            return XL.BAR_STACKED_100 if horiz else XL.COLUMN_STACKED_100
        if stacked:
            return XL.BAR_STACKED if horiz else XL.COLUMN_STACKED
        return XL.BAR_CLUSTERED if horiz else XL.COLUMN_CLUSTERED
    if kind == "line":
        return XL.LINE
    if kind == "area":
        return XL.AREA_STACKED if stacked else XL.AREA
    if kind == "pie":
        return XL.PIE
    if kind == "donut":
        return XL.DOUGHNUT
    return XL.XY_SCATTER

def add_chart(slide, plan, page, el):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.chart.data import CategoryChartData, XyChartData

    chart = el["chart"]
    kind = chart["kind"]
    bg = chart_surface(page, el)  # 承载面：页面底色 + 其下整块覆盖的色块；全图的结构色与退隐色都从它推
    style = chart_style(bg)
    fonts = plan.get("fonts") or {}
    family = el.get("font") or fonts.get("body") or LIMITS["defaults"]["fonts"]["body"]
    box = explicit_box(el)

    if kind == "table":
        return add_table(slide, plan, page, el, style, family)

    if kind == "scatter":
        data = XyChartData()
        for s in chart["data"]["series"]:
            series = data.add_series(str(s["name"]))
            for x, y in s["points"]:
                series.add_data_point(float(x), float(y))
    else:
        data = CategoryChartData()
        data.categories = chart["data"]["categories"]
        for s in chart["data"]["series"]:
            data.add_series(str(s["name"]), [float(v) for v in s["values"]])

    frame = slide.shapes.add_chart(_chart_type(chart), Inches(box[0]), Inches(box[1]),
                                   Inches(box[2]), Inches(box[3]), data)
    frame.name = f"{page['id']}.{el['id']}"
    obj = frame.chart
    if kind == "scatter":
        _scatter_marker_only(obj)
    _ensure_family(obj.font, family)
    obj.font.size = Pt(LIMITS["chart"]["font_pt"]["axis"])

    series = chart["data"]["series"]
    names = [str(s["name"]) for s in series]
    # 颜色只有一处来源：model.chart_color_map。扇区、系列、自定义图例与核验器读同一份映射，
    # 手绘色块与扇区对不上这类错色从根上不可能发生。
    colors = chart_color_map(plan, chart)
    flags = chart_flags(chart)
    plot = obj.plots[0]
    combo = set(chart.get("combo_line_series") or [])
    stacked = bool(chart.get("stacked"))

    if kind in ("pie", "donut"):
        obj.has_legend = flags["legend"]
        if obj.has_legend:
            _style_legend(obj, style, family)
        for j, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = RGBColor.from_string(colors[j % len(colors)].lstrip("#"))
            point.format.line.color.rgb = RGBColor.from_string(bg.lstrip("#"))
    elif kind == "scatter":
        from pptx.enum.chart import XL_MARKER_STYLE
        obj.has_legend = flags["legend"]
        for i, s in enumerate(plot.series):
            color = RGBColor.from_string(colors[i].lstrip("#"))
            s.format.line.fill.background()  # 散点不连线：连线会暗示不存在的顺序
            s.marker.style = XL_MARKER_STYLE.CIRCLE
            s.marker.size = 7
            s.marker.format.fill.solid()
            s.marker.format.fill.fore_color.rgb = color
            s.marker.format.line.color.rgb = color
    else:
        from pptx.enum.chart import XL_MARKER_STYLE
        obj.has_legend = flags["legend"]
        for i, s in enumerate(plot.series):
            color = RGBColor.from_string(colors[i].lstrip("#"))
            if kind == "line":
                s.format.line.color.rgb = color
                s.format.line.width = Pt(2.0)
                s.smooth = False
                s.marker.style = XL_MARKER_STYLE.NONE
                continue
            if names[i] in combo:
                s.format.line.color.rgb = color
                s.format.line.width = Pt(2.5)
                s.smooth = False
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = color
                continue
            if kind == "area":
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = color
                s.format.line.color.rgb = RGBColor.from_string(bg.lstrip("#"))
                s.format.line.width = Pt(0.75)
                continue
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = color
            s.format.line.fill.background()

    if obj.has_legend:
        _style_legend(obj, style, family)

    if kind in ("bar", "line", "area") and stacked:
        plot.overlap = LIMITS["chart"]["stacked_overlap"]
    if kind == "bar":
        plot.gap_width = int(chart.get("gap_width") or LIMITS["chart"]["gap_width"])

    if kind == "bar" and chart.get("highlight_category") is not None:
        hi = int(chart["highlight_category"])
        dim = _deemphasize(bg, "bar")
        for j, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = RGBColor.from_string((colors[0] if j == hi else dim).lstrip("#"))
            point.format.line.fill.background()

    _style_axes(obj, chart, style, family, kind)
    if chart.get("horizontal") and kind == "bar":
        _flip_category_axis(obj)

    if flags["data_labels"]:
        _write_labels(plot, chart_label_plan(plan, chart, bg), chart, style, family, _label_format(chart))

    if combo:
        line_plot = _split_combo(obj, combo, names, colors)
        if line_plot is not None and chart.get("combo_line_labels"):
            plan_by_name = {e["name"]: e for e in chart_label_plan(plan, chart, bg)}
            for ser in line_plot.findall(qn("c:ser")):
                entry = plan_by_name.get(_ser_label(ser), {})
                _attach(ser, _build_dlbls(entry.get("ink", style["body"]), _label_format(chart), "t",
                                          plate=entry.get("plate")))

    title = chart.get("title")
    if chart.get("show_title", True) and title:
        from pptx.enum.text import PP_ALIGN
        obj.has_title = True
        chart_title = obj.chart_title
        chart_title.text_frame.text = title
        for para in chart_title.text_frame.paragraphs:
            para.alignment = PP_ALIGN.LEFT
            for run in para.runs:
                _set_run_font(run, family, LIMITS["chart"]["font_pt"]["title"], style["title"], False)
        _left_align_chart_title(obj)
    else:
        obj.has_title = False
    return frame

def _style_legend(obj, style, family) -> None:
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_LEGEND_POSITION
    obj.legend.position = XL_LEGEND_POSITION.BOTTOM
    obj.legend.include_in_layout = False
    obj.legend.font.size = Pt(LIMITS["chart"]["font_pt"]["legend"])
    _ensure_family(obj.legend.font, family)
    obj.legend.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))

def _style_axes(obj, chart, style, family, kind: str) -> None:
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_TICK_LABEL_POSITION, XL_TICK_MARK
    axis_on = chart.get("axis_labels", True)
    axis_cfg = chart.get("value_axis") or {}

    if kind not in ("pie", "donut"):
        va = obj.value_axis
        if kind == "scatter":
            va.has_major_gridlines = axis_on
        else:
            va.has_major_gridlines = axis_on
        if axis_on:
            va.major_gridlines.format.line.color.rgb = RGBColor.from_string(style["grid"].lstrip("#"))
            va.major_gridlines.format.line.width = Pt(0.75)
            va.format.line.color.rgb = RGBColor.from_string(style["rule"].lstrip("#"))
        else:
            va.visible = False
            va.has_major_gridlines = False
        va.tick_labels.font.size = Pt(LIMITS["chart"]["font_pt"]["axis"])
        _ensure_family(va.tick_labels.font, family)
        va.tick_labels.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))
        if axis_cfg.get("min") is not None:
            va.minimum_scale = float(axis_cfg["min"])
        if axis_cfg.get("max") is not None:
            va.maximum_scale = float(axis_cfg["max"])
        if axis_cfg.get("major_unit") is not None:
            va.major_unit = float(axis_cfg["major_unit"])
        if axis_cfg.get("number_format"):
            va.tick_labels.number_format = str(axis_cfg["number_format"])
            va.tick_labels.number_format_is_linked = False

        ca = obj.category_axis
        ca.has_major_gridlines = False
        ca.tick_labels.font.size = Pt(LIMITS["chart"]["font_pt"]["axis"])
        _ensure_family(ca.tick_labels.font, family)
        ca.tick_labels.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))
        ca.format.line.color.rgb = RGBColor.from_string(style["rule"].lstrip("#"))
        if kind == "bar" and chart.get("horizontal"):
            # 横条图的类别轴是左侧一条竖脊，只承载标签：留下脊线与刻度只会像残留物。
            ca.format.line.fill.background()
            ca.major_tick_mark = XL_TICK_MARK.NONE
            ca.minor_tick_mark = XL_TICK_MARK.NONE
        # 类别标签放到绘图区最低处：含负值时标签不会压在零线下的柱上。
        # 旧写法 ca.tick_labels.position 在 python-pptx 里不是属性，赋值被静默吞掉，XML 里始终是 nextTo。
        ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW

    titles = chart.get("axis_titles") or {}
    pairs = (("y", obj.value_axis), ("x", obj.category_axis)) if kind not in ("pie", "donut") else ()
    for key, axis in pairs:
        text = titles.get(key)
        if not text:
            continue
        axis.has_title = True
        title_frame = axis.axis_title.text_frame
        title_frame.text = str(text)
        for para in title_frame.paragraphs:
            for run in para.runs:
                _set_run_font(run, family, LIMITS["chart"]["font_pt"]["legend"], style["body"], False)

def _split_combo(obj, combo: set, names: List[str], colors: List[str]):
    """把指定系列从柱图移到折线图：两条 plot 共用同一数值轴。返回折线层元素，供直标使用。"""

    plot_area = obj._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))
    bar = plot_area.find(qn("c:barChart"))
    if bar is None:
        return
    line = etree.Element(qn("c:lineChart"))
    etree.SubElement(line, qn("c:grouping")).set("val", "standard")
    etree.SubElement(line, qn("c:varyColors")).set("val", "0")
    moved = 0
    for ser in list(bar.findall(qn("c:ser"))):
        order = ser.find(qn("c:order"))
        idx = int(order.get("val")) if order is not None and order.get("val") else 0
        if idx < len(names) and names[idx] in combo:
            bar.remove(ser)
            line.append(ser)
            moved += 1
    if not moved:
        return None
    etree.SubElement(line, qn("c:marker")).set("val", "1")
    for ser in line.findall(qn("c:ser")):
        _add_line_marker(ser, colors)
    for ax in bar.findall(qn("c:axId")):
        etree.SubElement(line, qn("c:axId")).set("val", ax.get("val"))
    bar.addnext(line)
    return line

def _scatter_marker_only(obj) -> None:
    """默认线型散点会把点连成线，暗示不存在的顺序；散点只留标记。"""
    style = obj._chartSpace.find(".//" + qn("c:scatterStyle"))
    if style is not None:
        style.set("val", "marker")

def _label_format(chart: Dict[str, Any]) -> str:
    """直标带单位：读者看到“52”要能知道是百分比还是亿元。"""
    base = (chart.get("value_axis") or {}).get("number_format") or "#,##0.###"
    unit = chart_unit(chart)
    if chart.get("stacked") == "percent":
        return '0"%"'
    if not unit or "%" in base:
        return base
    if unit == "%":
        return base + '"%"'
    return base + '" ' + unit + '"'  # 长单位也要写进直标；不再静默丢弃

def _ser_label(ser) -> str:
    node = ser.find(qn("c:tx")).find(".//" + qn("c:v")) if ser.find(qn("c:tx")) is not None else None
    return node.text if node is not None and node.text else ""

# c:ser 里 dLbls 之后的兄弟（CT_BarSer / CT_LineSer / CT_AreaSer / CT_PieSer / CT_ScatterSer 的并集）。
_AFTER_SER_DLBLS = ("trendline", "errBars", "cat", "val", "xVal", "yVal", "smooth", "shape", "extLst")

def _attach(ser, dlbls) -> None:
    """把 dLbls 插到 ser 里 schema 要求的位置（dPt 之后、cat/val 之前），并替换已有的。"""
    for old in ser.findall(qn("c:dLbls")):
        ser.remove(old)
    for child in ser:
        if etree.QName(child).localname in _AFTER_SER_DLBLS:
            child.addprevious(dlbls)
            return
    ser.append(dlbls)

def _txpr(ink: str, size_pt: float, family: str):
    txpr = etree.Element(qn("c:txPr"))
    etree.SubElement(txpr, qn("a:bodyPr"))
    etree.SubElement(txpr, qn("a:lstStyle"))
    para = etree.SubElement(txpr, qn("a:p"))
    ppr = etree.SubElement(para, qn("a:pPr"))
    rpr = etree.SubElement(ppr, qn("a:defRPr"))
    rpr.set("sz", str(int(size_pt * 100)))
    etree.SubElement(etree.SubElement(rpr, qn("a:solidFill")), qn("a:srgbClr")).set("val", ink.lstrip("#").upper())
    for tag in ("a:latin", "a:ea", "a:cs"):
        etree.SubElement(rpr, qn(tag)).set("typeface", family)
    etree.SubElement(para, qn("a:endParaRPr")).set("lang", "zh-CN")
    return txpr

def _label_body(parent, ink, fmt, position, family, plate=None, show_cat=False, separator=None) -> None:
    """dLbl 与 dLbls 共有的子元素，严格按 CT_DLbl / Group_DLbls 的顺序写：
    numFmt, spPr, txPr, dLblPos, showLegendKey, showVal, showCatName, showSerName,
    showPercent, showBubbleSize, separator。写错顺序 PowerPoint 会提示修复。"""
    node = etree.SubElement(parent, qn("c:numFmt"))
    node.set("formatCode", fmt)
    node.set("sourceLinked", "0")
    sp = etree.SubElement(parent, qn("c:spPr"))
    if plate:
        etree.SubElement(etree.SubElement(sp, qn("a:solidFill")), qn("a:srgbClr")).set("val", plate.lstrip("#").upper())
    else:
        etree.SubElement(sp, qn("a:noFill"))
    etree.SubElement(etree.SubElement(sp, qn("a:ln")), qn("a:noFill"))
    parent.append(_txpr(ink, LIMITS["chart"]["font_pt"]["label"], family))
    if position:
        etree.SubElement(parent, qn("c:dLblPos")).set("val", position)
    for tag, on in (("c:showLegendKey", False), ("c:showVal", True), ("c:showCatName", show_cat),
                    ("c:showSerName", False), ("c:showPercent", False), ("c:showBubbleSize", False)):
        etree.SubElement(parent, qn(tag)).set("val", "1" if on else "0")
    if separator is not None:
        etree.SubElement(parent, qn("c:separator")).text = separator

def _build_dlbls(ink, fmt, position, family="", plate=None, show_cat=False, separator=None,
                 point_inks=None, leader_lines=False):
    """整套直标的唯一写入口。直标一律写在**系列**上：每个系列自己的墨色与位置，
    不依赖 plot 级继承，也不依赖 python-pptx 惰性生成的 dLbl。"""
    dlbls = etree.Element(qn("c:dLbls"))
    for idx, point_ink in enumerate(point_inks or []):
        dl = etree.SubElement(dlbls, qn("c:dLbl"))
        etree.SubElement(dl, qn("c:idx")).set("val", str(idx))
        _label_body(dl, point_ink, fmt, position, family, plate, show_cat, separator)
    _label_body(dlbls, ink, fmt, position, family, plate, show_cat, separator)
    if leader_lines:
        etree.SubElement(dlbls, qn("c:showLeaderLines")).set("val", "1")
    return dlbls

def _write_labels(plot, entries, chart, style, family, fmt) -> None:
    kind = chart["kind"]
    sers = plot._element.findall(qn("c:ser"))
    combo = set(chart.get("combo_line_series") or [])
    for i, ser in enumerate(sers):
        entry = entries[0] if kind in ("pie", "donut") else (entries[i] if i < len(entries) else None)
        if entry is None or (kind == "bar" and entry["name"] in combo):
            continue  # 组合图的折线层直标由 combo_line_labels 单独决定
        _attach(ser, _build_dlbls(
            entry["ink"], fmt, entry["position"], family, entry["plate"],
            show_cat=(kind == "pie"), separator=" " if kind == "pie" else None,
            point_inks=entry["points"], leader_lines=(kind == "pie")))

def _add_line_marker(ser, colors: List[str]) -> None:
    """折线上的小圆点：让读者在类别轴上定位读数。marker 必须插在 cat/val 之前。"""
    order = ser.find(qn("c:order"))
    idx = int(order.get("val")) if order is not None and order.get("val") else 0
    color = str(colors[idx % len(colors)]).lstrip("#").upper()
    marker = etree.Element(qn("c:marker"))
    etree.SubElement(marker, qn("c:symbol")).set("val", "circle")
    etree.SubElement(marker, qn("c:size")).set("val", "6")
    spPr = etree.SubElement(marker, qn("c:spPr"))
    fill = etree.SubElement(spPr, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:srgbClr")).set("val", color)
    ln = etree.SubElement(spPr, qn("a:ln"))
    etree.SubElement(ln, qn("a:solidFill")).append(etree.Element(qn("a:srgbClr")))
    ln.find(".//" + qn("a:srgbClr")).set("val", color)
    anchor = ser.find(qn("c:cat"))
    if anchor is None:
        anchor = ser.find(qn("c:val"))
    if anchor is not None:
        anchor.addprevious(marker)
    else:
        ser.append(marker)

def _legend_value(spec: Dict[str, Any], value: Any) -> str:
    """图例标签里的读数：与直标同源的单位写法，只在饼/环上启用（柱线图的图例只列系列名）。"""
    unit = chart_unit(spec)
    if spec.get("stacked") == "percent":
        return f"{float(value):g}%"
    return f"{float(value):g}" + unit

def add_legend(slide, plan, page, el):
    """自定义图例：色块与标签由被引用图表的**同一份**颜色映射生成。

    手绘图例各抄各的色（有的抄整稿 palette、有的照记忆挑、有的按叙述顺序排），
    与扇区对不上就是这么来的。现在图例只是 chart_color_map 的一个视图：
    改 palette 或 color_map，扇区与色块同时变；核验器按同一份映射逐块比对。
    """
    from pptx.util import Inches, Pt

    ref = str(el.get("for_chart") or "")
    target = next((e for e in page.get("elements") or [] if str(e.get("id")) == ref), None)
    if target is None or target.get("type") != "chart":
        raise ValueError(f"legend「{el.get('id')}」引用的 chart「{ref}」不在本页顶层")
    spec = target["chart"]
    if spec.get("kind") == "table":
        raise ValueError(f"legend「{el.get('id')}」引用的是表格：表格没有图例，用表头与细线分组")
    items, colors = chart_items(spec), chart_color_map(plan, spec)
    if not items:
        raise ValueError(f"legend「{el.get('id')}」引用的图表没有可图例化的类别或系列")

    box = explicit_box(el)
    style = chart_style(chart_surface(page, el))  # 图例自己的承载面，不一定等于图表的
    ink = str(el.get("color") or style["body"]).upper()
    family = str(el.get("font") or (plan.get("fonts") or {}).get("body")
                 or LIMITS["defaults"]["fonts"]["body"])
    size = float(el.get("font_size") or 18)
    swatch = min(float(el.get("swatch_in") or 0.16), box[2], box[3])
    gap = 0.14
    show_value = bool(el.get("show_value")) and spec.get("kind") in ("pie", "donut")
    values = list(((spec.get("data") or {}).get("series") or [{}])[0].get("values") or [])
    children = []
    name = f"{page['id']}.{el['id']}"

    if str(el.get("direction") or "vertical") == "horizontal":
        col = box[2] / len(items)
        for i, item in enumerate(items):
            cx = box[0] + i * col
            children.append(_swatch(slide, name, i, cx, box[1] + box[3] / 2 - swatch / 2, swatch, colors[i]))
            children.append(_swatch_label(slide, family, name, i, _legend_label(spec, item, values, i, show_value),
                                          cx + swatch + gap, box[1], col - swatch - gap, box[3], size, ink,
                                          str(el.get("align") or "left")))
    else:
        row = box[3] / len(items)
        for i, item in enumerate(items):
            top = box[1] + i * row
            sy = top + max((row - swatch) / 2, 0)
            children.append(_swatch(slide, name, i, box[0], sy, swatch, colors[i]))
            children.append(_swatch_label(slide, family, name, i, _legend_label(spec, item, values, i, show_value),
                                          box[0] + swatch + gap, top, box[2] - swatch - gap, row, size, ink,
                                          str(el.get("align") or "left")))
    group = slide.shapes.add_group_shape([c for c in children if c is not None])
    group.name = name
    return group

def _legend_label(spec: Dict[str, Any], item: str, values: List[Any], i: int, show_value: bool) -> str:
    if show_value and i < len(values):
        return f"{item} · {_legend_value(spec, values[i])}"
    return str(item)

def _swatch(slide, name: str, idx: int, x: float, y: float, size: float, color: str):
    from pptx.util import Inches
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(size), Inches(size))
    shape.name = f"{name}#sw{idx}"
    _strip_style(shape)
    shape.shadow.inherit = False
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(str(color).lstrip("#").upper())
    shape.line.fill.background()
    if shape.has_text_frame:
        shape.text_frame.text = ""
    return shape

def _swatch_label(slide, family: str, name: str, idx: int, text: str, x: float, y: float, w: float, h: float,
                  size: float, ink: str, align: str):
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

    frame = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(max(w, 0.3)), Inches(h))
    frame.name = f"{name}#tx{idx}"
    tf = frame.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)
    para.line_spacing = 1.1
    run = para.add_run()
    run.text = text
    _set_run_font(run, family, size, ink, False)
    return frame

def add_table(slide, plan, page, el, style, family):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    chart = el["chart"]
    headers = chart["data"]["headers"]
    rows = chart["data"]["rows"]
    box = explicit_box(el)
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(box[0]), Inches(box[1]),
                                   Inches(box[2]), Inches(box[3]))
    shape.name = f"{page['id']}.{el['id']}"
    table = shape.table
    table.first_row = False
    table.horz_banding = False
    tblPr = table._tbl.find(qn("a:tblPr"))
    if tblPr is not None:
        old = tblPr.find(qn("a:tableStyleId"))
        if old is not None:
            tblPr.remove(old)
        etree.SubElement(tblPr, qn("a:tableStyleId")).text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid

    weights = chart.get("column_weights") or [1] * len(headers)
    total = sum(weights)
    for i, weight in enumerate(weights):
        table.columns[i].width = Inches(box[2] * weight / total)

    body_pt = LIMITS["chart"]["font_pt"]["table"]
    aligns = chart.get("column_align") or ["left"] * len(headers)
    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}

    def paint(cell, text, ink, bold, align):
        cell.text = str(text)
        cell.margin_left = cell.margin_right = Inches(0.08)
        cell.margin_top = cell.margin_bottom = Inches(0.03)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        for para in cell.text_frame.paragraphs:
            para.alignment = align_map[align]
            for run in para.runs:
                _set_run_font(run, family, body_pt, ink, bold)

    for j, head in enumerate(headers):
        paint(table.cell(0, j), head, style["title"], True, aligns[j])
    for r, row in enumerate(rows):
        for j, value in enumerate(row):
            paint(table.cell(r + 1, j), value, style["body"], False, aligns[j])

    # 细底线：比灰底表格干净，也更像编辑设计
    for r in range(len(rows) + 1):
        for c in range(len(headers)):
            tcPr = table.cell(r, c)._tc.get_or_add_tcPr()
            for tag in ("a:lnT", "a:lnB", "a:lnL", "a:lnR"):
                for node in tcPr.findall(qn(tag)):
                    tcPr.remove(node)
                if tag in ("a:lnT", "a:lnB"):
                    line = etree.SubElement(tcPr, qn(tag))
                    line.set("w", str(int(0.75 * PT_EMU)))
                    line.set("cmpd", "sng")
                    fill = etree.SubElement(line, qn("a:solidFill"))
                    etree.SubElement(fill, qn("a:srgbClr")).set(
                        "val", (style["rule"] if r in (0, 1) else style["grid"]).lstrip("#").upper())
    return shape
