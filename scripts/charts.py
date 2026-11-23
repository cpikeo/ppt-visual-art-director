#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""原生图表与表格的构建（python-pptx + OOXML 补丁）。"""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from lxml import etree

from model import (LIMITS, PT_EMU, chart_color_map, chart_items, contrast_ratio,
                   explicit_box, is_dark, mix_toward, page_bg, qn)
from ooxml import _ensure_family, _set_run_font, _strip_style


def _deemphasize(bg: str, kind: str) -> str:
    """退隐色仍要读得出来：往远离底色的方向走，够对比就停。往底色里混只会更低。"""
    base = LIMITS["chart"]["deemphasis"]["dark" if is_dark(bg) else "light"]
    floor = LIMITS["chart"]["min_contrast"].get(kind, 2.0)
    target = "#FFFFFF" if is_dark(bg) else "#000000"
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
    bg = page_bg(page)
    style = LIMITS["chart"]["style"]["dark" if is_dark(bg) else "light"]
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
    plot = obj.plots[0]
    combo = set(chart.get("combo_line_series") or [])
    stacked = bool(chart.get("stacked"))

    if kind in ("pie", "donut"):
        obj.has_legend = bool(chart.get("legend", False))
        if obj.has_legend:
            _style_legend(obj, style, family)
        for j, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = RGBColor.from_string(colors[j % len(colors)].lstrip("#"))
            point.format.line.color.rgb = RGBColor.from_string(bg.lstrip("#"))
    elif kind == "scatter":
        from pptx.enum.chart import XL_MARKER_STYLE
        obj.has_legend = bool(chart.get("legend", len(series) > 1))
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
        obj.has_legend = bool(chart.get("legend", len(series) > 1))
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

    label_format = _label_format(chart)
    if kind in ("pie", "donut") and chart.get("data_labels"):
        _pie_labels(plot, chart, style, family, bg)
    elif chart.get("data_labels") and kind == "scatter":
        _add_dlbls(plot._element, style, label_format)
    elif chart.get("data_labels"):
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.number_format = label_format
        labels.number_format_is_linked = False
        labels.font.size = Pt(LIMITS["chart"]["font_pt"]["label"])
        _ensure_family(labels.font, family)
        labels.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))
        labels.position = _label_position(kind, stacked)
        _patch_dlbls(plot)
        if stacked and kind == "bar":
            # 堆叠柱的标签落在色块里：墨色必须按各系列自己的底色取，否则深色块上的深字读不出。
            for i, ser in enumerate(plot.series):
                sl = ser.data_labels
                sl.show_value = True
                sl.number_format = label_format
                sl.number_format_is_linked = False
                sl.font.size = Pt(LIMITS["chart"]["font_pt"]["label"])
                _ensure_family(sl.font, family)
                sl.font.color.rgb = RGBColor.from_string(_best_ink(colors[i], style, bg).lstrip("#"))
                sl.position = _label_position(kind, stacked)

    if combo:
        line_plot = _split_combo(obj, combo, names, colors)
        if line_plot is not None and chart.get("combo_line_labels"):
            _add_dlbls(line_plot, style, label_format, "t")

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

def _label_position(kind: str, stacked: bool):
    from pptx.enum.chart import XL_LABEL_POSITION as P
    if kind in ("pie", "donut"):
        return P.BEST_FIT
    if kind == "line":
        return P.ABOVE
    if kind == "scatter":
        return P.ABOVE
    if stacked:
        return P.CENTER
    return P.OUTSIDE_END

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
        try:
            axis.has_title = True
            title_frame = axis.axis_title.text_frame
            title_frame.text = str(text)
            for para in title_frame.paragraphs:
                for run in para.runs:
                    _set_run_font(run, family, LIMITS["chart"]["font_pt"]["legend"], style["body"], False)
        except Exception:
            pass

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
    unit = str(chart.get("unit") or "").strip()
    if chart.get("stacked") == "percent":
        return '0"%"'
    if not unit or "%" in base:
        return base
    if unit == "%":
        return base + '"%"'
    return base + '" ' + unit + '"'  # 长单位也要写进直标；不再静默丢弃

def _pie_labels(plot, chart: Dict[str, Any], style: Dict[str, str], family: str, bg: str) -> None:
    """饼与环的直标：饼图放到扇区外并带类别名；环图放在环带内、只留数值并按扇区明度换墨色。

    饼的薄片放不下字，所以外置；环带够宽，标签落在带内就得按带色取墨，
    否则深底上压深字——这是渲染器无关的读法。
    """
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_LABEL_POSITION

    kind = chart["kind"]
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.number_format = _label_format(chart)
    labels.number_format_is_linked = False
    labels.font.size = Pt(LIMITS["chart"]["font_pt"]["label"])
    _ensure_family(labels.font, family)
    labels.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))
    labels.position = XL_LABEL_POSITION.OUTSIDE_END if kind == "pie" else XL_LABEL_POSITION.INSIDE_END
    _patch_dlbls(plot, show_cat=(kind == "pie"), separator=" ")

    if kind != "donut":
        return
    fmt = _label_format(chart)
    for point in plot.series[0].points:
        try:
            slice_hex = "#" + str(point.format.fill.fore_color.rgb).upper()
        except (AttributeError, TypeError, ValueError):
            continue
        ink = _best_ink(slice_hex, style, bg)
        dl = point.data_label
        dl.font.size = Pt(LIMITS["chart"]["font_pt"]["label"])
        _ensure_family(dl.font, family)
        dl.font.color.rgb = RGBColor.from_string(ink.lstrip("#"))
        _point_numfmt(point, fmt)  # 必须写在字体之后：dLbl 是惰性生成的

_DLBL_ORDER = ("c:idx", "c:delete", "c:layout", "c:tx", "c:numFmt", "c:spPr", "c:txPr",
               "c:dLblPos", "c:showLegendKey", "c:showVal", "c:showCatName", "c:showSerName",
               "c:showPercent", "c:showBubbleSize", "c:separator", "c:extLst")

def _insert_ordered(parent, node, order=_DLBL_ORDER) -> None:
    """按 schema 顺序插入。dLbl 与 dLbls 共用一张顺序表：两者共有的子元素相对次序一致，
    写错了 PowerPoint 会拒绝打开——这是直标相关代码唯一的顺序来源。"""
    rank = order.index("c:" + etree.QName(node).localname)
    for i, child in enumerate(parent):
        tag = "c:" + etree.QName(child).localname
        if tag in order and order.index(tag) > rank:
            parent.insert(i, node)
            return
    parent.append(node)

def _dlbl_flag(dlbls, tag: str, on: bool) -> None:
    node = dlbls.find(qn(tag))
    if node is None:
        node = etree.Element(qn(tag))
        node.set("val", "1" if on else "0")
        _insert_ordered(dlbls, node)
    else:
        node.set("val", "1" if on else "0")

def _patch_dlbls(plot, *, show_cat: Optional[bool] = None, separator: Optional[str] = None) -> None:
    """补正 python-pptx 写出的 dLbls：环图默认 showVal=0，等于声明了直标却一个数都不显示。"""
    dlbls = plot._element.find(qn("c:dLbls"))
    if dlbls is None:
        return
    _dlbl_flag(dlbls, "c:showVal", True)
    if show_cat is not None:
        _dlbl_flag(dlbls, "c:showCatName", show_cat)
    if separator is not None:
        sep = dlbls.find(qn("c:separator"))
        if sep is None:
            sep = etree.Element(qn("c:separator"))
            _insert_ordered(dlbls, sep)
        sep.text = separator

def _point_numfmt(point: Any, fmt: str) -> None:
    """逐点 dLbl 一旦存在就不再继承 plot 级 numFmt，单位会静默丢掉（环图上只剩“47”）。

    python-pptx 没有开放的入口：DataLabel 没有 number_format，而 DataLabel._element
    指向的是 c:ser 而不是 c:dLbl——直接往那儿写会造出 PowerPoint 报修的顺序错误。
    所以按 idx 找到真正的 c:dLbl，再按 CT_DLbl 的子元素顺序插 c:numFmt。
    """
    dlbls = point._ser.find(qn("c:dLbls"))
    if dlbls is None:
        return
    for d in dlbls.findall(qn("c:dLbl")):
        hit = d.find(qn("c:idx"))
        if hit is None or hit.get("val") != str(point._idx):
            continue
        node = d.find(qn("c:numFmt"))
        if node is None:
            node = etree.Element(qn("c:numFmt"))
            _insert_ordered(d, node)
        node.set("formatCode", fmt)
        node.set("sourceLinked", "0")
        return

def _best_ink(surface: str, style: Dict[str, str], bg: str) -> str:
    """在给定表面上，深墨与白墨哪个对比更高就用哪个。只在这两档之间选，不发明第三种颜色。"""
    # 深墨必须真的是深色：深底样式里 style["title"] 是浅色，旧写法在荧光/白色色块上两个候选都是浅字（读不出）。
    dark, light = LIMITS["chart"]["style"]["light"]["title"], "#FFFFFF"
    return dark if contrast_ratio(dark, surface) >= contrast_ratio(light, surface) else light

def _add_dlbls(plot_element, style: Dict[str, str], number_format: str, position: Optional[str] = None) -> None:
    """python-pptx 没有散点/折线混排的直标入口，整块直接写 XML。"""
    for old in plot_element.findall(qn("c:dLbls")):
        plot_element.remove(old)
    labels = etree.Element(qn("c:dLbls"))
    fmt = etree.SubElement(labels, qn("c:numFmt"))
    fmt.set("formatCode", number_format)
    fmt.set("sourceLinked", "0")
    txPr = etree.SubElement(labels, qn("c:txPr"))
    etree.SubElement(txPr, qn("a:bodyPr"))
    etree.SubElement(txPr, qn("a:lstStyle"))
    para = etree.SubElement(txPr, qn("a:p"))
    pPr = etree.SubElement(para, qn("a:pPr"))
    def_rpr = etree.SubElement(pPr, qn("a:defRPr"))
    def_rpr.set("sz", str(int(LIMITS["chart"]["font_pt"]["label"] * 100)))
    fill = etree.SubElement(def_rpr, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:srgbClr")).set("val", style["body"].lstrip("#").upper())
    etree.SubElement(para, qn("a:endParaRPr")).set("lang", "zh-CN")
    if position:
        etree.SubElement(labels, qn("c:dLblPos")).set("val", position)
    # 只写 showVal 时，部分渲染器会把类别名与系列名一并显示；把其余开关显式关掉。
    etree.SubElement(labels, qn("c:showLegendKey")).set("val", "0")
    etree.SubElement(labels, qn("c:showVal")).set("val", "1")
    for tag in ("c:showCatName", "c:showSerName", "c:showPercent", "c:showBubbleSize"):
        etree.SubElement(labels, qn(tag)).set("val", "0")

    anchor = None
    for tag in ("c:dropLines", "c:hiLowLines", "c:upDownBars", "c:marker", "c:smooth", "c:axId"):
        anchor = plot_element.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(labels)
    else:
        plot_element.append(labels)

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
    unit = str(spec.get("unit") or "").strip()
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
    style = LIMITS["chart"]["style"]["dark" if is_dark(page_bg(page)) else "light"]
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
