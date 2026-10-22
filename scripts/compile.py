#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""compile.py · deck spec → 原生可编辑 PPTX（确定性）。

纪律：
  * 只做机械翻译：spec 里有什么就写什么；不做布局补偿、不猜单位、不改字号。
  * 文本写原生文本框（可编辑），图表写原生 chart part（可编辑数据），不用图片冒充。
  * 输出字节确定：ZIP 时间戳归零、条目排序、核心属性固定。同一 spec 两次编译字节相同。
  * 编译结果自带 readback 摘要（元素数、字号、颜色），供 fidelity 对账。
"""
from __future__ import annotations

import hashlib
import io
import re
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt
from pptx.oxml.ns import qn

COMPILER_VERSION = "1.0"
_ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)
_ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
_ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
_CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
}
_SHAPES = {
    "rect": MSO_SHAPE.RECTANGLE,
    "rounded_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "diamond": MSO_SHAPE.DIAMOND,
}


class CompileError(RuntimeError):
    pass


def _emu(px: float, dpi: float) -> Emu:
    return Emu(int(round(float(px) * 914400 / dpi)))


def _pt(px: float, dpi: float) -> Pt:
    return Pt(float(px) * 72.0 / dpi)


def _rgb(value, colors: dict) -> RGBColor:
    if isinstance(value, str) and value.startswith("#"):
        return RGBColor.from_string(value[1:].upper())
    if isinstance(value, str) and value in colors:
        return RGBColor.from_string(str(colors[value]).lstrip("#").upper())
    raise CompileError(f"无法解析颜色：{value!r}（只接受 #RRGGBB 或 theme token）")


def _set_fonts(run, cn: str, latin: str) -> None:
    """同时写 latin 与 East-Asian 字体：只写 latin 时中文会回落到系统默认字体。"""
    run.font.name = latin
    rpr = run.font._rPr
    for tag, face in (("a:ea", cn), ("a:cs", latin)):
        node = rpr.find(qn(tag))
        if node is None:
            node = rpr.makeelement(qn(tag), {})
            rpr.append(node)
        node.set("typeface", face)


def compile_deck(deck: dict, output: str | Path, base_dir=None) -> dict:
    dpi = float((deck.get("canvas") or {}).get("ppt_dpi") or 96)
    cw = float(deck["canvas"]["width"])
    ch = float(deck["canvas"]["height"])
    theme = deck.get("theme") or {}
    colors = theme.get("colors") or {}
    fonts = theme.get("fonts") or {}
    base_dir = base_dir or (deck.get("meta") or {}).get("base_dir") or None
    if base_dir:
        base_dir = str(base_dir)

    prs = Presentation()
    prs.slide_width = _emu(cw, dpi)
    prs.slide_height = _emu(ch, dpi)
    blank = prs.slide_layouts[6]
    report = {"compiler_version": COMPILER_VERSION, "slides": [], "counts": {}}
    counts = {"text": 0, "shape": 0, "chart": 0, "image": 0}

    for slide_spec in deck["slides"]:
        slide = prs.slides.add_slide(blank)
        slide._pvad_colors = colors   # 供 overlay 取色，不改对象模型
        _paint_background(slide, slide_spec, colors)
        readback = {"id": slide_spec["id"], "shapes": []}
        for element in slide_spec.get("elements") or []:
            etype = element["type"]
            counts[etype] = counts.get(etype, 0) + 1
            if etype == "text":
                _add_text(slide, element, dpi, colors, fonts)
            elif etype == "shape":
                _add_shape(slide, element, dpi, colors)
            elif etype == "chart":
                _add_chart(slide, element, dpi, colors, fonts)
            elif etype == "image":
                _add_image(slide, element, dpi, base_dir)
            else:
                raise CompileError(f"compile 不支持元素类型：{etype}")
            readback["shapes"].append({
                "id": element["id"], "type": etype,
                "x": float(element["x"]), "y": float(element["y"]),
                "width": float(element["width"]), "height": float(element["height"]),
            })
        report["slides"].append(readback)
    report["counts"] = counts

    target = Path(output)
    target.parent.mkdir(parents=True, exist_ok=True)
    buffer = io.BytesIO()
    prs.save(buffer)
    payload = _normalize_package(buffer.getvalue())
    # 先写临时文件再原子替换：中断不留「看起来完整、其实半截」的 PPTX
    staged = target.with_name(target.name + ".part")
    staged.write_bytes(payload)
    staged.replace(target)
    report["output_sha256"] = hashlib.sha256(payload).hexdigest()
    report["bytes"] = len(payload)
    return report


def _paint_background(slide, slide_spec: dict, colors: dict) -> None:
    token = slide_spec.get("background") or "background"
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _rgb(token, colors)


def _add_text(slide, element: dict, dpi: float, colors: dict, fonts: dict) -> None:
    box = slide.shapes.add_textbox(_emu(element["x"], dpi), _emu(element["y"], dpi),
                                   _emu(element["width"], dpi), _emu(element["height"], dpi))
    box.name = f"{element['id']}"
    tf = box.text_frame
    tf.word_wrap = bool(element.get("wrap", True))
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = _ANCHOR.get(str(element.get("anchor") or "top"), MSO_ANCHOR.TOP)
    lines = str(element.get("text") or "").split("\n")
    for i, line in enumerate(lines):
        paragraph = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        paragraph.alignment = _ALIGN.get(str(element.get("align") or "left"), PP_ALIGN.LEFT)
        if element.get("line_height"):
            paragraph.line_spacing = float(element["line_height"])
        run = paragraph.add_run()
        run.text = line
        run.font.size = _pt(element["size"], dpi)
        run.font.bold = bool(element.get("bold"))
        run.font.italic = bool(element.get("italic"))
        run.font.color.rgb = _rgb(element.get("color") or "ink", colors)
        _set_fonts(run, str(fonts.get("cn") or ""), str(fonts.get("latin") or ""))
        _set_tracking(run, element.get("tracking") or 0.0)


def _set_tracking(run, tracking_pt: float) -> None:
    """字距写 OOXML spc（1/100 pt）。中文大字收、小字放，都在策略里定。"""
    if not tracking_pt:
        return
    rpr = run.font._rPr
    rpr.set("spc", str(int(round(float(tracking_pt) * 100))))


def _add_shape(slide, element: dict, dpi: float, colors: dict) -> None:
    kind = element.get("shape")
    x, y = _emu(element["x"], dpi), _emu(element["y"], dpi)
    w, h = _emu(element["width"], dpi), _emu(element["height"], dpi)
    if kind in ("line", "arrow"):
        from pptx.enum.shapes import MSO_CONNECTOR
        connector = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, x, y, Emu(int(x) + int(w)), Emu(int(y) + int(h)))
        connector.name = element["id"]
        connector.line.color.rgb = _rgb(element.get("stroke") or "hairline", colors)
        connector.line.width = _pt(element.get("stroke_width") or 1.0, dpi)
        if kind == "arrow":
            _arrow_head(connector)
        return
    shape = slide.shapes.add_shape(_SHAPES.get(kind, MSO_SHAPE.RECTANGLE), x, y, w, h)
    shape.name = element["id"]
    fill = element.get("fill")
    if isinstance(fill, dict) and fill.get("type") == "none" or fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill if isinstance(fill, str) else fill.get("color"), colors)
    stroke = element.get("stroke")
    if stroke in (None, {"type": "none"}, "none"):
        shape.line.fill.background()
    else:
        shape.line.color.rgb = _rgb(stroke, colors)
        shape.line.width = _pt(element.get("stroke_width") or 1.0, dpi)
    if element.get("text"):
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        run = tf.paragraphs[0].add_run()
        run.text = str(element["text"])
        run.font.size = _pt(element.get("text_size") or 16, dpi)
        run.font.bold = bool(element.get("text_bold"))
        run.font.color.rgb = _rgb(element.get("text_color") or "ink", colors)
        if element.get("text_align"):
            tf.paragraphs[0].alignment = _ALIGN.get(element["text_align"], PP_ALIGN.LEFT)


def _arrow_head(connector) -> None:
    line = connector.line._get_or_add_ln()
    tail = line.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
    line.append(tail)


def _add_chart(slide, element: dict, dpi: float, colors: dict, fonts: dict) -> None:
    kind = element["chart_kind"]
    if kind == "progress":
        _add_progress(slide, element, dpi, colors)
        return
    data = CategoryChartData()
    data.categories = [str(c) for c in element["categories"]]
    for series in element["series"]:
        data.add_series(str(series["name"]), [float(v) for v in series["values"]])
    chart_type = XL_CHART_TYPE.DOUGHNUT if kind == "donut" else _CHART_TYPES[kind]
    frame = slide.shapes.add_chart(chart_type, _emu(element["x"], dpi), _emu(element["y"], dpi),
                                   _emu(element["width"], dpi), _emu(element["height"], dpi), data)
    frame.name = element["id"]
    chart = frame.chart
    chart.has_title = False
    chart.font.size = _pt(element.get("label_size") or 14, dpi)
    chart.font.name = str(fonts.get("latin") or "")
    chart.font.color.rgb = _rgb("muted", colors)
    multi = len(element["series"]) > 1
    chart.has_legend = multi or kind == "donut"
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.RIGHT if kind == "donut" else XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    palette = ["primary", "accent", "muted", "ink"]
    for i, series in enumerate(chart.plots[0].series):
        token = palette[i % len(palette)] if multi else "primary"
        if kind == "line":
            series.format.line.color.rgb = _rgb(token, colors)
            series.format.line.width = _pt(2.0, dpi)
        else:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = _rgb(token, colors)
    if kind == "donut":
        # 环图分片逐点上色：单序列的 donut 若只给 series 上色，PowerPoint 会用主题 accent
        # 系列自动铺色，成片就出现品牌色之外的颜色（案例 A 实测：出现 #1F4E5A 等默认色）。
        for i, point in enumerate(chart.plots[0].series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = _rgb(palette[i % len(palette)], colors)
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = bool(element.get("show_values"))
    if plot.has_data_labels:
        labels = plot.data_labels
        labels.font.size = _pt(element.get("value_size") or 14, dpi)
        labels.font.name = str(fonts.get("latin") or "")
        labels.font.color.rgb = _rgb("ink", colors)
        labels.number_format_is_linked = False
        labels.number_format = _number_format(element)
    if kind == "donut":
        _donut_hole(chart)
    _hide_axes(chart)


def _number_format(element: dict) -> str:
    unit = str(element.get("unit_suffix") or "")
    return f'#,##0.##"{unit}"' if unit else "#,##0.##"


def _donut_hole(chart) -> None:
    plot = chart.plots[0]
    hole = plot._element.find(qn("c:holeSize"))
    if hole is None:
        hole = plot._element.makeelement(qn("c:holeSize"), {"val": "58"})
        plot._element.append(hole)


def _hide_axes(chart) -> None:
    try:
        value_axis = chart.value_axis
        value_axis.has_major_gridlines = False
        value_axis.format.line.fill.background()
        value_axis.tick_labels.font.size = Pt(9)
        category_axis = chart.category_axis
        category_axis.format.line.color.rgb = RGBColor(0xD8, 0xD3, 0xCB)
        category_axis.tick_labels.font.size = Pt(9)
    except (AttributeError, ValueError):
        pass


def colors_of(slide) -> dict:
    """取当前主题色（编译期由 compile_deck 注入到 slide 上的私有属性）。"""
    return getattr(slide, "_pvad_colors", {}) or {}


def _set_alpha(shape, alpha: float) -> None:
    """给纯色填充加透明度（OOXML a:alpha，单位 1/1000 %）。"""
    solid = shape.fill._xPr.find(qn("a:solidFill"))
    if solid is None:
        return
    clr = solid.find(qn("a:srgbClr"))
    if clr is None:
        return
    node = clr.makeelement(qn("a:alpha"), {"val": str(int(round(max(0.0, min(1.0, alpha)) * 100000)))})
    clr.append(node)


def _add_image(slide, element: dict, dpi: float, base_dir) -> None:
    """写入原生图片。cover 时按已算好的裁切比例真正裁切，而不是拉伸变形。"""
    from assets import resolve_source
    path = resolve_source(element.get("source"), base_dir)
    if path is None:
        raise CompileError(f"素材不存在：{element.get('source')!r}（相对 {base_dir}）")
    picture = slide.shapes.add_picture(str(path), _emu(element["x"], dpi), _emu(element["y"], dpi),
                                       _emu(element["width"], dpi), _emu(element["height"], dpi))
    picture.name = element["id"]
    crop = element.get("crop") or [0.0, 0.0, 0.0, 0.0]
    if any(crop):
        picture.crop_left, picture.crop_top, picture.crop_right, picture.crop_bottom = (
            float(crop[0]), float(crop[1]), float(crop[2]), float(crop[3]))
    overlay = float(element.get("overlay") or 0.0)
    if overlay > 0:
        # 压字用的半透明层：真实存在一个矩形（可编辑、可删），不靠改原图
        veil = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, _emu(element["x"], dpi), _emu(element["y"], dpi),
                                      _emu(element["width"], dpi), _emu(element["height"], dpi))
        veil.name = f"{element['id']}.overlay"
        veil.line.fill.background()
        veil.fill.solid()
        veil.fill.fore_color.rgb = _rgb("ink", colors_of(slide))
        _set_alpha(veil, overlay)


def _add_progress(slide, element: dict, dpi: float, colors: dict) -> None:
    """进度条用形状画：值是量，不是装装饰。"""
    values = [float(v) for s in element["series"] for v in s["values"]]
    peak = max(values) if values else 1.0
    ratio = (values[0] / peak) if peak else 0.0
    x, y = element["x"], element["y"]
    w, h = element["width"], element["height"]
    track = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _emu(x, dpi), _emu(y, dpi),
                                   _emu(w, dpi), _emu(h, dpi))
    track.name = f"{element['id']}.track"
    track.fill.solid()
    track.fill.fore_color.rgb = _rgb("hairline", colors)
    track.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _emu(x, dpi), _emu(y, dpi),
                                 _emu(max(w * ratio, 2), dpi), _emu(h, dpi))
    bar.name = f"{element['id']}.fill"
    bar.fill.solid()
    bar.fill.fore_color.rgb = _rgb("primary", colors)
    bar.line.fill.background()


_EMBEDDED_OOXML = (".xlsx", ".xlsm", ".docx", ".pptx")
_CORE_TIME = re.compile(rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:(?:created|modified)>)")
_FIXED_ISO = b"1980-01-01T00:00:00Z"


def _normalize_package(payload: bytes) -> bytes:
    """重写 ZIP：固定时间戳与内嵌文档的时间戳，使同一 spec 的产物字节稳定。"""
    source = zipfile.ZipFile(io.BytesIO(payload))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for name in sorted(source.namelist()):
            data = source.read(name)
            if name.endswith(_EMBEDDED_OOXML) or name.startswith("docProps/"):
                data = _normalize_ooxml(data)
            info = zipfile.ZipInfo(name, date_time=_ZIP_EPOCH)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            target.writestr(info, data)
    return out.getvalue()


def _normalize_ooxml(blob: bytes) -> bytes:
    """内嵌图表工作簿/文档属性里的时间戳与 ZIP 时间同样归零。"""
    if not blob.startswith(b"PK"):
        return _CORE_TIME.sub(lambda m: m.group(1) + _FIXED_ISO + m.group(2), blob)
    try:
        source = zipfile.ZipFile(io.BytesIO(blob))
    except zipfile.BadZipFile:
        return blob
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for name in sorted(source.namelist()):
            data = source.read(name)
            if name.endswith((".xml", ".rels")):
                data = _CORE_TIME.sub(lambda m: m.group(1) + _FIXED_ISO + m.group(2), data)
            info = zipfile.ZipInfo(name, date_time=_ZIP_EPOCH)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            target.writestr(info, data)
    return out.getvalue()


def readback(path: str | Path) -> dict:
    """读回 PPTX 的可见事实（几何/文字/字号/颜色/图表数据/底色），供 fidelity 对账。"""
    prs = Presentation(str(path))
    dpi = 96.0
    slides = []
    for slide in prs.slides:
        shapes = []
        for shape in slide.shapes:
            item = {"name": shape.name,
                    "x": round(shape.left / 914400 * dpi, 2),
                    "y": round(shape.top / 914400 * dpi, 2),
                    "width": round(shape.width / 914400 * dpi, 2),
                    "height": round(shape.height / 914400 * dpi, 2)}
            if shape.has_text_frame:
                item["type"] = "text"
                item["text"] = shape.text_frame.text
                sizes, colors = [], []
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.size:
                            sizes.append(round(float(run.font.size.pt) / 72 * 96, 2))
                        try:
                            if run.font.color and run.font.color.rgb:
                                colors.append(str(run.font.color.rgb))
                        except (AttributeError, TypeError):
                            pass
                item["sizes"] = sizes
                item["colors"] = colors
                item["anchor"] = str(shape.text_frame.vertical_anchor)
            elif getattr(shape, "has_chart", False):
                chart = shape.chart
                item["type"] = "chart"
                item["chart_type"] = str(chart.chart_type)
                item["kind"] = _chart_kind_name(str(chart.chart_type))
                item["categories"] = [str(c) for c in chart.plots[0].categories]
                item["series"] = [{"name": s.name,
                                   "values": [float(v) if v is not None else None for v in s.values]}
                                  for s in chart.series]
            elif _is_picture(shape):
                item["type"] = "image"
                try:
                    blob = shape.image.blob
                    item["image_sha256"] = hashlib.sha256(blob).hexdigest()
                    item["image_blob"] = blob
                    item["image_px"] = [int(shape.image.size[0]), int(shape.image.size[1])]
                except (AttributeError, ValueError):
                    pass
                item["crop"] = [round(float(shape.crop_left), 4), round(float(shape.crop_top), 4),
                                round(float(shape.crop_right), 4), round(float(shape.crop_bottom), 4)]
                item["fit"] = "cover" if any(item["crop"]) else "contain"
            else:
                item["type"] = "shape"
                item["shape_kind"] = _shape_kind_name(shape)
                fill, stroke = _shape_paint(shape)
                item["fill"] = fill
                item["stroke"] = stroke
            shapes.append(item)
        slides.append({"id": slide.name if hasattr(slide, "name") else "",
                       "background": _background(prs, slide),
                       "shapes": shapes})
    return {"slides": slides,
            "slide_size": [round(prs.slide_width / 914400 * dpi, 2),
                           round(prs.slide_height / 914400 * dpi, 2)]}


def _background(prs, slide):
    try:
        fill = slide.background.fill
        if fill.type is not None and fill.fore_color and fill.fore_color.rgb is not None:
            return "#" + str(fill.fore_color.rgb)
    except (AttributeError, TypeError, ValueError):
        pass
    return None


def _is_picture(shape) -> bool:
    from pptx.shapes.picture import Picture
    return isinstance(shape, Picture)


def _shape_kind_name(shape) -> str:
    name = str(getattr(shape, "shape_type", "") or "")
    if "LINE" in name or "CONNECTOR" in name:
        return "line"
    return "rect"


def _shape_paint(shape):
    fill = stroke = None
    try:
        fill = "#" + str(shape.fill.fore_color.rgb) if shape.fill.type is not None else None
    except (AttributeError, TypeError, ValueError):
        fill = None
    try:
        stroke = "#" + str(shape.line.color.rgb)
    except (AttributeError, TypeError, ValueError):
        stroke = None
    return fill, stroke


def _chart_kind_name(chart_type: str) -> str:
    text = chart_type.upper()
    if "DOUGHNUT" in text or "DONUT" in text:
        return "donut"
    if "LINE" in text:
        return "line"
    if "BAR" in text and "COLUMN" not in text:
        return "bar"
    if "COLUMN" in text:
        return "column"
    return "unknown"
