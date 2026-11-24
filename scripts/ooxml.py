#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OOXML 低层写入：填充、渐变、描边、圆角、双色调、run 属性。compile 与 revise 共用。"""

from __future__ import annotations

from typing import Any, Dict, Optional

from lxml import etree

from model import Box, DASH, LIMITS, corner_adj, qn


# a:blip 里排在 duotone 之后的元素（ECMA-376 CT_Blip 的 EG_ColorTransform 次序）。
_BLIP_AFTER_DUOTONE = ("fillOverlay", "grayscale", "hsl", "lum", "tint")

# ---------------------------------------------------------------- 绘图底件


def _strip_style(shape) -> None:
    """删掉 p:style：主题的 effectRef 会带来投影，LibreOffice 与 PowerPoint 表现还不一致。
    形状的填充与描边全部由 plan 显式给出，不需要主题引用。"""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)

def _alpha_on(color_node, alpha: float) -> None:
    for old in color_node.findall(qn("a:alpha")):
        color_node.remove(old)
    etree.SubElement(color_node, qn("a:alpha")).set("val", str(int(round(alpha * 100000))))

def _apply_fill_alpha(shape, alpha: float) -> None:
    solid = shape.fill._xPr.find(qn("a:solidFill"))
    if solid is None:
        return
    _alpha_on(solid[0], alpha)

def _apply_gradient(shape, grad: Dict[str, Any]) -> None:
    """两停线性渐变：两端各自可声明透明度。用于压图蒙版、光带与大面积色场过渡。"""
    from pptx.dml.color import RGBColor

    shape.fill.gradient()
    grad_fill = shape.fill._xPr.find(qn("a:gradFill"))
    gsLst = grad_fill.find(qn("a:gsLst"))
    while len(gsLst) > 2:  # 默认模板偶尔给三停，删掉多余的那端
        gsLst.remove(gsLst[-1])
    stops = list(shape.fill.gradient_stops)
    stops[0].position = 0.0
    stops[0].color.rgb = RGBColor.from_string(str(grad["from"]).lstrip("#").upper())
    stops[1].position = 1.0
    stops[1].color.rgb = RGBColor.from_string(str(grad["to"]).lstrip("#").upper())
    for i, key in enumerate(("from_alpha", "to_alpha")):
        if grad.get(key) is not None:
            clr = gsLst[i].find(qn("a:srgbClr"))
            if clr is not None:
                _alpha_on(clr, float(grad[key]))
    for old in grad_fill.findall(qn("a:lin")):
        grad_fill.remove(old)
    lin = etree.Element(qn("a:lin"))
    lin.set("ang", str(int(round((float(grad.get("angle", 0)) % 360) * 60000))))
    lin.set("scaled", "0")
    grad_fill.insert(1, lin)

def _set_corner_radius(shape, radius_in: float, box: Box) -> None:
    """默认 16.7% 常常过圆或过方；圆角半径按英寸声明，按短边换算后写入 adj。"""
    adj = corner_adj(radius_in, box[2], box[3])
    prst = shape._element.spPr.find(qn("a:prstGeom"))
    if prst is None:
        return
    for old in prst.findall(qn("a:avLst")):
        prst.remove(old)
    av = etree.SubElement(prst, qn("a:avLst"))
    gd = etree.SubElement(av, qn("a:gd"))
    gd.set("name", "adj")
    gd.set("fmla", f"val {adj}")

def _apply_duotone(pic, duo: Dict[str, Any]) -> None:
    """把图压进整稿色域：暗部→一个色、亮部→另一个色。原生可编辑，随时可摘。

    a:blip 的子元素有固定次序（alphaModFix 在 duotone 之前）。插错位置时 PowerPoint
    多数情况仍能开，但严格校验器会报错——顺序由这里统一负责。
    """
    blip = pic._element.blipFill.find(qn("a:blip"))
    if blip is None:
        return
    for old in blip.findall(qn("a:duotone")):
        blip.remove(old)
    node = etree.Element(qn("a:duotone"))
    for key in ("dark", "light"):
        clr = etree.SubElement(node, qn("a:srgbClr"))
        clr.set("val", str(duo[key]).lstrip("#").upper())
    for i, child in enumerate(blip):
        local = etree.QName(child).localname
        if local in _BLIP_AFTER_DUOTONE:
            blip.insert(i, node)
            return
    blip.append(node)

def _apply_stroke(shape, stroke: Optional[Dict[str, Any]], arrow: str = "none") -> None:
    from pptx.util import Pt
    from pptx.dml.color import RGBColor

    if not stroke:
        shape.line.fill.background()
        return
    shape.line.color.rgb = RGBColor.from_string(str(stroke["color"]).lstrip("#").upper())
    shape.line.width = Pt(float(stroke.get("width_pt") or LIMITS["geometry"]["stroke_default_pt"]))
    ln = shape.line._get_or_add_ln()
    if stroke.get("cap"):
        ln.set("cap", str(stroke["cap"]))
    if stroke.get("dash") and stroke["dash"] != "solid":
        for old in ln.findall(qn("a:prstDash")):
            ln.remove(old)
        etree.SubElement(ln, qn("a:prstDash")).set("val", DASH[str(stroke["dash"])])
    ends = {"end": ("tail",), "both": ("head", "tail")}.get(arrow, ())
    for side in ends:
        node = etree.SubElement(ln, qn(f"a:{side}End"))
        node.set("type", "triangle")
        node.set("w", "med")
        node.set("len", "med")

def _ensure_family(font, family: str) -> None:
    """图表文字也要写东亚字体：只写 latin 时，中文会被渲染器换掉，整稿字族就断了。"""
    font.name = family
    rPr = getattr(font, "_rPr", None)
    if rPr is None:
        return
    for tag in ("a:ea", "a:cs"):
        node = rPr.find(qn(tag))
        if node is None:
            node = etree.SubElement(rPr, qn(tag))
        node.set("typeface", family)

def set_run_attrs(run, *, font: Optional[str] = None, font_ea: Optional[str] = None,
                  size: Optional[float] = None, color: Optional[str] = None,
                  bold: Optional[bool] = None, spacing: Optional[float] = None) -> None:
    """只写已知属性；None 表示"不知道"，留空让目标文件继续继承。

    改稿通路的全部意义是"只改被点名的东西"。把继承值写成显式值，会把 44pt 的标题
    压成 18pt、把主题粗体改成常规，而保留核对仍然报 READY——这是最坏的一种失败。
    """
    from pptx.util import Pt
    from pptx.dml.color import RGBColor

    font_obj = run.font
    if size is not None:
        font_obj.size = Pt(float(size))
    if bold is not None:
        font_obj.bold = bool(bold)
    if font is not None:
        font_obj.name = font
    # 东亚/复杂文字字族只在**明确给出**时写。原稿 run 上本来没有 a:ea 时凭空补一个，
    # 等于把继承改成了显式值：样式其实没变，保留核对却会误报，目标机行为也变了。
    if font_ea is not None:
        rPr = run._r.get_or_add_rPr()
        for tag in ("a:ea", "a:cs"):
            node = rPr.find(qn(tag))
            if node is None:
                node = etree.SubElement(rPr, qn(tag))
            node.set("typeface", font_ea)
    if color is not None:
        font_obj.color.rgb = RGBColor.from_string(color.lstrip("#").upper())
    if spacing is not None:
        rPr = run._r.get_or_add_rPr()
        if spacing:
            rPr.set("spc", str(int(round(spacing * 100))))
        else:
            rPr.attrib.pop("spc", None)

def _set_run_font(run, family: str, size: float, color: str, bold: bool, spacing: float = 0.0) -> None:
    """新建对象一律连东亚字族一起写：只写 latin 时中文会被渲染器换掉，整稿字族就断了。
    同时写 lang：缺省时中文在 PowerPoint 里被当成未标注语言，拼写检查满屏红线，断行规则也不稳。"""
    set_run_attrs(run, font=family, font_ea=family, size=size, color=color, bold=bold, spacing=spacing)
    has_cjk = any("\u3400" <= ch <= "\u9fff" for ch in (run.text or ""))
    run._r.get_or_add_rPr().set("lang", "zh-CN" if has_cjk else "en-US")


def stable_xlsx(blob: bytes) -> bytes:
    """内嵌工作簿由 xlsxwriter 写出，zip 条目时间与 docProps/core.xml 的 created/modified 都是“此刻”，
    导致同一 plan 每次编译得到不同字节。把它们钉成固定值：相同输入 → 相同 PPTX（可复现）。"""
    import io
    import re
    import zipfile
    fixed = (2000, 1, 1, 0, 0, 0)
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*", rb"\g<1>2000-01-01T00:00:00Z", data)
            info = zipfile.ZipInfo(item.filename, fixed)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = item.external_attr
            zout.writestr(info, data)
    return out.getvalue()


def stabilize_package(path) -> None:
    """把整个包里的 zip 条目时间钉成固定值，内嵌工作簿同样处理：相同输入 → 相同字节。
    只动 zip 元数据与工作簿时间戳，不改任何部件的 XML。"""
    import os
    import zipfile
    from pathlib import Path
    path = Path(path)
    tmp = path.with_name(path.name + ".stable")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/embeddings/") and item.filename.endswith(".xlsx"):
                data = stable_xlsx(data)
            info = zipfile.ZipInfo(item.filename, (2000, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = item.external_attr
            zout.writestr(info, data)
    os.replace(tmp, path)
