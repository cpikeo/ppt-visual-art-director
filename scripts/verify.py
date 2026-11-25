#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
核验交付物：plan ↔ PPTX ↔ 真实渲染。

用法：
    python scripts/verify.py plan.json out.pptx [--render-dir DIR] [--report report.json]

核验分三层，各自给出结论，不合并成一个分数：
  1. 对象核对   成品里每个对象的类型、几何、文字（逐段逐 run）、样式、图表载荷与着色、
                表格全部单元格、路径顶点、组合结构、图片字节是否与 plan 一致。
  2. 渲染核对   用 LibreOffice 把本次 PPTX 转 PDF，再由 PyMuPDF 取页尺寸、文字词框、
                实际字体与 PNG。缺工具时如实记 SKIP，并说明哪些判断因此无法做出。
  3. 需人判断   构图、主次、概念是否成立、目标 PowerPoint 是否正常——机器不给结论。

exit 0 = 没有阻断项；1 = 有；2 = 无法读取输入。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import model as M  # noqa: E402
import checks as K  # noqa: E402
import charts as H  # noqa: E402
import chart_audit as A  # noqa: E402

EMU = M.EMU
TEXT_TYPES = M.TEXT_TYPES
qn = M.qn


def _in(value: Optional[int]) -> Optional[float]:
    return None if value is None else round(value / EMU, 3)


def _norm_ws(text: str) -> str:
    return " ".join(text.replace("\r\n", "\n").split())


# ---------------------------------------------------------------- 对象核对


def _collect(shapes) -> Dict[str, Any]:
    """按对象名建立索引，递归进入组合；同时保留每个组合的直接子对象顺序。"""
    index: Dict[str, Any] = {}
    order: Dict[str, List[str]] = {}

    def walk(container) -> List[str]:
        names = []
        for shape in container:
            index[shape.name] = shape
            names.append(shape.name)
            if getattr(shape, "shape_type", None) == 6:  # MSO_SHAPE_TYPE.GROUP
                order[shape.name] = walk(shape.shapes)
        return names

    order["slide"] = walk(shapes)
    return {"index": index, "order": order}


def verify_objects(plan: Dict[str, Any], pptx: Path, asset_root: Path, f: M.Report) -> Dict[str, Any]:
    from pptx import Presentation

    prs = Presentation(str(pptx))
    pages = plan["pages"]
    facts: Dict[str, Any] = {"slides": len(prs.slides), "expected_slides": len(pages), "images": {}}
    if len(prs.slides) != len(pages):
        f.add("HARD", "V-PAGES", "整稿", f"成品 {len(prs.slides)} 页 ≠ plan {len(pages)} 页")

    for index, page in enumerate(pages):
        if index >= len(prs.slides):
            break
        pid = page["id"]
        slide = prs.slides[index]
        book = _collect(slide.shapes)
        boxes = M.page_geometry(page)

        expected_top = [f"{pid}.{el['id']}" for el in page["elements"]]
        actual_top = [n for n in book["order"].get("slide", []) if n != "_page_number"]
        if actual_top != expected_top:
            f.add("HARD", "V-ORDER", pid, f"叠放顺序 {actual_top} ≠ 声明顺序 {expected_top}")

        verify_background(page, slide, pid, f)

        for el in M.all_elements(page):
            name = f"{pid}.{el['id']}"
            shape = book["index"].get(name)
            if shape is None:
                f.add("HARD", "V-OBJ", name, "成品里找不到该对象")
                continue
            want = boxes.get(el["id"])
            if want is None:
                continue
            box = (_in(shape.left), _in(shape.top), _in(shape.width), _in(shape.height))
            if box[3] is None:
                f.add("HARD", "V-GEOM", name, "对象没有几何信息")
            elif any(abs(a - b) > 0.02 for a, b in zip(box, want)):
                f.add("HARD", "V-GEOM", name, f"几何 {box} ≠ plan {tuple(round(v, 3) for v in want)}")
            # 旋转对所有类型统一核对：按类型分头查，就会漏掉没实现的那几种。
            want_rot = float(el.get("rotation") or 0)
            got_rot = float(getattr(shape, "rotation", 0) or 0)
            if _angle_diff(got_rot, want_rot) > 0.2:
                f.add("HARD", "V-STYLE", name, f"旋转 {got_rot}° ≠ plan {want_rot}°")
            kind = el["type"]
            alt = M.alt_text(el)
            if alt:
                got_alt = (shape._element.xpath(".//p:cNvPr")[0].get("descr") or "")
                if got_alt != alt:
                    f.add("HARD", "V-ALT", name, f"替代文字 {got_alt!r} ≠ plan {alt!r}")
            if kind in TEXT_TYPES:
                verify_text(plan, page, el, shape, name, f)
            elif kind in ("shape", "divider"):
                verify_shape(page, el, shape, name, f)
            elif kind == "path":
                verify_path(page, el, shape, name, f)
            elif kind == "image":
                verify_image(plan, el, shape, name, facts, asset_root, f)
            elif kind == "chart":
                verify_chart(plan, page, el, shape, name, f)
            elif kind == "legend":
                verify_legend(plan, page, el, shape, name, f)
            elif kind == "group":
                kids = book["order"].get(name, [])
                want_kids = [f"{pid}.{c['id']}" for c in el["elements"]]
                if kids != want_kids:
                    f.add("HARD", "V-OBJ", name, f"组合成员 {kids} ≠ plan {want_kids}")

        want_notes = page.get("speaker_notes") or ""
        got_notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
        if got_notes != want_notes:
            f.add("HARD", "V-NOTES", pid, f"备注 {got_notes[:30]!r} ≠ plan {want_notes[:30]!r}")
        verify_page_number(plan, page, book, index, pid, f)
    return facts


def verify_background(page: Dict[str, Any], slide, pid: str, f: M.Report) -> None:
    bg = slide._element.find(qn("p:cSld")).find(qn("p:bg"))
    value = None
    if bg is not None:
        bgPr = bg.find(qn("p:bgPr"))
        if bgPr is not None:
            fill = bgPr.find(qn("a:solidFill"))
            if fill is not None:
                clr = fill.find(qn("a:srgbClr"))
                if clr is not None:
                    value = clr.get("val")
    want = M.page_bg(page).lstrip("#").upper()
    if value is None:
        f.add("HARD", "V-STYLE", pid, f"页面底色不是纯色填充（期望 #{want}）")
    elif value.upper() != want:
        f.add("HARD", "V-STYLE", pid, f"页面底色 #{value} ≠ plan #{want}")


def _expected_runs(plan, page, el) -> List[Dict[str, Any]]:
    st = M.text_style(plan, page, el)
    out = []
    for seg, pieces in (M.paragraph_runs(el) or [(str(el.get("content", "")), [{"text": str(el.get("content", ""))}])]):
        for piece in pieces:
            out.append({
                "text": piece.get("text", ""),
                "size": float(piece.get("font_size") or st["size"]),
                "color": str(piece.get("color") or st["color"]).upper(),
                "bold": bool(piece.get("bold", st["bold"])),
                "family": piece.get("font") or st["family"],
                "spacing": float(piece.get("letter_spacing_pt", st["spacing"])),
            })
    return out


def _run_family(run) -> Tuple[Optional[str], Optional[str]]:
    rPr = run._r.find(qn("a:rPr"))
    ea = None
    if rPr is not None:
        node = rPr.find(qn("a:ea"))
        if node is not None:
            ea = node.get("typeface")
    return run.font.name, ea


def _angle_diff(a: float, b: float) -> float:
    d = abs((a - b) % 360)
    return min(d, 360 - d)


def verify_text(plan, page, el, shape, name, f: M.Report) -> None:
    from pptx.enum.dml import MSO_COLOR_TYPE

    if not getattr(shape, "has_text_frame", False):
        f.add("HARD", "V-OBJ", name, "该对象没有文本框（plan 声明为文字）")
        return
    actual = shape.text_frame.text
    if actual != el["content"]:
        if _norm_ws(actual) == _norm_ws(el["content"]):
            f.add("SOFT", "V-TEXT", name, f"只差空白与换行：成品「{actual!r}」")
        else:
            f.add("HARD", "V-TEXT", name, f"文字不一致：成品「{actual[:60]!r}」≠ plan「{el['content'][:60]!r}」")

    st = M.text_style(plan, page, el)
    paragraphs = shape.text_frame.paragraphs
    want_paras = M.paragraph_runs(el) or [(el["content"], [])]
    if len(paragraphs) != len(want_paras):
        f.add("HARD", "V-TEXT", name, f"段落数 {len(paragraphs)} ≠ plan {len(want_paras)}")
    for i, para in enumerate(paragraphs[: len(want_paras)]):
        got_align = str(para.alignment).split(" ")[0] if para.alignment is not None else "NONE"
        if got_align.upper() != st["align"].upper():
            f.add("HARD", "V-STYLE", f"{name}#p{i + 1}", f"对齐 {got_align} ≠ {st['align']}")
        if para.line_spacing is not None and abs(float(para.line_spacing) - st["line_spacing"]) > 0.02:
            f.add("SOFT", "V-STYLE", f"{name}#p{i + 1}",
                  f"行距 {float(para.line_spacing):.2f} ≠ plan {st['line_spacing']:.2f}")

    pad_x, pad_y = st["pad"]
    for side, want_pad in (("left", pad_x), ("right", pad_x), ("top", pad_y), ("bottom", pad_y)):
        got_pad = getattr(shape.text_frame, f"margin_{side}")
        if got_pad is None or abs(got_pad / EMU - want_pad) > 0.005:
            f.add("SOFT", "V-STYLE", name, f"{side} 内边距 {None if got_pad is None else round(got_pad / EMU, 3)}in ≠ {want_pad}in")

    anchor = shape.text_frame.vertical_anchor
    want_anchor = {"top": "TOP", "middle": "MIDDLE", "bottom": "BOTTOM"}[st["valign"]]
    if anchor is not None and want_anchor not in str(anchor):
        f.add("SOFT", "V-STYLE", name, f"垂直对齐 {anchor} ≠ plan {st['valign']}")

    expected = _expected_runs(plan, page, el)
    runs = [r for p in shape.text_frame.paragraphs for r in p.runs]
    if len(runs) != len(expected):
        f.add("HARD", "V-RUNS", name, f"run 数量 {len(runs)} ≠ plan {len(expected)}（混排会被拆散或合并）")
        return
    for i, (run, want) in enumerate(zip(runs, expected)):
        tag = f"{name}#run{i + 1}"
        if run.text != want["text"]:
            f.add("HARD", "V-RUNS", tag, f"片段文字「{run.text!r}」≠「{want['text']!r}」")
        size = run.font.size.pt if run.font.size else None
        if size is None or abs(size - want["size"]) > 0.5:
            f.add("HARD", "V-STYLE", tag, f"字号 {size} ≠ {want['size']}")
        if bool(run.font.bold) != want["bold"]:
            f.add("HARD", "V-STYLE", tag, f"字重 bold={bool(run.font.bold)} ≠ {want['bold']}")
        color = run.font.color
        if color is None or color.type != MSO_COLOR_TYPE.RGB:
            f.add("HARD", "V-STYLE", tag, "墨色不是显式 RGB，目标机可能随主题变化")
        elif f"#{color.rgb}".upper() != "#" + want["color"].lstrip("#").upper():
            f.add("HARD", "V-STYLE", tag, f"墨色 #{color.rgb} ≠ 预期 {want['color']}")
        latin, ea = _run_family(run)
        if latin != want["family"]:
            f.add("HARD", "V-FONT", tag, f"西文字体「{latin}」≠「{want['family']}」")
        if ea != want["family"]:
            f.add("HARD", "V-FONT", tag, f"东亚字体「{ea}」≠「{want['family']}」（中文会回退到别的字族）")
        rPr = run._r.find(qn("a:rPr"))
        got_spc = float(rPr.get("spc") or 0) / 100.0 if rPr is not None else 0.0
        if abs(got_spc - want["spacing"]) > 0.05:
            f.add("HARD", "V-STYLE", tag, f"字距 {got_spc}pt ≠ {want['spacing']}pt")


def verify_shape(page, el, shape, name, f: M.Report) -> None:
    from pptx.enum.dml import MSO_COLOR_TYPE

    prst = shape._element.find(".//" + qn("a:prstGeom"))
    if el["type"] == "shape" and prst is not None:
        want = {"rect": "rect", "round_rect": "roundRect", "oval": "ellipse", "triangle": "triangle",
                "diamond": "diamond", "chevron": "chevron", "arrow_right": "rightArrow",
                "arrow_up": "upArrow"}[el.get("shape_kind") or "rect"]
        if prst.get("prst") != want:
            f.add("HARD", "V-OBJ", name, f"形状 {prst.get('prst')} ≠ {want}")
    if el.get("gradient"):
        verify_gradient(el["gradient"], shape, name, f)
    elif el.get("color"):
        color = shape.fill.fore_color
        if color.type != MSO_COLOR_TYPE.RGB:
            f.add("HARD", "V-STYLE", name, "填充不是显式 RGB")
        elif f"#{color.rgb}".upper() != "#" + str(el["color"]).lstrip("#").upper():
            f.add("HARD", "V-STYLE", name, f"填充 #{color.rgb} ≠ {el['color']}")
        alpha = _fill_alpha(shape)
        want_alpha = float(el.get("fill_alpha", 1.0))
        if abs(alpha - want_alpha) > 0.02:
            f.add("HARD", "V-STYLE", name, f"填充透明度 {alpha:.2f} ≠ {want_alpha:.2f}")
    if el.get("stroke"):
        ln = shape.line
        try:
            got = f"#{ln.color.rgb}".upper()
        except (AttributeError, TypeError):
            got = None
        if got != "#" + str(el["stroke"]["color"]).lstrip("#").upper():
            f.add("HARD", "V-STYLE", name, f"描边色 {got} ≠ {el['stroke']['color']}")
        width = ln.width.pt if ln.width else 0
        want_w = float(el["stroke"].get("width_pt") or M.LIMITS["geometry"]["stroke_default_pt"])
        if abs(width - want_w) > 0.1:
            f.add("HARD", "V-STYLE", name, f"描边 {width}pt ≠ {want_w}pt")
        if el["stroke"].get("dash") and el["stroke"]["dash"] != "solid":
            node = ln._get_or_add_ln().find(qn("a:prstDash"))
            if node is None:
                f.add("HARD", "V-STYLE", name, f"声明了虚线 {el['stroke']['dash']}，成品没有")
    if el.get("corner_in") is not None:
        want = M.corner_adj(float(el["corner_in"]), float(el["w"]), float(el["h"]))
        got = _got_adj(shape)
        if got is None or abs(got - want) > 400:
            f.add("HARD", "V-STYLE", name, f"圆角 adj {got} ≠ 期望 {want}（半径 {el['corner_in']}in）")


def _got_adj(shape) -> Optional[int]:
    gd = shape._element.spPr.find(".//" + qn("a:avLst") + "/" + qn("a:gd"))
    if gd is None or gd.get("name") != "adj":
        return None
    fmla = str(gd.get("fmla") or "")
    try:
        return int(fmla.split()[-1])
    except (IndexError, ValueError):
        return None


def verify_gradient(grad: Dict[str, Any], shape, name: str, f: M.Report) -> None:

    node = shape._element.find(".//" + qn("a:gradFill"))
    if node is None:
        f.add("HARD", "V-STYLE", name, "声明了渐变色，成品不是渐变填充")
        return
    stops = node.findall(".//" + qn("a:gsLst") + "/" + qn("a:gs"))
    if len(stops) != 2:
        f.add("HARD", "V-STYLE", name, f"渐变停靠点 {len(stops)} 个 ≠ 2")
        return
    got_colors, got_alphas = [], []
    for gs in stops:
        srgb = gs.find(".//" + qn("a:srgbClr"))
        got_colors.append("#" + srgb.get("val").upper() if srgb is not None else None)
        alpha = gs.find(".//" + qn("a:alpha"))
        got_alphas.append(int(alpha.get("val")) / 100000.0 if alpha is not None else 1.0)
    want_colors = ["#" + str(grad["from"]).lstrip("#").upper(), "#" + str(grad["to"]).lstrip("#").upper()]
    if got_colors != want_colors:
        f.add("HARD", "V-STYLE", name, f"渐变两端 {got_colors} ≠ {want_colors}")
    want_alphas = [float(grad.get("from_alpha", 1.0)), float(grad.get("to_alpha", 1.0))]
    if any(abs(a - b) > 0.02 for a, b in zip(got_alphas, want_alphas)):
        f.add("HARD", "V-STYLE", name, f"渐变透明度 {[round(a, 2) for a in got_alphas]} ≠ {want_alphas}")
    lin = node.find(qn("a:lin"))
    got_ang = float(lin.get("ang")) / 60000.0 % 360 if lin is not None and lin.get("ang") else 0.0
    if _angle_diff(got_ang, float(grad.get("angle", 0)) % 360) > 0.5:
        f.add("HARD", "V-STYLE", name, f"渐变角度 {got_ang:.0f}° ≠ {float(grad.get('angle', 0)) % 360:.0f}°")


def _fill_alpha(shape) -> float:
    solid = shape.fill._xPr.find(qn("a:solidFill"))
    if solid is None or not len(solid):
        return 1.0
    alpha = solid[0].find(qn("a:alpha"))
    return 1.0 if alpha is None else int(alpha.get("val")) / 100000.0


def _path_points(shape) -> List[Tuple[float, float]]:
    """读回 custGeom 的顶点，换算成页面绝对英寸（路径坐标系按 a:path 的 w/h 归一）。"""
    path = shape._element.find(".//" + qn("a:pathLst") + "/" + qn("a:path"))
    if path is None:
        return []
    w = float(path.get("w") or 0) or float(shape.width or 0)
    h = float(path.get("h") or 0) or float(shape.height or 0)
    sx = (shape.width / w) if w else 1.0
    sy = (shape.height / h) if h else 1.0
    out: List[Tuple[float, float]] = []
    for node in path:
        tag = etree.QName(node).localname
        if tag not in ("moveTo", "lnTo"):
            continue
        pt = node.find(qn("a:pt"))
        if pt is None:
            continue
        out.append(((shape.left + float(pt.get("x")) * sx) / EMU,
                    (shape.top + float(pt.get("y")) * sy) / EMU))
    return out


def verify_path(page, el, shape, name, f: M.Report) -> None:
    pts = M.resolve_points(el, M.page_geometry(page))
    want_kind = "connector" if len(pts) == 2 else "freeform"
    has_cust = shape._element.find(".//" + qn("a:custGeom")) is not None
    is_conn = shape._element.tag.endswith("}cxnSp")
    if want_kind == "connector" and not is_conn:
        f.add("HARD", "V-OBJ", name, "两点路径应为直线连接符，成品不是")
    if want_kind == "freeform" and not has_cust:
        f.add("HARD", "V-OBJ", name, "多点路径应为自由折线，成品没有自定义几何")

    # 顶点位置也要读回来对：只数个数，等于只要"形状还在"就算过。
    if is_conn:
        got = [(shape.begin_x / EMU, shape.begin_y / EMU), (shape.end_x / EMU, shape.end_y / EMU)]
    else:
        got = _path_points(shape)
    if len(got) != len(pts):
        if got or pts:
            f.add("HARD", "V-GEOM", name, f"路径顶点 {len(got)} ≠ plan {len(pts)}")
    else:
        # 连接符的起止点可能因翻转而对调，按无序比对。
        pairs = list(zip(pts, got)) if not is_conn else _best_pairing(pts, got)
        for (want, have) in pairs:
            if abs(want[0] - have[0]) > 0.02 or abs(want[1] - have[1]) > 0.02:
                f.add("HARD", "V-GEOM", name,
                      f"路径顶点 ({have[0]:.2f},{have[1]:.2f}) ≠ plan ({want[0]:.2f},{want[1]:.2f})")
                break
    if has_cust:
        path = shape._element.find(".//" + qn("a:pathLst") + "/" + qn("a:path"))
        closed = path is not None and path.find(qn("a:close")) is not None
        if closed != bool(el.get("closed")):
            f.add("HARD", "V-GEOM", name, f"闭合={closed} ≠ plan {bool(el.get('closed'))}：填充面会缺一条边")

    try:
        got_line = f"#{shape.line.color.rgb}".upper()
    except (AttributeError, TypeError):
        got_line = None
    if got_line != "#" + str(el["stroke"]["color"]).lstrip("#").upper():
        f.add("HARD", "V-STYLE", name, f"线色 {got_line} ≠ {el['stroke']['color']}")
    ln = shape.line._get_or_add_ln()
    want_arrow = str(el.get("arrow") or "none")
    has_tail = ln.find(qn("a:tailEnd")) is not None and ln.find(qn("a:tailEnd")).get("type") not in (None, "none")
    has_head = ln.find(qn("a:headEnd")) is not None and ln.find(qn("a:headEnd")).get("type") not in (None, "none")
    if want_arrow == "none" and (has_tail or has_head):
        f.add("HARD", "V-STYLE", name, "未声明箭头，成品却有端点箭头")
    if want_arrow in ("end", "both") and not has_tail:
        f.add("HARD", "V-STYLE", name, "声明了尾部箭头，成品没有")
    if want_arrow == "both" and not has_head:
        f.add("HARD", "V-STYLE", name, "声明了两端箭头，成品只有一端")


def _best_pairing(want: List[Tuple[float, float]], have: List[Tuple[float, float]]):
    """连接符的首尾可能被翻转写反：取两种配对里误差更小的那种。"""
    def cost(pairs) -> float:
        return sum(abs(a[0] - b[0]) + abs(a[1] - b[1]) for a, b in pairs)

    straight = list(zip(want, have))
    swapped = [(want[0], have[1]), (want[1], have[0])]
    return straight if cost(straight) <= cost(swapped) else swapped


def verify_image(plan, el, shape, name, facts, asset_root: Path, f: M.Report) -> None:
    from PIL import Image
    import io

    blob = getattr(shape, "image", None)
    if blob is None:
        f.add("HARD", "V-OBJ", name, "该对象不是图片")
        return
    actual_hash = hashlib.sha256(shape.image.blob).hexdigest()
    facts["images"][name] = actual_hash
    path = M.resolve_asset(asset_root, el["image_id"])
    if path is None:
        f.add("SOFT", "V-ASSET", name, "找不到素材原件，无法比对内嵌字节")
        return
    file_hash = hashlib.sha256(path.read_bytes()).hexdigest()
    if file_hash != actual_hash:
        f.add("SOFT", "V-ASSET", name, "内嵌字节与素材原件不同（可能被重新编码，或有损压缩）")
    try:
        with Image.open(io.BytesIO(shape.image.blob)) as embedded, Image.open(path) as original:
            if embedded.size != original.size:
                f.add("HARD", "V-ASSET", name, f"内嵌像素 {embedded.size} ≠ 原件 {original.size}")
    except Exception as exc:  # 解码失败就是失败：静默放行等于没核对
        f.add("HARD", "V-ASSET", name, f"无法解码内嵌图片或素材原件，像素核对无法完成：{type(exc).__name__}")
    want_crop = M._crop_geometry(shape.image.size[0], shape.image.size[1],
                                 float(el["w"]), float(el["h"]), el.get("crop_focus"))
    got_crop = (shape.crop_left, shape.crop_right, shape.crop_top, shape.crop_bottom)
    if any(abs(a - b) > 0.01 for a, b in zip(got_crop, want_crop)):
        f.add("HARD", "V-STYLE", name, f"裁切 {tuple(round(v, 3) for v in got_crop)} ≠ {tuple(round(v, 3) for v in want_crop)}")

    blip = shape._element.blipFill.find(qn("a:blip"))
    node = blip.find(qn("a:alphaModFix")) if blip is not None else None
    got_alpha = int(node.get("amt")) / 100000.0 if node is not None else 1.0
    want_alpha = float(el.get("alpha", 1.0))
    if abs(got_alpha - want_alpha) > 0.02:
        f.add("HARD", "V-STYLE", name, f"图片透明度 {got_alpha:.2f} ≠ {want_alpha:.2f}")

    duo = el.get("duotone")
    got_duo = blip.find(qn("a:duotone")) if blip is not None else None
    if duo:
        if got_duo is None:
            f.add("HARD", "V-STYLE", name, "声明了双色调，成品的图片没有该效果")
        else:
            colors = ["#" + (c.get("val") or "").upper() for c in got_duo.findall(qn("a:srgbClr"))]
            want = ["#" + str(duo["dark"]).lstrip("#").upper(), "#" + str(duo["light"]).lstrip("#").upper()]
            if colors != want:
                f.add("HARD", "V-STYLE", name, f"双色调 {colors} ≠ {want}（暗部在前、亮部在后）")
    elif got_duo is not None:
        f.add("SOFT", "V-STYLE", name, "未声明双色调，成品图片却有该效果")


def _series_from_xml(shape) -> List[Dict[str, Any]]:
    """跨 plot 读取全部系列：组合图里柱与线分属两个 plot，按名称核对才可靠。"""
    out = []
    for ser in shape.chart._chartSpace.iter(qn("c:ser")):
        name_el = ser.find(qn("c:tx")).find(".//" + qn("c:v")) if ser.find(qn("c:tx")) is not None else None
        name = name_el.text if name_el is not None else ""
        values = []
        for holder in (qn("c:val"), qn("c:yVal")):
            node = ser.find(holder)
            if node is None:
                continue
            cache = node.find(".//" + qn("c:numCache"))
            if cache is None:
                continue
            values = [float(p.findtext(qn("c:v"))) for p in cache.findall(qn("c:pt"))]
            break
        xs = []
        xnode = ser.find(qn("c:xVal"))
        if xnode is not None:
            xcache = xnode.find(".//" + qn("c:numCache"))
            if xcache is not None:
                xs = [float(p.findtext(qn("c:v"))) for p in xcache.findall(qn("c:pt"))]
        out.append({"name": name, "values": values, "x": xs})
    return out


def verify_chart(plan, page, el, shape, name, f: M.Report) -> None:
    spec = el["chart"]
    kind = spec["kind"]
    if kind == "table":
        if not shape.has_table:
            f.add("HARD", "V-CHART", name, "该对象不是原生表格")
            return
        table = shape.table
        headers = [str(v) for v in spec["data"]["headers"]]
        rows = spec["data"]["rows"]
        if len(table.rows) != len(rows) + 1 or len(table.columns) != len(headers):
            f.add("HARD", "V-CHART", name,
                  f"表格 {len(table.rows)}×{len(table.columns)} ≠ plan {len(rows) + 1}×{len(headers)}")
        for c in range(min(len(headers), len(table.columns))):
            got = table.cell(0, c).text
            if got != headers[c]:
                f.add("HARD", "V-CHART", name, f"表头第 {c + 1} 列「{got}」≠「{headers[c]}」")
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                if r + 1 >= len(table.rows) or c >= len(table.columns):
                    continue
                got = table.cell(r + 1, c).text
                if got != str(value):
                    f.add("HARD", "V-CHART", name, f"单元格({r + 1},{c + 1})「{got}」≠「{value}」")
        widths = spec.get("column_weights") or [1] * len(headers)
        total = sum(widths)
        for i, weight in enumerate(widths):
            if i >= len(table.columns):
                break
            want = float(el["w"]) * weight / total
            got = table.columns[i].width / EMU
            if abs(got - want) > 0.02:
                f.add("SOFT", "V-CHART", name, f"第 {i + 1} 列宽 {got:.2f}in ≠ {want:.2f}in")
        return

    if not shape.has_chart:
        f.add("HARD", "V-CHART", name, "该对象不是原生图表")
        return
    obj = shape.chart
    if obj.chart_type != H._chart_type(spec):
        f.add("HARD", "V-CHART", name, f"图表类型 {obj.chart_type} ≠ {H._chart_type(spec)}")
    combo = spec.get("combo_line_series") or []
    plots = len(obj.plots)
    want_plots = 2 if combo else 1
    if plots != want_plots:
        f.add("HARD", "V-CHART", name, f"plot 数量 {plots} ≠ {want_plots}（组合图应有柱与线两层）")
    if combo and obj._chartSpace.find(".//" + qn("c:lineChart")) is None:
        f.add("HARD", "V-CHART", name, "声明了 combo_line_series，成品没有折线层")

    actual = _series_from_xml(shape)
    by_name = {s["name"]: s for s in actual}
    for spec_series in spec["data"]["series"]:
        got = by_name.get(str(spec_series["name"]))
        if got is None:
            f.add("HARD", "V-CHART", name, f"缺少系列「{spec_series['name']}」")
            continue
        if kind == "scatter":
            want = [(round(float(x), 6), round(float(y), 6)) for x, y in spec_series["points"]]
            have = [(round(x, 6), round(y, 6)) for x, y in zip(got["x"], got["values"])]
            if have != want:
                f.add("HARD", "V-CHART", name, f"系列「{spec_series['name']}」坐标点数或数值 {have} ≠ plan {want}")
        else:
            want = [round(float(v), 6) for v in spec_series["values"]]
            have = [round(v, 6) for v in got["values"]]
            if have != want:
                f.add("HARD", "V-CHART", name, f"系列「{spec_series['name']}」数值 {have} ≠ plan {want}")

    if kind != "scatter":
        cats = [str(c) for c in obj.plots[0].categories]
        want_cats = [str(c) for c in spec["data"]["categories"]]
        if [_norm_ws(c) for c in cats] != [_norm_ws(c) for c in want_cats]:
            f.add("HARD", "V-CHART", name, f"类别 {cats} ≠ plan {want_cats}")

    title = spec.get("title")
    has_title = obj.has_title and obj.chart_title.text_frame.text.strip() != ""
    if spec.get("show_title", True) and title:
        if not has_title:
            f.add("HARD", "V-CHART", name, "缺少图题")
        elif obj.chart_title.text_frame.text != title:
            f.add("HARD", "V-CHART", name, f"图题「{obj.chart_title.text_frame.text}」≠「{title}」")
    elif has_title:
        f.add("SOFT", "V-CHART", name, f"未声明图题，成品却有「{obj.chart_title.text_frame.text}」")

    labels = M.chart_flags(spec)["data_labels"] or bool(spec.get("combo_line_labels"))
    has_labels = _labels_on(obj._chartSpace)
    if labels and not has_labels:
        f.add("HARD", "V-CHART", name, "声明了直标，成品没有数据标签")
    if not labels and has_labels and kind not in ("pie", "donut", "scatter"):
        f.add("SOFT", "V-CHART", name, "未声明直标，成品却有数据标签")

    # 有效色审计：只读成品 XML，不读 plan 的任何颜色（见 chart_audit.py）。
    surface = M.chart_surface(page, el)
    for sev, code, message in A.audit_chart(obj._chartSpace, surface, kind):
        f.add(sev, code, name, message)
    legend_els = [e for e in M.all_elements(page) if e.get("type") == "legend" and str(e.get("for_chart")) == str(el.get("id"))]
    for sev, code, message in A.audit_readout(obj._chartSpace, kind, bool(legend_els),
                                              any(e.get("show_value") for e in legend_els)):
        f.add(sev, code, name, message)

    if spec.get("highlight_category") is not None and not list(obj._chartSpace.iter(qn("c:dPt"))):
        f.add("HARD", "V-CHART", name, "声明了 highlight_category，但成品没有逐点着色")

    verify_chart_style(plan, page, el, shape, name, f)


def _labels_on(chart_space) -> bool:
    """默认生成的 c:dLbls 里 showVal=0，等于没开直标；只看元素在不在会误判。"""
    for lbl in chart_space.iter(qn("c:dLbls")):
        show = lbl.find(qn("c:showVal"))
        if show is not None and show.get("val") not in (None, "0", "false"):
            return True
    return False


def _ser_name(ser) -> str:
    tx = ser.find(qn("c:tx"))
    if tx is None:
        return ""
    node = tx.find(".//" + qn("c:v"))
    return node.text if node is not None and node.text else ""


def verify_legend(plan, page, el, shape, name: str, f: M.Report) -> None:
    """自定义图例逐块核对：色块必须等于颜色映射，标签必须等于类别/系列名。

    这是"环形图与图例颜色不匹配"的收口检查：手绘色块一旦与扇区取色脱节，这里直接阻断。
    """
    if getattr(shape, "shapes", None) is None:
        f.add("HARD", "V-OBJ", name, "图例不是组合对象")
        return
    ref = str(el.get("for_chart") or "")
    target = next((e for e in (page.get("elements") or []) if str(e.get("id")) == ref), None)
    if target is None or target.get("type") != "chart":
        f.add("HARD", "V-OBJ", name, f"for_chart「{ref}」在本页找不到 chart 元素")
        return
    spec = target["chart"]
    items = M.chart_items(spec)
    colors = M.chart_color_map(plan, spec)
    values = list(((spec.get("data") or {}).get("series") or [{}])[0].get("values") or [])
    show_value = bool(el.get("show_value")) and spec.get("kind") in ("pie", "donut")
    kids = {c.name: c for c in shape.shapes}
    for i, (item, color) in enumerate(zip(items, colors)):
        sw = kids.get(f"{name}#sw{i}")
        if sw is None:
            f.add("HARD", "V-OBJ", name, f"缺少第 {i + 1} 个色块「{item}」")
            continue
        try:
            got = f"#{sw.fill.fore_color.rgb}".upper()
        except (AttributeError, TypeError, ValueError):
            got = None
        if got != color.upper():
            f.add("HARD", "V-CHART", name, f"色块「{item}」{got} ≠ 颜色映射 {color.upper()}：图例与图形错色")
        tx = kids.get(f"{name}#tx{i}")
        if tx is None:
            f.add("HARD", "V-OBJ", name, f"缺少第 {i + 1} 个标签「{item}」")
            continue
        runs = [r for para in tx.text_frame.paragraphs for r in para.runs]
        try:
            ink = f"#{runs[0].font.color.rgb}".upper()
        except (IndexError, AttributeError, TypeError, ValueError):
            ink = None
        surface = M.chart_surface(page, el)
        if ink is not None and M.contrast_ratio(ink, surface) < M.LABEL_INK_MIN:
            f.add("HARD", "V-LABEL-INK", name,
                  f"图例标签「{item}」墨色 {ink} 对承载面 {surface} 仅 {M.contrast_ratio(ink, surface):.2f}:1")
        want = H._legend_label(spec, str(item), values, i, show_value)
        got_text = tx.text_frame.text
        if _norm_key(got_text) != _norm_key(want):
            f.add("HARD", "V-TEXT", name, f"图例标签「{got_text}」≠「{want}」")
    expected = {f"{name}#sw{i}" for i in range(len(items))} | {f"{name}#tx{i}" for i in range(len(items))}
    extra = sorted(n for n in kids if n not in expected)
    if extra:
        f.add("SOFT", "V-OBJ", name, f"图例里有计划外的成员 {extra}：多出来的色块没有数据对应")


def verify_chart_style(plan, page, el, shape, name, f: M.Report) -> None:
    """着色、柱宽、轴开关与轴名：这些决定读者能不能读出数，且全部来自 plan。"""

    spec = el["chart"]
    kind = spec["kind"]
    obj = shape.chart
    # 与 compile 同一份映射：颜色对不上就是 HARD，不给"两边一起错"留通道。
    palette = M.chart_color_map(plan, spec)
    order = {name: i for i, name in enumerate(M.chart_items(spec))}

    if kind in ("pie", "donut"):
        cats = spec["data"]["categories"]
        points = list(obj._chartSpace.iter(qn("c:dPt")))
        if len(points) < len(cats):
            f.add("HARD", "V-CHART", name, f"逐点着色 {len(points)} 个 ≠ 类别 {len(cats)} 个")
        for j, dpt in enumerate(points[: len(cats)]):
            srgb = dpt.find(".//" + qn("a:srgbClr"))
            want = palette[j] if j < len(palette) else None
            if want is None or srgb is None or f"#{srgb.get('val').upper()}" != want:
                f.add("HARD", "V-CHART", name, f"第 {j + 1} 个扇区颜色 ≠ 调色板 {want}")
    elif kind == "scatter":
        # 散点的颜色在 marker 上，不在系列线上；读成品 marker，与调色板逐系列对账。
        for ser in obj._chartSpace.iter(qn("c:ser")):
            sname = _ser_name(ser)
            if sname not in order:
                continue
            want = palette[order[sname]] if order[sname] < len(palette) else None
            mk = A.marker_of(ser)
            if want is None or mk["fill"] != want:
                f.add("HARD", "V-CHART", name, f"散点系列「{sname}」marker 填充 {mk['fill']} ≠ 调色板 {want}")
            if mk["line"] is not None and mk["line"] != want:
                f.add("HARD", "V-CHART", name, f"散点系列「{sname}」marker 边线 {mk['line']} ≠ 调色板 {want}")
    else:
        for ser in obj._chartSpace.iter(qn("c:ser")):
            sname = _ser_name(ser)
            if sname not in order:
                continue
            spPr = ser.find(qn("c:spPr"))
            srgb = spPr.find(".//" + qn("a:srgbClr")) if spPr is not None else None
            if srgb is None:
                continue
            want = palette[order[sname]] if order[sname] < len(palette) else None
            if want is None or f"#{srgb.get('val').upper()}" != want:
                f.add("HARD", "V-CHART", name, f"系列「{sname}」颜色 #{srgb.get('val')} ≠ 调色板 {want}")

    hi = spec.get("highlight_category")
    if hi is not None and kind == "bar":
        want = palette[0]
        for dpt in obj._chartSpace.iter(qn("c:dPt")):
            idx = dpt.find(qn("c:idx"))
            srgb = dpt.find(".//" + qn("a:srgbClr"))
            if idx is None or srgb is None:
                continue
            if int(idx.get("val")) == int(hi) and f"#{srgb.get('val').upper()}" != want:
                f.add("HARD", "V-CHART", name, f"强调类别的颜色 ≠ 调色板首色 {want}")

    if kind == "bar" and spec.get("gap_width"):
        got = obj.plots[0].gap_width
        if got is not None and abs(int(got) - int(spec["gap_width"])) > 1:
            f.add("SOFT", "V-CHART", name, f"柱宽间隙 {got} ≠ plan {spec['gap_width']}")

    if kind in ("pie", "donut"):
        want_legend = M.chart_flags(spec)["legend"]
        if bool(obj.has_legend) != want_legend:
            f.add("SOFT", "V-CHART", name, f"图例 {bool(obj.has_legend)} ≠ plan {want_legend}")
        _verify_labels(plan, page, el, obj, name, f)
        return
    axis_on = spec.get("axis_labels", True)
    try:
        if bool(obj.value_axis.visible) != bool(axis_on):
            f.add("SOFT", "V-CHART", name, f"数值轴显示 {bool(obj.value_axis.visible)} ≠ plan {bool(axis_on)}")
    except (AttributeError, ValueError):
        # 读不到就说读不到：静默 pass 会让"没查"和"查过没事"长得一模一样。
        f.add("INFO", "I-AXIS", name, "读不到该类图表的数值轴显示状态，轴开关未核对")

    cfg = spec.get("value_axis") or {}
    va = obj.value_axis
    for key, attr in (("min", "minimum_scale"), ("max", "maximum_scale"), ("major_unit", "major_unit")):
        if cfg.get(key) is None:
            continue
        got = getattr(va, attr, None)
        if got is None or abs(float(got) - float(cfg[key])) > 1e-6:
            f.add("HARD", "V-CHART", name, f"数值轴 {key} {got} ≠ plan {cfg[key]}")
    if cfg.get("min") is None and spec.get("kind") in ("bar", "area") and H._all_non_negative(spec):
        got = getattr(va, "minimum_scale", None)
        if got is None or abs(float(got)) > 1e-9:
            f.add("HARD", "V-CHART", name, f"柱/面积图数值轴起点 {got}：未声明截轴时必须从 0 起，否则差异被无声放大")

    titles = spec.get("axis_titles") or {}
    for key, axis in (("y", va), ("x", getattr(obj, "category_axis", None))):
        if not titles.get(key) or axis is None:
            continue
        try:
            got = axis.axis_title.text_frame.text if axis.has_title else ""
        except (AttributeError, ValueError):
            f.add("INFO", "I-AXIS", name, f"读不到 {key} 轴名，声明了「{titles[key]}」但未核对")
            continue
        if got != str(titles[key]):
            f.add("HARD", "V-CHART", name, f"{key} 轴名「{got}」≠「{titles[key]}」")

    combo = spec.get("combo_line_series") or []
    if combo:
        line = obj._chartSpace.find(".//" + qn("c:lineChart"))
        got = [_ser_name(ser) for ser in line.findall(qn("c:ser"))] if line is not None else []
        if sorted(got) != sorted(str(c) for c in combo):
            f.add("HARD", "V-CHART", name, f"折线层系列 {got} ≠ plan {combo}")

    _verify_labels(plan, page, el, obj, name, f)
    want_legend = M.chart_flags(spec)["legend"]
    if bool(obj.has_legend) != want_legend:
        f.add("SOFT", "V-CHART", name, f"图例 {bool(obj.has_legend)} ≠ plan {want_legend}")


def _verify_labels(plan, page, el, obj, name: str, f: M.Report) -> None:
    """成品直标 == chart_label_plan 的决策：开关、位置、墨色、衬底、逐扇区墨色、单位、类别名。

    这是“plan → 成品”的一致性；读者实际读不读得清由 chart_audit 另算，两者互不替代。
    """
    spec = el["chart"]
    kind = spec["kind"]
    surface = M.chart_surface(page, el)
    entries = M.chart_label_plan(plan, spec, surface)
    by_name = {e["name"]: e for e in entries}
    combo = set(spec.get("combo_line_series") or [])
    unit = M.chart_unit(spec)
    for ser in obj._chartSpace.iter(qn("c:ser")):
        sname = _ser_name(ser)
        entry = entries[0] if kind in ("pie", "donut") else by_name.get(sname)
        want_on = bool(spec.get("combo_line_labels")) if sname in combo else M.chart_flags(spec)["data_labels"]
        dl = ser.find(qn("c:dLbls"))
        if dl is None:  # 他人或旧版写法：直标挂在 plot 级，系列继承
            dl = ser.getparent().find(qn("c:dLbls"))
        got_on = dl is not None and (dl.find(qn("c:showVal")) is not None
                                     and dl.find(qn("c:showVal")).get("val") in ("1", "true"))
        if got_on != want_on:
            f.add("HARD", "V-CHART", name, f"系列「{sname}」直标 {got_on} ≠ plan {want_on}")
        if not want_on or entry is None or dl is None:
            continue
        pos = dl.find(qn("c:dLblPos"))
        got_pos = None if pos is None else pos.get("val")
        if got_pos != entry["position"]:
            f.add("SOFT", "V-CHART", name, f"系列「{sname}」直标位置 {got_pos} ≠ {entry['position']}")
        ink = A._ink(dl.find(qn("c:txPr")))
        if ink != entry["ink"]:
            f.add("HARD", "V-LABEL-INK", name, f"系列「{sname}」直标墨色 {ink} ≠ 决策表 {entry['ink']}")
        plate = A._srgb(dl.find(qn("c:spPr")))
        if plate != entry["plate"]:
            f.add("HARD", "V-LABEL-INK", name, f"系列「{sname}」直标衬底 {plate} ≠ 决策表 {entry['plate']}")
        fmt = dl.find(qn("c:numFmt"))
        code = "" if fmt is None else (fmt.get("formatCode") or "")
        if unit and unit not in code and spec.get("stacked") != "percent" and "%" not in code:
            f.add("HARD", "V-CHART", name, f"系列「{sname}」直标数字格式「{code or '未设置'}」丢掉了单位「{unit}」")
        if kind in ("pie", "donut"):
            cat = dl.find(qn("c:showCatName"))
            got_cat = cat is not None and cat.get("val") in ("1", "true")
            if got_cat != (kind == "pie"):
                f.add("SOFT", "V-CHART", name, f"类别名开关 {got_cat} ≠ {kind == 'pie'}"
                                               f"（{'饼图外置直标要带类别名' if kind == 'pie' else '环带内只留数值，类别交给图例'}）")
        for k, want in enumerate(entry["points"] or []):
            lbl = next((d for d in dl.findall(qn("c:dLbl")) if d.find(qn("c:idx")).get("val") == str(k)), None)
            got = A._ink(lbl.find(qn("c:txPr"))) if lbl is not None else None
            if got != want:
                f.add("HARD", "V-LABEL-INK", name, f"环图第 {k + 1} 个扇区标签墨色 {got} ≠ 决策表 {want}")
            fmt = None if lbl is None else lbl.find(qn("c:numFmt"))
            if unit and (fmt is None or unit not in (fmt.get("formatCode") or "")):
                f.add("HARD", "V-CHART", name, f"环图第 {k + 1} 个标签的数字格式丢掉了单位「{unit}」")


def verify_page_number(plan, page, book, index, pid, f: M.Report) -> None:
    cfg = M.page_numbers(plan)
    pn = book["index"].get("_page_number")
    if pid in cfg["hide"]:
        if pn is not None:
            f.add("HARD", "V-PAGENO", pid, "声明隐藏页码，但成品里有页码对象")
        return
    label = str(cfg["start"] + index)
    if pn is None:
        f.add("HARD", "V-PAGENO", pid, "缺少页码对象")
        return
    if pn.text_frame.text != label:
        f.add("HARD", "V-PAGENO", pid, f"页码文字「{pn.text_frame.text}」≠「{label}」")
    order = book["order"].get("slide", [])
    if order and order[-1] != "_page_number":
        f.add("SOFT", "V-PAGENO", pid, "页码不在最上层：压到角区的色块会把它盖住")
    box = M.pn_box(plan, page)
    pos = (_in(pn.left), _in(pn.top))
    if abs(pos[0] - box["x"]) > 0.02 or abs(pos[1] - box["y"]) > 0.02:
        f.add("SOFT", "V-PAGENO", pid, f"页码位置 {pos} ≠ 保留框 ({box['x']},{box['y']})")
    runs = [r for p in pn.text_frame.paragraphs for r in p.runs]
    if runs and runs[0].font.size and abs(runs[0].font.size.pt - M.LIMITS["page_number"]["font_pt"]) > 0.5:
        f.add("SOFT", "V-PAGENO", pid, f"页码字号 {runs[0].font.size.pt}pt 与整稿约定不符")
    if runs:
        from pptx.enum.dml import MSO_COLOR_TYPE
        want_ink = M.pn_ink(plan, page)
        color = runs[0].font.color
        if color is None or color.type != MSO_COLOR_TYPE.RGB:
            f.add("HARD", "V-PAGENO", pid, "页码墨色不是显式 RGB")
        elif f"#{color.rgb}".upper() != want_ink:
            f.add("HARD", "V-PAGENO", pid, f"页码墨色 #{color.rgb} ≠ 按实际落点算的 {want_ink}")


def verify_structure(plan: Dict[str, Any], pptx: Path, facts: Dict[str, Any], f: M.Report) -> Dict[str, Any]:
    """文件结构：图表有对应的原生图表部件与内嵌工作簿，图片有对应的 media 部件。

    这一层查的是"文件里到底有什么"，对象核对查的是"python-pptx 读出来对不对"。
    一个图表被写成图片、内嵌工作簿丢失导致无法编辑数据，都只有在这一层才看得见。
    """
    import zipfile

    with zipfile.ZipFile(pptx) as zf:
        names = zf.namelist()
    charts = sum(1 for p in plan["pages"] for el in M.all_elements(p)
                 if el.get("type") == "chart" and (el.get("chart") or {}).get("kind") != "table")
    chart_parts = [n for n in names if n.startswith("ppt/charts/chart") and n.endswith(".xml")]
    books = [n for n in names if n.startswith("ppt/embeddings/")]
    media = [n for n in names if n.startswith("ppt/media/")]
    slides = [n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    distinct_images = len(set((facts.get("images") or {}).values()))

    if len(slides) != len(plan["pages"]):
        f.add("HARD", "V-PARTS", "结构", f"slide 部件 {len(slides)} 个 ≠ 页数 {len(plan['pages'])}")
    if len(chart_parts) != charts:
        f.add("HARD", "V-PARTS", "结构",
              f"原生图表部件 {len(chart_parts)} 个 ≠ 声明的图表 {charts} 个（图表可能被写成图片）")
    if len(books) < len(chart_parts):
        f.add("HARD", "V-PARTS", "结构",
              f"内嵌工作簿 {len(books)} 个 < 图表 {len(chart_parts)} 个（数据不可编辑）")
    if distinct_images and len(media) < distinct_images:
        f.add("HARD", "V-PARTS", "结构",
              f"media 部件 {len(media)} 个 < 不同的图片字节 {distinct_images} 份（有图片没进文件）")
    return {"slides": len(slides), "charts": len(chart_parts), "embeddings": len(books),
            "media": len(media), "distinct_images": distinct_images}


# ---------------------------------------------------------------- 渲染核对


def find_soffice() -> Optional[str]:
    """优先 SOFFICE 环境变量（便携版、AppImage 解包版），再找 PATH 与常见安装位置。"""
    from shutil import which
    override = os.environ.get("SOFFICE")
    if override and Path(override).is_file():
        return override
    for name in ("soffice", "libreoffice"):
        for ext in ("", ".exe"):
            found = which(name + ext)
            if found:
                return found
    for path in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                 "/usr/bin/soffice", "/usr/local/bin/soffice"):
        if Path(path).is_file():
            return path
    return None


def render(pptx: Path, render_dir: Path, f: M.Report) -> Optional[Path]:
    soffice = find_soffice()
    if soffice is None:
        f.add("INFO", "V-RENDER", "环境", "找不到 LibreOffice；页尺寸、断行、字体替换与溢出均未核验")
        return None
    try:
        import pymupdf  # noqa: F401
    except ImportError:
        f.add("INFO", "V-RENDER", "环境", "找不到 PyMuPDF（pip install -r requirements-render.txt）；渲染核对跳过")
        return None
    render_dir.mkdir(parents=True, exist_ok=True)
    for stale in render_dir.glob("*.pdf"):
        if stale.stem == pptx.stem:
            stale.unlink()
    env = dict(os.environ)
    env.pop("PYTHONPATH", None)
    # 独立用户配置目录：避免与用户正在运行的 LibreOffice 抢锁，也让并发核验互不影响。
    import tempfile
    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile:
        cmd = [soffice, "--headless", "--norestore", f"-env:UserInstallation={Path(profile).as_uri()}",
               "--convert-to", "pdf", "--outdir", str(render_dir), str(pptx)]
        ok, err = _run_soffice(cmd, env)
    if not ok:
        f.add("INFO", "V-RENDER", "环境", f"LibreOffice 调用失败：{err}")
        return None
    pdf = render_dir / (pptx.stem + ".pdf")
    if not pdf.is_file():
        f.add("INFO", "V-RENDER", "环境", "LibreOffice 未产出 PDF（可能是文件被占用）")
        return None
    return pdf


def environment_facts(rendered: bool) -> Dict[str, Any]:
    """验证等级的事实：在哪个环境、用什么渲染。PowerPoint 只能由真实 Windows/macOS 环境证明，这里永远不替它说话。"""
    import platform
    info: Dict[str, Any] = {"os": platform.platform(), "python": platform.python_version(),
                            "renderer": "none", "powerpoint": "NOT VERIFIED"}
    soffice = find_soffice() if rendered else None
    if soffice:
        try:
            out = subprocess.run([soffice, "--version"], capture_output=True, text=True, timeout=60).stdout.strip()
        except (OSError, subprocess.TimeoutExpired):
            out = "LibreOffice（版本未知）"
        info["renderer"] = out or "LibreOffice"
    return info


def _run_soffice(cmd: List[str], env: Dict[str, str]) -> Tuple[bool, str]:
    try:
        subprocess.run(cmd, capture_output=True, timeout=240, env=env, check=False)
        return True, ""
    except (OSError, subprocess.TimeoutExpired) as exc:
        return False, str(exc)


def _declared_families(plan, page) -> set:
    """这一页**声明过**的全部字族：元素级、run 级、图表、页码。

    只收元素级会漏掉"某个 run 单独换了字族"——那样的字在目标机上缺失时，
    渲染核对不会报，因为声明集合里根本没有它。
    """
    families = set()
    for el in M.all_elements(page):
        if el.get("type") in TEXT_TYPES:
            families.add(M.text_style(plan, page, el)["family"])
            for run in el.get("runs") or []:
                if run.get("font"):
                    families.add(str(run["font"]))
        elif el.get("type") in ("chart", "legend"):
            families.add(el.get("font") or (plan.get("fonts") or {}).get("body") or M.LIMITS["defaults"]["fonts"]["body"])
    if page["id"] not in M.page_numbers(plan)["hide"]:
        families.add((plan.get("fonts") or {}).get("body") or M.LIMITS["defaults"]["fonts"]["body"])
    return families


def _font_present(fam: str, fonts_here, aliases: Dict[str, List[str]]) -> bool:
    """声明的字族名（微软雅黑）→ 渲染器报的字体文件名（microsoftyahei），按别名表对上。"""
    keys = aliases.get(fam.lower()) or [fam.lower().replace(" ", "")]
    return any(any(key in rendered for key in keys) for rendered in fonts_here)


def _overlap_area(a, b) -> float:
    ix = max(0.0, min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]))
    return ix * iy


def _norm_key(text: str) -> str:
    """比对用的归一化：去掉全部空白（含全角空格与不换行空格）。

    字族被替换时断行位置会变，但字不会多也不会少——比字，不比空。
    """
    return "".join(str(text).split())


def _split_line_segments(line: Dict[str, Any]) -> List[List[Dict[str, Any]]]:
    """把 PyMuPDF 的一行按字号比（>1.25）或水平间隙（>0.6 em）切成段；普通行原样返回一段。"""
    spans = [sp for sp in line.get("spans", []) if sp.get("text", "") != ""]
    if not spans:
        return []
    segments: List[List[Dict[str, Any]]] = [[spans[0]]]
    for prev, cur in zip(spans, spans[1:]):
        ps, cs = float(prev.get("size") or 0) or 1.0, float(cur.get("size") or 0) or 1.0
        gap = cur["bbox"][0] - prev["bbox"][2]
        if max(ps, cs) / min(ps, cs) > 1.25 or gap > 0.6 * max(ps, cs):
            segments.append([cur])
        else:
            segments[-1].append(cur)
    return segments


def _render_text(pdf_page) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """一次遍历取出渲染里的行与 span。

    行用于文字完整性与排印判断，span 用于越界与字族判断；分两次 get_text 只是
    把同一份 PDF 解析两遍。坐标一律换算成英寸：调用方比的是 plan 的英寸框，
    混着点与英寸比会永远"不在框内"，检查也就永远不会响。
    """
    lines: List[Dict[str, Any]] = []
    spans: List[Dict[str, Any]] = []
    for block in pdf_page.get_text("dict")["blocks"]:
        blines = block.get("lines", [])
        for i, line in enumerate(blines):
            direction = tuple(line.get("dir", (1.0, 0.0)))
            # 同一基线上字号悬殊（巨型数字＋小单位）或间隔很大的 span，PyMuPDF 会并成一行，
            # 中心落到别的元素框里，归属就错了（12.1 前把“24.6｜亿元”判成单位缺失）。按字号/间隙切成独立段。
            for seg in _split_line_segments(line):
                text = "".join(span.get("text", "") for span in seg)
                if not text.strip():
                    continue
                x0 = min(sp["bbox"][0] for sp in seg); y0 = min(sp["bbox"][1] for sp in seg)
                x1 = max(sp["bbox"][2] for sp in seg); y1 = max(sp["bbox"][3] for sp in seg)
                lines.append({"text": text, "dir": direction,
                              "bbox": (x0 / 72, y0 / 72, x1 / 72, y1 / 72),
                              "last": i == len(blines) - 1, "n_lines": len(blines)})
            for span in line.get("spans", []):
                sx0, sy0, sx1, sy1 = span.get("bbox") or line["bbox"]
                if span.get("origin") and span.get("size"):
                    # span.bbox 是字体 ascent+descent 的行盒（Noto CJK 约 1.45 em），巨型数字会比墨迹高出一截而被误报越框；
                    # 越界判断按墨迹范围：基线上 0.95 em、基线下 0.2 em 之内。
                    oy, sz = span["origin"][1], span["size"]
                    sy0, sy1 = max(sy0, oy - 0.95 * sz), min(sy1, oy + 0.2 * sz)
                spans.append({"text": span.get("text", ""), "font": span.get("font", ""),
                              "bbox": (sx0 / 72, sy0 / 72, sx1 / 72, sy1 / 72)})
    return lines, spans


def _inside(box: Tuple[float, float, float, float], bbox, tol: float) -> bool:
    cx, cy = (bbox[0] + bbox[2]) / 2, (bbox[1] + bbox[3]) / 2
    return (box[0] - tol <= cx <= box[0] + box[2] + tol
            and box[1] - tol <= cy <= box[1] + box[3] + tol)


def expected_text(el: Dict[str, Any], page: Optional[Dict[str, Any]] = None) -> List[str]:
    """成品上应该读到的字，按"块"返回。

    文字元素是整段：段内换行由渲染决定，去掉空白后仍是同一串字。
    图表与表格必须逐块比（图题、每个类别、每个单元格）：它们的阅读顺序由渲染器
    决定，拼成一整串再去比对，字都在也会因为顺序不同而判成缺失。
    """
    kind = el.get("type")
    if kind in TEXT_TYPES:
        if el.get("list"):
            # 原生列表：渲染器会在每项前插入符号/编号，整段拼接比不上；逐项比。
            return [part for part in str(el.get("content") or "").split("\n") if part.strip()]
        return [str(el.get("content") or "")]
    if kind == "legend":
        # 图例文字在组合里：逐条比标签，缺一条就是错色或漏条目。
        target = None
        if page is not None:
            target = next((e for e in page.get("elements") or []
                           if str(e.get("id")) == str(el.get("for_chart"))), None)
        if target is None or target.get("type") != "chart":
            return []
        spec = target["chart"]
        values = list(((spec.get("data") or {}).get("series") or [{}])[0].get("values") or [])
        show_value = bool(el.get("show_value")) and spec.get("kind") in ("pie", "donut")
        return [H._legend_label(spec, str(item), values, i, show_value)
                for i, item in enumerate(M.chart_items(spec))]
    chart = el.get("chart") or {}
    if kind != "chart":
        return []
    out: List[str] = []
    # 表格不画图题（checks.py 的 W-CHART 已提示），核验也不能要求渲染里有它。
    if chart.get("kind") != "table" and chart.get("show_title", True) and chart.get("title"):
        out.append(str(chart["title"]))
    data = chart.get("data") or {}
    # 类别只在真的会画出来时才算预期文字：轴关了、环图没图例，类别就不会出现在渲染里，
    # 一律要求会让核对永远误报（旧版对环图 + 手绘图例报 V-TEXT-RENDER 就是这么来的）。
    kind_c = chart.get("kind")
    if kind_c in ("pie", "donut"):
        flags = M.chart_flags(chart)
        if flags["legend"] or (kind_c == "pie" and flags["data_labels"]):
            out.extend(str(c) for c in data.get("categories") or [])
    elif chart.get("axis_labels", True):
        out.extend(str(c) for c in data.get("categories") or [])
    out.extend(str(v) for v in data.get("headers") or [])
    for row in data.get("rows") or []:
        out.extend(str(v) for v in row)
    # 与 compile 一致：legend 缺省时两个以上系列才显示图例；显式 false 就不会渲染系列名。
    show_legend = M.chart_flags(chart)["legend"]
    # 饼与环的图例列的是类别而不是系列名，系列名不会出现在渲染里。
    for series in data.get("series") or []:
        if show_legend and chart.get("kind") not in ("pie", "donut"):
            out.append(str(series.get("name") or ""))
    out.extend(_label_pieces(chart))
    return [piece for piece in out if _norm_key(piece)]


def _label_pieces(chart: Dict[str, Any]) -> List[str]:
    """直标的预期文字：默认数字格式（#,##0.###）加图表级单位，与 compile 的 _label_format 对齐。

    自定义数字格式、百分比堆叠与饼环不在此核对（它们的文字规则不同），留给人看渲染。
    """
    if chart.get("kind") not in ("bar", "line") or not chart.get("data_labels"):
        return []
    if (chart.get("value_axis") or {}).get("number_format") or chart.get("stacked") == "percent":
        return []
    unit = M.chart_unit(chart)
    out: List[str] = []
    # 组合折线层只有声明 combo_line_labels 才有直标（compile 同一规则）。
    skip = set() if chart.get("combo_line_labels") else set(chart.get("combo_line_series") or [])
    for series in (chart.get("data") or {}).get("series") or []:
        if series.get("name") in skip:
            continue
        for value in series.get("values") or []:
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                out.append(f"{float(value):,.3f}".rstrip("0").rstrip(".") + unit)
    return out


def _owners(lines: List[Dict[str, Any]], boxes: Dict[str, Tuple[float, float, float, float]]
            ) -> Dict[str, List[Dict[str, Any]]]:
    """每一行渲染文字只归属一个元素：中心落在的最小声明框。

    一行字若同时落在大框和小框里，两边都算，大框的核对就会借小框的字过关。
    """
    tol = M.LIMITS["render_check"]["span_tol_in"]
    owned: Dict[str, List[Dict[str, Any]]] = {}
    for line in lines:
        hits = [(w * h, name) for name, (bx, by, w, h) in boxes.items()
                if _inside((bx, by, w, h), line["bbox"], tol)]
        if hits:
            owned.setdefault(min(hits)[1], []).append(line)
    return owned


def _check_render_text(page, lines, boxes, page_key: str, pid: str, f: M.Report) -> int:
    """预期文字 vs 渲染文字，逐块核对。

    只查"框有没有超出"是不够的：字体缺失吞字、渲染器省略长文本、上层色块盖住文字，
    都能让框看起来正好、字却没了。这里比的是字，不是框。
    """
    checks = 0
    owned = _owners(lines, boxes)
    for el in M.all_elements(page):
        pieces = [p for p in expected_text(el, page) if _norm_key(p)]
        if not pieces:
            continue
        pool = _norm_key("".join(ln["text"] for ln in owned.get(str(el.get("id")), [])))
        missing = [p for p in pieces if _norm_key(p) not in pool]
        checks += 1
        if not missing:
            continue
        where = f"{pid}/{el.get('id')}"
        detail = "、".join(f"「{p[:12]}」" for p in missing[:3])
        elsewhere = any(_norm_key(p) in page_key for p in missing)
        why = "（页面别处有这些字，但不在本元素框内：多半是溢出或被挪位）" if elsewhere else ""
        if el.get("type") in TEXT_TYPES:
            f.add("HARD", "V-TEXT-RENDER", where,
                  f"渲染里读不到本元素的文字：{detail}（共缺 {len(missing)}/{len(pieces)} 块）{why}。"
                  f"查字体是否存在、是否被上层遮住；不要靠缩字号掩盖")
        else:
            f.add("SOFT", "V-TEXT-RENDER", where,
                  f"图表/表格里读不到本元素的文字：{detail}（共缺 {len(missing)}/{len(pieces)} 块）{why}")
    return checks


def _check_typography(page, lines, boxes, pid: str, f: M.Report) -> None:
    """渲染实测的排印缺陷：行首收尾标点、行尾起始标点、末行孤字。

    这三条只有断行之后才能判断——计划里是整段文字，断点由字族、字距与框宽共同决定。
    全部记 SOFT：孤字也可能是刻意的短行，由人确认。

    孤字按**元素**归组判断，不按 PDF 的 block 判断：LibreOffice 常把同一文本框的
    每一行写成独立 block，按 block 取"最后一行"永远只有一行，检查会永远沉默。
    """
    lead_bad = "，。、；：？！）〕」』】》"
    tail_bad = "（〔「『【《"
    groups: Dict[str, List[Dict[str, Any]]] = {}
    for line in lines:
        text = line["text"].strip()
        if not text:
            continue
        tol = M.LIMITS["render_check"]["line_tol_in"]
        host = next((name for name, box in boxes.items() if _inside(box, line["bbox"], tol)), None)
        where = f"{pid}/{host}" if host else pid
        if text[0] in lead_bad:
            f.add("SOFT", "V-PUNCT", where, f"行首是收尾标点「{text[:8]}」：断点要往前挪一个字")
        if text[-1] in tail_bad:
            f.add("SOFT", "V-PUNCT", where, f"行尾是起始标点「{text[-8:]}」：断点落在括号前了")
        if host is not None:
            groups.setdefault(host, []).append(line)

    for host, rows in groups.items():
        horizontal = [ln for ln in rows if ln["dir"][0] > 0.9]
        if len(horizontal) < 2:
            continue
        horizontal.sort(key=lambda ln: (round(ln["bbox"][1], 1), ln["bbox"][0]))
        last = horizontal[-1]["text"].strip()
        total = _norm_key("".join(ln["text"] for ln in horizontal))
        min_chars = M.LIMITS["render_check"]["orphan_min_chars"]
        if len(last) == 1 and ord(last) >= 0x2E80 and len(total) >= min_chars:
            f.add("SOFT", "V-ORPHAN", f"{pid}/{host}",
                  f"末行只剩「{last}」一个字：改文案或改框宽，不要用空格推")


def _check_render_rotation(page, lines, boxes, pid: str, f: M.Report) -> None:
    """旋转是否真的进了渲染：XML 里写了 rot、画面上字却没转向，是常见的半失效。"""
    import math

    for el in M.all_elements(page):
        if el.get("type") not in TEXT_TYPES:
            continue
        rot = float(el.get("rotation") or 0) % 360
        if min(rot, 360 - rot) < M.LIMITS["render_check"]["rotation_min_deg"]:
            continue
        box = boxes.get(str(el.get("id")))
        if box is None:
            continue
        here = [ln for ln in lines if _inside(box, ln["bbox"], 0.25)]
        if not here:
            continue
        want = (math.cos(math.radians(rot)), math.sin(math.radians(rot)))
        got = max((ln["dir"] for ln in here), key=lambda d: abs(d[0]) + abs(d[1]))
        diff = math.degrees(math.acos(max(-1.0, min(1.0, want[0] * got[0] + want[1] * got[1]))))
        if diff > M.LIMITS["render_check"]["rotation_tol_deg"]:
            f.add("SOFT", "V-ROT", f"{pid}/{el.get('id')}",
                  f"声明旋转 {rot:.0f}°，渲染字向 {round(got[0], 2), round(got[1], 2)}（差 {diff:.0f}°）："
                  f"可能只转了框，没有转字")


def verify_render(plan: Dict[str, Any], pptx: Path, pdf: Path, render_dir: Path, f: M.Report) -> Dict[str, Any]:
    import pymupdf
    from PIL import Image

    doc = pymupdf.open(str(pdf))
    w_in, h_in = M.canvas(plan)
    dpi = M.LIMITS["preview"]["dpi"]
    facts: Dict[str, Any] = {"pages": doc.page_count, "dpi": dpi, "pngs": [], "fonts": {}, "text_checks": 0}
    if doc.page_count != len(plan["pages"]):
        f.add("HARD", "V-PAGES", "渲染", f"PDF {doc.page_count} 页 ≠ plan {len(plan['pages'])} 页")

    pngs: List[Path] = []
    all_fonts: set = set()
    face_notes: List[str] = []
    face_seen: set = set()
    face_states: set = set()
    aliases = {k.lower(): [a.lower() for a in v] for k, v in M.LIMITS.get("font_aliases", {}).items()
               if not k.startswith("_")}
    for index, pdf_page in enumerate(doc):
        spec_page = plan["pages"][index] if index < len(plan["pages"]) else {"id": f"p{index + 1}", "elements": []}
        pid = spec_page["id"]
        # 声明框按旋转后的外接框算：文字转 -90° 之后，字的落点本来就不在原框里。
        # 用原框做归属会把它判成"不在任何元素范围内"——那是量错了，不是设计错了。
        boxes: Dict[str, Tuple[float, float, float, float]] = {}
        for el in M.all_elements(spec_page):
            if el.get("type") in TEXT_TYPES + ("chart", "legend") and "x" in el:
                boxes[str(el.get("id"))] = M.rot_box(
                    (float(el["x"]), float(el["y"]), float(el["w"]), float(el["h"])), el.get("rotation"))
        page_w = round(pdf_page.rect.width / 72, 3)
        page_h = round(pdf_page.rect.height / 72, 3)
        if abs(page_w - w_in) > 0.05 or abs(page_h - h_in) > 0.05:
            f.add("HARD", "V-PAGES", pid, f"页面尺寸 {page_w}×{page_h}in ≠ 画布 {w_in}×{h_in}in")

        declared = _declared_families(plan, spec_page)
        cfg = M.page_numbers(plan)
        pn = None if pid in cfg["hide"] else M.pn_box(plan, spec_page)

        pix = pdf_page.get_pixmap(dpi=dpi)
        png_path = render_dir / f"p{index + 1:02d}.png"
        pix.save(str(png_path))
        pngs.append(png_path)

        with Image.open(png_path) as img:
            gray = img.convert("L").resize((320, 180))
            ink = sum(1 for v in gray.tobytes() if v < 245) / max(320 * 180, 1)
            if ink < M.LIMITS["preview"]["blank_ink_ratio"]:
                f.add("SOFT", "V-BLANK", pid, f"页面近乎空白（墨迹占比 {ink:.4%}）")

        page_box = (0.0, 0.0, page_w, page_h)
        fonts_here: set = set()
        lines, spans = _render_text(pdf_page)
        for span in spans:
            text = span["text"].strip()
            if not text:
                continue
            fonts_here.add(font_family_base(span["font"]))
            sx0, sy0, sx1, sy1 = span["bbox"]
            if sx0 < -0.02 or sy0 < -0.02 or sx1 > page_box[2] + 0.02 or sy1 > page_box[3] + 0.02:
                # 元素自己声明了出血（框伸出画布，compile 已给 W-GEOM）：被裁是设计，只提示；否则是漏出。
                declared_bleed = any(
                    (bx < 0 or by < 0 or bx + bw > page_box[2] or by + bh > page_box[3])
                    and bx <= (sx0 + sx1) / 2 <= bx + bw and by <= (sy0 + sy1) / 2 <= by + bh
                    for bx, by, bw, bh in boxes.values())
                if declared_bleed:
                    f.add("INFO", "V-BLEED", pid, f"文字按声明出血到页外：{text[:24]}（确认被裁的部分不含信息）")
                else:
                    f.add("HARD", "V-CLIP", pid, f"文字超出页面：{text[:24]}")
                continue
            tol = max(0.06, 0.15 * (sy1 - sy0))
            if any(sx0 >= bx - tol and sy0 >= by - tol and sx1 <= bx + bw + tol and sy1 <= by + bh + tol
                   for bx, by, bw, bh in boxes.values()):
                continue
            if pn is not None and pn["x"] - 0.05 <= sx0 and sx1 <= pn["x"] + pn["w"] + 0.05 \
                    and pn["y"] - 0.05 <= sy0 and sy1 <= pn["y"] + pn["h"] + 0.05:
                continue  # 页码由 V-PAGENO 单独查；按区域放过，不按"是不是数字"放过
            host = max(boxes.items(), key=lambda kv: _overlap_area(
                (sx0, sy0, sx1 - sx0, sy1 - sy0), kv[1]), default=None)
            if host is not None and _overlap_area((sx0, sy0, sx1 - sx0, sy1 - sy0), host[1]) > 0:
                f.add("SOFT", "V-OVERFLOW", pid,
                      f"「{text[:20]}」超出了 {host[0]} 的声明框（框太小或字体被替换）；"
                      f"按真实渲染放大框或改文案，不要压字号")
            else:
                f.add("SOFT", "V-OVERFLOW", pid, f"「{text[:20]}」不在任何声明元素的范围里")

        page_key = _norm_key("".join(ln["text"] for ln in lines))
        facts["text_checks"] += _check_render_text(spec_page, lines, boxes, page_key, pid, f)
        _check_typography(spec_page, lines, boxes, pid, f)
        _check_render_rotation(spec_page, lines, boxes, pid, f)

        if pn is not None:
            want = str(cfg["start"] + index)
            box = pn
            hit = False
            for word in pdf_page.get_text("words"):
                if word[4].strip() == want:
                    cx, cy = (word[0] + word[2]) / 2 / 72, (word[1] + word[3]) / 2 / 72
                    if box["x"] - 0.35 <= cx <= box["x"] + box["w"] + 0.35 and \
                       box["y"] - 0.2 <= cy <= box["y"] + box["h"] + 0.2:
                        hit = True
                        break
            if not hit:
                f.add("HARD", "V-PAGENO", pid, f"渲染里角区找不到页码「{want}」")

        facts["fonts"][pid] = sorted(fonts_here)
        all_fonts |= fonts_here
        # 一句话说清"声明的没出现、出现的没声明"：分成两条只会把同一件事说两遍。
        extra = sorted(fam for fam in fonts_here
                       if not any(_font_present(d, {fam}, aliases) for d in declared))
        missing = sorted(d for d in declared if not _font_present(d, fonts_here, aliases))
        if extra or missing:
            facts["font_substituted"] = True
            f.add("SOFT", "V-FONT", pid,
                  f"声明 {sorted(declared)}，渲染实际 {sorted(fonts_here)}"
                  f"{'；未出现 ' + str(missing) if missing else ''}"
                  f"{'；多出 ' + str(extra) if extra else ''}"
                  f"——该页断行只代表替代字体，不能当作目标机结论")

        # 12.0.0 字面核验：家族对上了不等于简体字形对上了。Noto/思源 CJK 以 .ttc 集合
        # 分发，LibreOffice 按集合首个字面导出（常为 jp），声明 SC 而实际取 JP 字面时，
        # 直、骨、门、者一类字形会按日文字形渲染——家族级"OK"不能读成"简体字形已核验"。
        want_script = _declared_script(declared)
        got_scripts = sorted({s for fam in fonts_here for s in _render_scripts(fam)} - {None})
        collection_only = _collection_suffix_only(fonts_here) and got_scripts == ["jp"]
        if want_script and got_scripts and want_script not in got_scripts and collection_only:
            if ("collection", want_script) not in face_seen:
                face_seen.add(("collection", want_script))
                f.add("SOFT", "W-FONT-FACE", pid,
                      f"声明 {SCRIPT_LABEL.get(want_script, want_script)}，渲染用的是 Noto/思源 CJK 的 .ttc 集合，"
                      f"PDF 字体名里的 jp 只是集合首面的固定后缀，无法判定实际字面——字面 NOT VERIFIED；"
                      f"需在目标机或装有独立 SC 字体文件的环境人工复核（直/骨/门/者）")
        elif want_script and got_scripts and want_script not in got_scripts:
            key = (want_script, tuple(got_scripts))
            if key not in face_seen:
                face_seen.add(key)
                face_notes.append(f"声明 {SCRIPT_LABEL.get(want_script, want_script)}，"
                                  f"PDF 里实际取到的字面是 "
                                  f"{SCRIPT_LABEL.get(got_scripts[0], got_scripts[0])}（{pid} 起）")
                f.add("SOFT", "W-FONT-FACE", pid,
                      f"声明 {SCRIPT_LABEL.get(want_script, want_script)} 字族，PDF 里实际取到的字面是 "
                      f"{SCRIPT_LABEL.get(got_scripts[0], got_scripts[0])}"
                      f"——家族匹配通过，但字面不同（{', '.join(got_scripts)}）；"
                      f"直/骨/门/者一类共享码位的字形可能按别的地区字面渲染，"
                      f"简体字形未被核验，目标机必须人工复核")

        # P0（13.5.0）：字面只在“页面确有中日韩文字”时才有意义；无法确认也是状态，不是通过。
        page_cjk = any("\u4e00" <= ch <= "\u9fff" for ln in lines for ch in ln["text"])
        if page_cjk:
            if want_script and got_scripts and want_script in got_scripts:
                face_states.add("VERIFIED")
            elif want_script and got_scripts and not collection_only:
                face_states.add("MISMATCH")
            elif want_script and got_scripts:
                face_states.add("UNVERIFIED")
            else:
                face_states.add("UNVERIFIED")
                if ("unverified", declared and tuple(sorted(declared))) not in face_seen:
                    face_seen.add(("unverified", declared and tuple(sorted(declared))))
                    f.add("SOFT", "W-FONT-FACE", pid,
                          f"页面含中文，但声明字族 {sorted(declared)} 或渲染字体 {sorted(fonts_here)} "
                          f"无法判定地区字面（SC/TC/JP/KR）——字面 NOT VERIFIED，状态不能是 READY；"
                          f"声明带地区后缀的字族（如 Noto Sans CJK SC），或在目标机人工复核")

    facts["fonts_all"] = sorted(all_fonts)
    if face_notes:
        f.add("INFO", "I-FONT-FACE", "deck",
              "；".join(face_notes) + "——「字体：OK」只到家族，不含字面")
    facts["font_face"] = ("MISMATCH" if "MISMATCH" in face_states else "UNVERIFIED" if "UNVERIFIED" in face_states
                          else "VERIFIED" if face_states else "NOT_APPLICABLE")
    facts["pngs"] = make_contact_sheet(pngs, render_dir)
    return facts


SCRIPT_LABEL = {"sc": "简体（SC/Hans）", "tc": "繁体（TC/Hant）", "jp": "日文（JP）", "kr": "韩文（KR）"}
_SCRIPT_MARKERS = (("cjkjp", "jp"), ("cjktc", "tc"), ("cjkkr", "kr"), ("cjksc", "sc"),
                   ("hansansjp", "jp"), ("hansanssc", "sc"), ("hansanskr", "kr"),
                   ("hansseriftc", "tc"), ("hansserifsc", "sc"))


def _script_of(name: str) -> Optional[str]:
    """字体名里的地区字面（sc/tc/jp/kr）；分不出来返回 None。

    先看集合名标记（cjkjp、hansanssc…），再看名字结尾的地区码——Noto CJK 的 .ttc 在
    PDF 里常是 notosanscjkjp-bold-vkana 这种带字重与变体后缀的形式，先按标记找才不会被
    后缀挡住，也不会把 notosanscjkjp 里的 "sc" 误当成简体。
    """
    key = "".join(ch for ch in str(name).lower() if ch.isascii() and (ch.isalnum() or ch == " "))
    key = key.replace(" ", "")
    if not key:
        return None
    for marker, code in _SCRIPT_MARKERS:
        if marker in key:
            return code
    for tail, code in (("hans", "sc"), ("hant", "tc"), ("sc", "sc"), ("tc", "tc"),
                       ("jp", "jp"), ("kr", "kr")):
        if key.endswith(tail):
            return code
    return None


def _collection_suffix_only(fonts_here) -> bool:
    """渲染字体全是 Noto/思源 CJK 的 .ttc 集合：LibreOffice 导出 PDF 时字体名固定带首个字面（jp），
    这个后缀不代表实际取到的字面——无法据此判"不一致"，只能判"无法确认"。"""
    keys = ["".join(ch for ch in str(f).lower() if ch.isalnum()) for f in fonts_here]
    return bool(keys) and all(any(t in k for t in ("notosanscjk", "notoserifcjk", "sourcehansans", "sourcehanserif"))
                              for k in keys)


def _declared_script(declared) -> Optional[str]:
    """整组声明里能确定的地区字面：都确定且一致才返回，混着或没写就返回 None（不误报）。"""
    found = sorted({s for s in (_script_of(d) for d in declared) if s})
    return found[0] if len(found) == 1 else None


def _render_scripts(fam: str) -> set:
    """渲染字体名里出现的地区字面集合。"""
    key = "".join(ch for ch in str(fam).lower() if ch.isascii() and (ch.isalnum() or ch == " "))
    key = key.replace(" ", "")
    out = set()
    for marker, code in _SCRIPT_MARKERS:
        if marker in key:
            out.add(code)
    if not out:
        tail = _script_of(fam)
        if tail:
            out.add(tail)
    return out


def font_family_base(name: str) -> str:
    base = name.split("+")[-1]
    for suffix in ("-BoldItalic", "-BoldOblique", "-Bold", "-Italic", "-Oblique", "-Regular",
                   "-Light", "-Medium", "-Semibold", "-Black", "-Heavy", "-Thin"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
    return base.lower().replace(" ", "")


def make_contact_sheet(pngs: List[Path], render_dir: Path) -> List[str]:
    from PIL import Image
    if not pngs:
        return []
    try:
        with Image.open(pngs[0]) as probe:
            ratio = probe.height / probe.width
    except OSError:
        return [str(p) for p in pngs]
    cols = 2 if len(pngs) <= 4 else 3
    thumb_w = M.LIMITS["preview"]["contact_sheet_width_px"] // cols
    thumb_h = int(thumb_w * ratio)
    rows = (len(pngs) + cols - 1) // cols
    pad = 24
    sheet = Image.new("RGB", (cols * thumb_w + pad * (cols + 1), rows * thumb_h + pad * (rows + 1)), "#E9E6DF")
    for i, path in enumerate(pngs):
        with Image.open(path) as img:
            thumb = img.convert("RGB").resize((thumb_w, thumb_h), Image.LANCZOS)
        x = pad + (i % cols) * (thumb_w + pad)
        y = pad + (i // cols) * (thumb_h + pad)
        sheet.paste(thumb, (x, y))
    out = render_dir / "contact_sheet.png"
    sheet.save(out)
    return [str(out)] + [str(p) for p in pngs]


# ---------------------------------------------------------------- 主流程


# 状态 ↔ 进程退出码：唯一映射。自动化只看退出码即可区分“可交付核验”“需补环境/人工”“有硬问题”。
# 2 保留给输入/契约错误（找不到文件、plan 不合契约、路径守卫），不是核验结论。
EXIT_BY_STATUS = {"READY": 0, "BLOCK": 1, "PARTIAL": 4}


def decide_status(blocked: bool, render_state: str, render_facts: Dict[str, Any]) -> Tuple[str, str]:
    """BLOCK / PARTIAL / READY 的唯一判定点，返回 (status, fonts_state)。

    BLOCK：任一层有硬问题。
    READY：对象一致 + 渲染完成且无硬问题 + 声明字体真实参与 + 字面（SC/TC/JP/KR）已确认或页面无中日韩文字。
    其余一律 PARTIAL：渲染没跑、字体被替换、字面不一致或无法确认。
    READY 仍不包含 PowerPoint 实机（NOT VERIFIED）与人工签收（NOT REVIEWED）。
    """
    if render_state != "PASS":
        fonts_state = "NOT CHECKED"
    elif render_facts.get("font_substituted"):
        fonts_state = "SUBSTITUTED"
    else:
        face = render_facts.get("font_face", "UNVERIFIED")
        fonts_state = {"VERIFIED": "OK", "NOT_APPLICABLE": "OK", "MISMATCH": "FACE MISMATCH"}.get(face, "FACE UNVERIFIED")
    if blocked:
        return "BLOCK", fonts_state
    return ("READY" if render_state == "PASS" and fonts_state == "OK" else "PARTIAL"), fonts_state


def main(argv: Optional[List[str]] = None) -> int:
    M.utf8_stdio()
    parser = argparse.ArgumentParser(description="核验 plan ↔ PPTX ↔ 渲染")
    parser.add_argument("plan")
    parser.add_argument("pptx")
    parser.add_argument("--assets", default=None, help="素材目录；须与编译时一致，缺省为 plan 同级 assets/ 或 plan 所在目录")
    parser.add_argument("--render-dir", default=None, help="写 PDF/PNG/缩略总览的目录")
    parser.add_argument("--report", default=None, help="核验报告 JSON；已存在时需 --force，且不得是 plan / pptx")
    parser.add_argument("--force", action="store_true", help="允许替换已存在的报告")
    args = parser.parse_args(argv)

    plan_path = Path(args.plan).resolve()
    pptx = Path(args.pptx).resolve()
    if not pptx.is_file():
        print(f"找不到 {pptx}")
        return 2
    rep = M.Report()
    render_arg = Path(args.render_dir).resolve() if args.render_dir else pptx.parent / "render" / pptx.stem
    M.guard_paths(rep, {"plan": plan_path, "pptx": pptx},
                  [("--report", Path(args.report) if args.report else None, ".json")], args.force,
                  dirs={"--render-dir": render_arg})
    if rep.blocked:
        M.show(rep.items)
        print("\n核验中止：路径会覆盖输入或工作目录，未写任何文件。")
        return 2
    plan = M.load_plan(plan_path, rep)
    if plan is None or rep.blocked:
        M.show(rep.items)
        print("\nplan 本身不合契约：先修 plan，核验才有意义。")
        return 2

    problems = M.package_problems(pptx)
    if problems:
        for problem in problems:
            print(f"[HARD] V-FORMAT {pptx}：{problem}")
        print("\n核验中止：成品不是合格的无宏 .pptx（PPTM 与扩展名不符的包都不接受）。")
        return 1

    f = M.Report()
    key = pptx.read_bytes()
    sha = hashlib.sha256(key).hexdigest()
    facts: Dict[str, Any] = {"pptx": str(pptx), "sha256": sha, "bytes": len(key), "render": {"done": False}}
    asset_root = M.default_asset_root(plan_path, args.assets)
    facts["asset_root"] = str(asset_root)

    # 两层各自出结论，不合并：合并之后"渲染没跑"会被"对象都对"盖掉。
    obj_rep = M.Report()
    facts.update(verify_objects(plan, pptx, asset_root, obj_rep))
    for issue in K.design_signals(plan).items:
        obj_rep.items.append(issue)
    facts["structure"] = verify_structure(plan, pptx, facts, obj_rep)

    # 每份 pptx 一个子目录：同一目录里核验两份稿时，PNG 与缩略总览不会互相覆盖。
    render_dir = render_arg
    render_rep = M.Report()
    pdf = render(pptx, render_dir, render_rep)
    if pdf is not None:
        facts["render"] = {"done": True, "pdf": str(pdf)}
        facts["render"].update(verify_render(plan, pptx, pdf, render_dir, render_rep))
    facts["environment"] = environment_facts(pdf is not None)
    f.extend(obj_rep)
    f.extend(render_rep)

    M.show(f.items)

    # READY = 对象一致 + 渲染完成 + 声明字体都真实参与渲染。任一缺失只能是 PARTIAL：
    # 渲染缺失或字体被替换时，断行与宽度都只代表替代字体，不是目标环境结论。
    objects = "FAIL" if obj_rep.blocked else "PASS"
    render_state = "SKIP" if pdf is None else ("FAIL" if render_rep.blocked else "PASS")
    status, fonts_state = decide_status(f.blocked, render_state, facts["render"])
    facts["layers"] = {"objects": objects, "render": render_state, "fonts": fonts_state,
                       "powerpoint": "NOT VERIFIED", "human": "NOT REVIEWED"}
    if fonts_state == "SUBSTITUTED":
        print("字体替换：声明字体未全部参与渲染，断行与宽度只代表替代字体，状态不能是 READY。")
    elif fonts_state.startswith("FACE"):
        print("字面未确认或不一致：家族匹配不等于地区字面一致，状态不能是 READY。")
    print(f"\n对象核对：{objects}（硬 {obj_rep.count('HARD')}）｜渲染核对：{render_state}｜字体：{fonts_state}"
          f"｜提示 {f.count('SOFT')}｜说明 {f.count('INFO')}｜状态：{status}")
    print("READY 只说明对象与渲染一致，不说明设计成立；PARTIAL 说明视觉判断没有任何机器依据。")
    st = facts.get("structure") or {}
    print(f"文件结构：slide {st.get('slides')}｜原生图表 {st.get('charts')}｜内嵌工作簿 {st.get('embeddings')}"
          f"｜media {st.get('media')}")
    brief = plan.get("brief") or {}
    if brief:
        print(f"概念：{str(brief.get('concept', ''))[:80]}")
        print(f"否决：{'；'.join(str(v) for v in brief.get('veto') or [])[:120]}")
    print(f"PPTX SHA-256：{sha}")
    print(f"验证环境：渲染 = {facts['environment']['renderer']}｜Microsoft PowerPoint = NOT VERIFIED（需由人在真实 PowerPoint 中复核；本包从未做过实机验证）")
    if facts["render"].get("done"):
        print(f"渲染：{facts['render']['pdf']}")
        print(f"渲染字体：{facts['render'].get('fonts_all')}")
        print(f"逐元素文字核对：{facts['render'].get('text_checks', 0)} 处"
              f"｜缩略总览：{render_dir / 'contact_sheet.png'}")
        print("机器能证明的到此为止：构图、主次、概念是否成立，必须自己看 PNG 并在目标 PowerPoint 复核。")
    else:
        print("未渲染：视觉相关判断一律没有依据。不要声称「已预览」「版面无溢出」「断行正常」。")

    code = EXIT_BY_STATUS[status]
    print(f"退出码：{code}（READY=0｜BLOCK=1｜PARTIAL=4｜输入或契约错误=2）")
    if args.report:
        M.write_atomic(Path(args.report), json.dumps({"status": status, "exit_code": code, **facts,
                                                      "findings": f.to_dict()}, ensure_ascii=False, indent=2).encode("utf-8"),
                       args.force)
    return code


if __name__ == "__main__":
    sys.exit(main())
