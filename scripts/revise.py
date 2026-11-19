#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
在既有 PPTX 上做定点修改：只动声明的对象，其余原样保留。

用法：
    python scripts/revise.py 原稿.pptx edit.json 输出.pptx [--assets DIR] [--report revise.json] [--force]

为什么不是重新编译：母版、动画、切换、SmartArt、嵌入对象、备注在重建里回不来。
改一页标题、挪一个框、换一张图、改一组数，都不该付这个代价。
本脚本不建对象，只在原文件上做声明的改动，再写成一个新文件（原稿不动）。

每次运行都会做保留核对：页数、母版、版式、主题字体、宏、逐页的动画/切换/SmartArt/
嵌入对象/备注、以及**未被点名的对象**是否与改动前逐项一致。核对结果写进 --report。
exit 0 = 写成功且保留核对通过；2 = 硬错误（未写文件）；3 = 运行期失败。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import compile as C  # noqa: E402  诊断类型、字体写入、颜色与几何工具共用一份实现
import inspect_pptx as I  # noqa: E402  保留核对复用只读解析器

qn = C.qn

OPS = ("text", "box", "style", "fill", "line", "order", "image", "chart_data", "delete", "background")
# 这些字段一旦出现在操作里，就说明它有意改样式；只改字时一个都不该有。
STYLE_KEYS = ("font", "font_ea", "font_size", "bold", "color", "letter_spacing_pt",
              "align", "line_spacing", "pad_in", "valign")
EMU = C.EMU

# 操作级属性白名单：每个操作只接受表中的键，拼错或不适用的键直接拒绝（不静默忽略）。
COMMON_KEYS = {"op", "page", "target", "target_regex", "target_index", "note"}
OP_KEYS = {
    "text": {"content", "runs", "space_after_pt"} | (set(STYLE_KEYS) - {"pad_in", "valign"}),
    "box": {"x", "y", "w", "h"},
    "style": set(STYLE_KEYS),
    "fill": {"color", "alpha"},
    "line": {"color", "width_pt", "dash", "hide"},
    "order": {"to"},
    "image": {"image_id"},
    "chart_data": {"data"},
    "delete": set(),
    "background": {"color"},
}
# 每种操作授权改动的对象属性（对象指纹的顶层键）。段落样式与文本框字段再按声明收窄。
GRANT_TOP = {
    "text": {"text"}, "box": {"box_in"}, "style": set(), "fill": {"fill"},
    "line": {"stroke", "stroke_pt"}, "image": {"sha256", "px", "ppi", "crop"},
    "chart_data": {"plots", "payload"}, "order": set(), "delete": set(), "background": set(),
}
STYLE_SIG = {"font": "font", "font_ea": "font_ea", "font_size": "font_size", "bold": "bold",
             "color": "color", "align": "align", "line_spacing": "line_spacing",
             "letter_spacing_pt": "letter_spacing_pt"}
STYLE_FRAME = {"pad_in": "ins", "valign": "anchor"}


# ---------------------------------------------------------------- 目标解析


def slide_at(prs, page: Any) -> Any:
    if not isinstance(page, int) or isinstance(page, bool) or not (1 <= page <= len(prs.slides)):
        raise C.BadOp(f"page 必须是 1..{len(prs.slides)} 的整数，收到 {page!r}")
    return prs.slides[page - 1]


def resolve(shape_list, op: Dict[str, Any]) -> Any:
    """按名称、正则或 z 序定位对象。名称多处重名时必须改用 target_index，不替人猜。"""
    name = op.get("target")
    regex = op.get("target_regex")
    index = op.get("target_index")
    if index is not None:
        if not isinstance(index, int) or not (0 <= index < len(shape_list)):
            raise C.BadOp(f"target_index 必须是 0..{len(shape_list) - 1} 的整数，收到 {index!r}")
        return shape_list[index]
    if regex is not None:
        hits = [s for s in shape_list if re.search(str(regex), s.name)]
    elif name is not None:
        hits = [s for s in shape_list if s.name == str(name)]
    else:
        raise C.BadOp("缺少 target / target_regex / target_index")
    if not hits:
        listing = ", ".join(f"{i}:{s.name}" for i, s in enumerate(shape_list))[:400]
        raise C.BadOp(f"找不到目标 {name or regex!r}；本页对象为 {listing}")
    if len(hits) > 1:
        raise C.BadOp(f"目标 {name or regex!r} 命中 {len(hits)} 个对象，改用 target_index："
                      f"{[s.name for s in hits][:6]}")
    return hits[0]


# ---------------------------------------------------------------- 文本


def run_style(run) -> Dict[str, Any]:
    """一个 run 的**显式**样式；没声明的字段是 None。

    None 表示"继承自版式或主题"，不能替它编一个值——真实稿的标题多半是 44pt 主题样式、
    run 上什么都没有，编成 18pt 等于把设计改坏，而保留核对不会响。
    """
    from pptx.enum.dml import MSO_COLOR_TYPE

    color = None
    try:
        if run.font.color and run.font.color.type == MSO_COLOR_TYPE.RGB:
            color = f"#{run.font.color.rgb}"
    except (AttributeError, ValueError):
        color = None
    rPr = run._r.find(qn("a:rPr"))
    ea = None
    if rPr is not None and rPr.find(qn("a:ea")) is not None:
        ea = rPr.find(qn("a:ea")).get("typeface")
    spc = rPr.get("spc") if rPr is not None else None
    return {"font": run.font.name, "font_ea": ea,
            "size": run.font.size.pt if run.font.size else None,
            "bold": run.font.bold, "color": color,
            "spacing": int(spc) / 100.0 if spc else None}


def paragraph_styles(shape) -> List[Dict[str, Any]]:
    """逐段：该段的 run 样式列表，加上段落级对齐与行距。

    只取第一段的第一个 run 是"改一处、坏一处"的根源：多段标题、副题、灰色眉题
    各有各的字级，用第一个 run 的样式重写全部，等于把整块文字压成一种。
    """
    out: List[Dict[str, Any]] = []
    for para in shape.text_frame.paragraphs:
        out.append({
            "runs": [run_style(run) for run in para.runs],
            "align": str(para.alignment).split(" ")[0].lower() if para.alignment is not None else None,
            "line_spacing": para.line_spacing,
        })
    return out


def base_style(shape) -> Dict[str, Any]:
    """兼容旧调用：第一段第一个 run 的样式与段落属性。"""
    paras = paragraph_styles(shape)
    if not paras:
        return {}
    first = paras[0]["runs"][0] if paras[0]["runs"] else {}
    return {**first, "align": paras[0]["align"], "line_spacing": paras[0]["line_spacing"]}


def _first(*values: Any) -> Any:
    """第一个非 None 的值。False 与 0 都是合法值，不能用 or 链代替。"""
    for value in values:
        if value is not None:
            return value
    return None


def _write_run(run, piece: Dict[str, Any], op: Dict[str, Any], base: Dict[str, Any]) -> None:
    """run 级 > 操作级 > 原稿该 run 的显式值；三层都没有就不写，让继承继续生效。"""
    C.set_run_attrs(
        run,
        font=_first(piece.get("font"), op.get("font"), base.get("font")),
        font_ea=_first(piece.get("font_ea"), op.get("font_ea"), base.get("font_ea")),
        size=_first(piece.get("font_size"), op.get("font_size"), base.get("size")),
        color=_first(piece.get("color"), op.get("color"), base.get("color")),
        bold=_first(piece.get("bold"), op.get("bold"), base.get("bold")),
        spacing=_first(piece.get("letter_spacing_pt"), op.get("letter_spacing_pt"), base.get("spacing")),
    )


def _mixed(runs: List[Dict[str, Any]]) -> bool:
    """同一段里是否存在两种以上不同的 run 样式（混排）。"""
    sigs = {(r.get("size"), r.get("bold"), r.get("font"), r.get("font_ea"), r.get("color")) for r in runs}
    return len(sigs) > 1


def write_text(shape, op: Dict[str, Any], rep: C.Report, page: int, name: str) -> None:
    """换字，同时守住原有排版：逐段对齐/行距、逐 run 字号字重字色字族能继承就继承。

    保真规则：新段落 i 的样式基线取原稿第 i 段；若该段新 run 数与原 run 数一致，
    就按位置一一继承（深色主句 + 灰色括注这种混排不会被压平）。数量对不上时整段
    退回该段首个 run 的样式，并在报告里披露——混排被统一了，读者该知道。
    """
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Pt

    tf = shape.text_frame
    old = paragraph_styles(shape)
    content = str(op.get("content", ""))
    runs = op.get("runs")
    if runs and "".join(str(r.get("text", "")) for r in runs) != content:
        raise C.BadOp("runs 拼接与 content 不一致（含 \\n）")
    paras = C.paragraph_runs({"content": content, "runs": runs})
    if paras is None:
        raise C.BadOp("runs 分段与 content 对不上")

    fallback = old[0]["runs"][0] if old and old[0]["runs"] else {}
    collapsed: List[int] = []

    for extra in list(tf.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)
    first = tf.paragraphs[0]
    for run in list(first.runs):
        run._r.getparent().remove(run._r)

    for idx, (seg, pieces) in enumerate(paras):
        para = first if idx == 0 else tf.add_paragraph()
        src = old[idx] if idx < len(old) else old[0] if old else {"runs": [], "align": None, "line_spacing": None}
        align = op.get("align") or src.get("align")
        if align:
            para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                              "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}[align]
        spacing = float(op["line_spacing"]) if op.get("line_spacing") else src.get("line_spacing")
        if spacing:
            para.line_spacing = spacing
        if op.get("space_after_pt"):
            para.space_after = Pt(float(op["space_after_pt"]))

        src_runs = src.get("runs") or []
        if src_runs and _mixed(src_runs) and len(pieces or [{"text": seg}]) != len(src_runs):
            collapsed.append(idx + 1)
        for j, piece in enumerate(pieces or ([{"text": seg}] if seg else [])):
            base = src_runs[j] if len(src_runs) == len(pieces or []) and j < len(src_runs) else \
                (src_runs[0] if src_runs else fallback)
            run = para.add_run()
            run.text = piece.get("text", "")
            _write_run(run, piece, op, base)

    if collapsed:
        rep.soft("W-RUNS", f"第 {page} 页 {name}",
                 f"第 {collapsed} 段原有混排（同一段内字号/字色/字重不同），新内容 run 数与之不同，"
                 f"已统一为该段首个 run 的样式；要保留分区样式请用 runs 明确写出")


def apply_style(shape, op: Dict[str, Any]) -> None:
    """只改声明的属性。一个 color 操作不该顺手把字号、字重、字距也重写一遍。"""
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.util import Pt

    tf = shape.text_frame
    for para in tf.paragraphs:
        if op.get("align"):
            para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                              "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}[op["align"]]
        if op.get("line_spacing"):
            para.line_spacing = float(op["line_spacing"])
        for run in para.runs:
            # 显式声明 font 就是要换字族：连东亚字族一起写，否则中文仍留在原字族上。
            C.set_run_attrs(run, font=op.get("font"), font_ea=op.get("font_ea") or op.get("font"),
                            size=op.get("font_size"), color=op.get("color"),
                            bold=op.get("bold"), spacing=op.get("letter_spacing_pt"))
    if op.get("valign"):
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                              "bottom": MSO_ANCHOR.BOTTOM}[op["valign"]]
    if op.get("pad_in") is not None:
        pad = float(op["pad_in"])
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(pad * 72)


# ---------------------------------------------------------------- 几何与外观


def apply_box(shape, op: Dict[str, Any]) -> None:
    from pptx.util import Inches
    for key, attr in (("x", "left"), ("y", "top"), ("w", "width"), ("h", "height")):
        if op.get(key) is not None:
            setattr(shape, attr, Inches(float(op[key])))


def apply_fill(shape, op: Dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor
    from lxml import etree

    color = op.get("color")
    if color is None:
        raise C.BadOp("fill 需要 color")
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(str(color).lstrip("#"))
    if op.get("alpha") is not None:
        alpha = max(0.0, min(1.0, float(op["alpha"])))
        srgb = shape.fill.fore_color._xFill.find(qn("a:srgbClr"))
        if srgb is not None:
            for old in srgb.findall(qn("a:alpha")):
                srgb.remove(old)
            etree.SubElement(srgb, qn("a:alpha")).set("val", str(int(round(alpha * 100000))))


def apply_line(shape, op: Dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    if op.get("color"):
        shape.line.color.rgb = RGBColor.from_string(str(op["color"]).lstrip("#"))
    if op.get("width_pt") is not None:
        shape.line.width = Pt(float(op["width_pt"]))
    if op.get("dash"):
        from pptx.enum.dml import MSO_LINE_DASH_STYLE
        shape.line.dash_style = {"solid": MSO_LINE_DASH_STYLE.SOLID, "dash": MSO_LINE_DASH_STYLE.DASH,
                                 "dot": MSO_LINE_DASH_STYLE.ROUND_DOT, "dash_dot": MSO_LINE_DASH_STYLE.DASH_DOT,
                                 "long_dash": MSO_LINE_DASH_STYLE.LONG_DASH}[str(op["dash"])]
    if op.get("hide"):
        shape.line.fill.background()


def apply_order(shape, op: Dict[str, Any]) -> None:
    """层级也是版面：把色块压到底、把标签提到顶，是编辑里最常见的一步。"""
    to = op.get("to")
    node = shape._element
    parent = node.getparent()
    if to == "back":
        parent.remove(node)
        parent.insert(2, node)  # spTree: nvGrpSpPr, grpSpPr, 之后才是第一个对象
    elif to == "front":
        parent.remove(node)
        parent.append(node)
    else:
        raise C.BadOp("order 的 to 取 back 或 front")


def apply_image(shape, path: Path) -> None:

    if not path.is_file():
        raise C.BadOp(f"素材不存在：{path}")
    image_part, rId = shape.part.get_or_add_image_part(str(path))
    blip = shape._element.blipFill.find(qn("a:blip"))
    if blip is None:
        raise C.BadOp("目标不是图片对象")
    blip.set(qn("r:embed"), rId)


def apply_chart(shape, op: Dict[str, Any]) -> None:
    from pptx.chart.data import CategoryChartData

    if not getattr(shape, "has_chart", False):
        raise C.BadOp("目标不是图表对象")
    chart = shape.chart
    if len(chart.plots) > 1:
        raise C.BadOp("该图有多个图层（组合图）：改数据会同时打乱两层格式，请整稿重建")
    data = op.get("data") or {}
    cats, series = data.get("categories") or [], data.get("series") or []
    if not cats or not series:
        raise C.BadOp("chart_data 需要 data.categories 与 data.series")
    existing = [s.name for s in chart.plots[0].series]
    if len(series) != len(existing):
        raise C.BadOp(f"系列数 {len(series)} ≠ 原图 {len(existing)}；加减系列会破坏配色与格式，请重建")
    for i, s in enumerate(series):
        if len(s.get("values") or []) != len(cats):
            raise C.BadOp(f"系列「{s.get('name')}」值数 {len(s.get('values') or [])} ≠ 类别数 {len(cats)}")
        if s.get("name") and s["name"] != existing[i]:
            raise C.BadOp(f"第 {i + 1} 个系列名 {s['name']!r} ≠ 原图 {existing[i]!r}；改名会丢图例格式")
    payload = CategoryChartData()
    payload.categories = cats
    for s in series:
        payload.add_series(s["name"], [float(v) for v in s["values"]])
    chart.replace_data(payload)


def apply_background(slide, op: Dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor

    color = op.get("color")
    if color is None:
        raise C.BadOp("background 需要 color")
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(str(color).lstrip("#"))


# ---------------------------------------------------------------- 主流程


def _chart_parts(shape) -> set:
    """图表改数会重写自己的 chart 部件与内嵌工作簿：这两个部件因此可以变。

    只放行"这个图表自己的"部件，别的图表被顺手改写照样会报——放行整片
    `ppt/charts/` 就等于关掉了这条检查。
    """
    out: set = set()
    try:
        part = shape.chart.part
    except (AttributeError, ValueError):
        return out
    out.add(str(part.partname).lstrip("/"))
    for rel in getattr(part, "rels", {}).values():
        target = getattr(rel, "target_part", None)
        name = str(getattr(target, "partname", "") or "").lstrip("/")
        if name:
            out.add(name)
    return out


def check_keys(kind: Any, op: Dict[str, Any], where: str) -> None:
    """操作级白名单：未知操作、拼错或不适用的键一律拒绝。静默忽略的结果是报告写着"已改"、文件里什么都没变。"""
    if kind not in OP_KEYS:
        raise C.BadOp(f"{where}: 未知操作 {kind!r}；可用 {', '.join(OPS)}")
    extra = sorted(set(op) - COMMON_KEYS - OP_KEYS[kind])
    if extra:
        raise C.BadOp(f"{where}: {kind} 不接受 {extra}；该操作只认 {sorted(OP_KEYS[kind])}")


def empty_grant() -> Dict[str, Any]:
    return {"top": set(), "sig": set(), "frame": set(), "text": False}


def grant_for(kind: str, op: Dict[str, Any]) -> Dict[str, Any]:
    """一次操作授权改动的属性：顶层键、声明过的段落样式字段、声明过的文本框字段。"""
    declared = set(op) | {k for run in op.get("runs") or [] for k in run}
    grant = empty_grant()
    grant["top"] = set(GRANT_TOP[kind])
    grant["sig"] = {STYLE_SIG[k] for k in declared if k in STYLE_SIG}
    grant["frame"] = {STYLE_FRAME[k] for k in declared if k in STYLE_FRAME}
    grant["text"] = kind == "text"
    return grant


def merge_grant(into: Dict[str, Any], extra: Dict[str, Any]) -> None:
    for key in ("top", "sig", "frame"):
        into[key] |= extra[key]
    into["text"] = into["text"] or extra["text"]


def run_ops(prs, ops: List[Dict[str, Any]], asset_root: Path, log: List[Dict[str, Any]],
            rep: C.Report) -> Tuple[Dict[Tuple[int, int], Dict[str, Any]], set, set, set]:
    """逐条执行操作。对象以 (页, shape id) 为键；返回每个被点名对象的授权清单。"""
    grants: Dict[Tuple[int, int], Dict[str, Any]] = {}
    bg_touched: set = set()
    deleted: set = set()
    movable: set = set()
    for n, op in enumerate(ops, 1):
        where = f"op{n}"
        kind = op.get("op")
        check_keys(kind, op, where)
        slide = slide_at(prs, op.get("page"))
        if kind == "background":
            apply_background(slide, op)
            bg_touched.add(int(op["page"]))
            log.append({"op": n, "kind": kind, "page": op["page"]})
            continue
        shape = resolve(slide.shapes, op)
        key = (int(op["page"]), shape.shape_id)
        merge_grant(grants.setdefault(key, empty_grant()), grant_for(kind, op))
        if kind == "text":
            if op.get("content") is None:
                raise C.BadOp(f"{where}: text 需要 content")
            write_text(shape, op, rep, int(op["page"]), shape.name)
        elif kind == "box":
            if not any(op.get(k) is not None for k in ("x", "y", "w", "h")):
                raise C.BadOp(f"{where}: box 至少给 x/y/w/h 之一")
            apply_box(shape, op)
        elif kind == "style":
            if not getattr(shape, "has_text_frame", False):
                raise C.BadOp(f"{where}: {shape.name} 没有文本可改样式")
            apply_style(shape, op)
        elif kind == "fill":
            apply_fill(shape, op)
        elif kind == "line":
            apply_line(shape, op)
        elif kind == "image":
            if not op.get("image_id"):
                raise C.BadOp(f"{where}: image 需要 image_id")
            path = C.resolve_asset(asset_root, str(op["image_id"]))
            if path is None:
                raise C.BadOp(f"{where}: 找不到素材 {op['image_id']!r}（素材根 {asset_root}）")
            apply_image(shape, path)
        elif kind == "chart_data":
            apply_chart(shape, op)
            movable |= _chart_parts(shape)
        elif kind == "order":
            apply_order(shape, op)
        elif kind == "delete":
            shape._element.getparent().remove(shape._element)
            deleted.add(key)
        log.append({"op": n, "kind": kind, "page": op["page"], "target": shape.name,
                    "note": op.get("note", "")})
    return grants, bg_touched, deleted, movable


def ink_box(obj: Dict[str, Any]) -> Optional[Tuple[float, float, float, float]]:
    """文字实际落点的大致范围：声明框常比字宽，用声明框判遮挡会天天误报。

    估法与本包其他处一致（中文一字一 em、西文约 0.55 em），宽不超过声明框。
    """
    box = obj.get("box_in")
    if not box or None in box:
        return None
    if obj.get("kind") != "text":
        return tuple(box)  # type: ignore[return-value]
    paras = (obj.get("styles") or {}).get("paragraphs") or []
    if not paras:
        return tuple(box)  # type: ignore[return-value]
    widest, tallest = 0.0, 0.0
    for para in paras:
        size = max([r.get("size_pt") or 0 for r in para.get("runs") or []] or [0])
        if not size:
            continue
        text = para.get("text") or ""
        latin = sum(1 for ch in text if ord(ch) < 128)
        cjk = max(len(text) - latin, 0)
        widest = max(widest, (cjk + latin * 0.55) * size / 72)
        tallest += size * 1.35 / 72
    if not widest:
        return tuple(box)  # type: ignore[return-value]
    return (box[0], box[1], min(box[2], max(widest, 0.2)), min(box[3], max(tallest, 0.1)))


def post_check(after: Dict[str, Any], f: C.Report) -> None:
    """改完再按同一套版面规则过一遍：挪过的文字有没有被上层色块盖住、有没有出画布。

    新建稿由 compile 的 E-OVERLAP 守；改动稿不走 plan，所以这里补同一道闸。
    """
    w_in, h_in = after.get("canvas_in") or (13.333, 7.5)
    for page in after["pages"]:
        objects = page["objects"]
        boxes = [(o.get("name"), o.get("kind"), o.get("box_in"), o.get("fill"), o) for o in objects]
        for i, (name, kind, box, _, obj) in enumerate(boxes):
            if not box or None in box:
                continue
            if box[0] < -0.02 or box[1] < -0.02 or box[0] + box[2] > w_in + 0.02 or box[1] + box[3] > h_in + 0.02:
                if kind in ("text", "chart", "table"):
                    f.soft("W-BOUND", f"第 {page['index']} 页 {name}", "改动后超出画布，会被裁掉")
                continue
            if kind != "text":
                continue
            glyph = ink_box(obj) or box
            for up_name, up_kind, up_box, up_fill, _ in boxes[i + 1:]:
                if not up_box or None in up_box:
                    continue
                opaque = bool(up_fill) or up_kind in ("picture", "chart", "table")
                if not opaque:
                    continue
                ix, iy = C._intersection(glyph, up_box)
                area = ix * iy
                if area >= 0.35 and area >= 0.3 * max(glyph[2] * glyph[3], 1e-6):
                    f.soft("W-COVER", f"第 {page['index']} 页 {name}",
                           f"文字落在上层的 {up_name} 上（约 {area:.1f} in²）；挪位置或调层级")
                    break


_SIG_FIELDS = ("font_size", "bold", "font", "font_ea", "color", "align", "line_spacing", "letter_spacing_pt")


def _text_sig(obj: Dict[str, Any]) -> Optional[List[set]]:
    """逐段样式签名：每段用到的（字号、字重、西文/东亚字族、墨色、对齐、行距）集合。

    只看第一个 run 会放过"第二段被压成 32pt"这种最常见的漂移——副题、眉题、括注
    在真实稿里各有各的字级。按段比对而不是按 run 位置比对：换行与增删 run 是内容
    编辑，压平整段样式才是改设计。段落级属性也进签名：只改字的操作顺手改掉对齐，
    是同一类漂移，只是发生在上层。
    """
    paras = (obj.get("styles") or {}).get("paragraphs") or []
    out: List[set] = []
    for para in paras:
        align = para.get("align")
        spacing = para.get("line_spacing")
        out.append({(r.get("size_pt"), r.get("bold"), r.get("font"), r.get("ea"), r.get("color"),
                     align, spacing, r.get("spacing")) for r in para.get("runs") or []})
    return out or None


def _mask(sigs: set, allowed: frozenset) -> set:
    """把"本操作合法声明过"的字段抹平，只留没被授权的位。"""
    return {tuple(None if _SIG_FIELDS[i] in allowed else v for i, v in enumerate(s)) for s in sigs}


def _style_drift(old: Optional[List[set]], new: Optional[List[set]], allowed: frozenset = frozenset()):
    """只改字的操作不许给任何一段换一套样式：新段落用到的样式必须是它原来就有的。

    允许"少用"（内容删掉了那句灰色括注），不允许"新出现"（14pt 副题被压成 32pt，
    或继承的 44pt 被写死成 18pt）。同一目标上另有 `style` 操作显式声明过的字段不参与
    判定——否则"先换字、再改字号"这种正常写法会被自己的闸拦住。
    """
    if not old or not new:
        return None
    olds = [_mask(s, allowed) for s in old]
    news = [_mask(s, allowed) for s in new]

    def ref_for(index: int) -> Optional[set]:
        # 段落数变多时用相邻段作参照：新增的那一段没有"原来的样式"可比。
        for j in range(min(index, len(olds) - 1), -1, -1):
            if olds[j]:
                return olds[j]
        for j in range(index + 1, len(olds)):
            if olds[j]:
                return olds[j]
        return None

    for i, sigs in enumerate(news):
        if not sigs:
            continue
        ref = ref_for(i)
        if not ref:
            continue
        stray = sigs - ref
        if stray:
            return i + 1, sorted(ref, key=str), sorted(stray, key=str)
    return None


def _short(value: Any, n: int = 80) -> str:
    text = repr(value)
    return text if len(text) <= n else text[:n] + "…"


def _object_drift(obj: Dict[str, Any], twin: Dict[str, Any], grant: Optional[Dict[str, Any]],
                  where: str, f: C.Report) -> None:
    """对象的改动是否都落在授权之内。未被点名的对象（grant 为 None）必须逐项一致。"""
    top = grant["top"] if grant else set()
    for key in sorted((set(obj) | set(twin)) - {"styles", "frame"}):
        if key not in top and obj.get(key) != twin.get(key):
            f.hard("E-PRESERVE", where,
                   f"未授权的属性「{key}」被改动：{_short(obj.get(key))} → {_short(twin.get(key))}")
    granted_frame = grant["frame"] if grant else set()
    fb, fa = obj.get("frame") or {}, twin.get("frame") or {}
    bad = sorted(k for k in set(fb) | set(fa) if k not in granted_frame and fb.get(k) != fa.get(k))
    if bad:
        f.hard("E-PRESERVE", where, f"未授权的文本框属性被改动：{bad}")
    allowed = frozenset(grant["sig"]) if grant else frozenset()
    sig_b, sig_a = _text_sig(obj), _text_sig(twin)
    if grant and grant["text"]:
        drift = _style_drift(sig_b, sig_a, allowed)
        if drift:
            para_no, ref, stray = drift
            f.hard("E-PRESERVE", where,
                   f"只改文字的操作用改了样式：第 {para_no} 段出现原稿没有的"
                   f"（字号, 粗体, 西文字族, 东亚字族, 墨色, 对齐, 行距, 字距）={stray}（该段原有 {ref}）")
    elif [_mask(s, allowed) for s in sig_b or []] != [_mask(s, allowed) for s in sig_a or []]:
        f.hard("E-PRESERVE", where, "未授权的段落样式被改动（字号/字重/字族/墨色/对齐/行距/字距）")


FLAG_LABEL = {"hidden": "隐藏", "transition": "切换", "animation": "动画", "smartart": "SmartArt",
              "ole": "嵌入对象", "media": "音视频", "notes": "备注"}


def compare(before: Dict[str, Any], after: Dict[str, Any], grants: Dict[Tuple[int, int], Dict[str, Any]],
            bg_touched: set, deleted: set, f: C.Report) -> Dict[str, Any]:
    """保留核对：没有任何操作授权改动的东西，必须与原稿逐项一致。这是"没有顺手改坏别处"的证据。"""
    for key, label in (("slides", "页数"), ("masters", "母版数"), ("layouts", "版式数"), ("macros", "宏")):
        if before.get(key) != after.get(key):
            f.hard("E-PRESERVE", label, f"{before.get(key)} → {after.get(key)}")
    if before.get("theme_fonts") != after.get("theme_fonts"):
        f.hard("E-PRESERVE", "主题字体", "改动前后不一致")

    removed = 0
    for i, b_page in enumerate(before["pages"]):
        where_page = f"第 {i + 1} 页"
        if i >= len(after["pages"]):
            f.hard("E-PRESERVE", where_page, "页面丢失")
            continue
        a_page = after["pages"][i]
        # 逐页特性（隐藏、切换、动画、SmartArt、嵌入、音视频、备注）没有任何操作授权改它们
        for flag in sorted(set(b_page["flags"]) | set(a_page["flags"])):
            bv, av = b_page["flags"].get(flag), a_page["flags"].get(flag)
            if bv != av:
                f.hard("E-PRESERVE", where_page, f"页面特性「{FLAG_LABEL.get(flag, flag)}」被改动：{bv} → {av}")
        if b_page.get("notes") != a_page.get("notes"):
            f.hard("E-PRESERVE", where_page, "备注文字被改动")
        if b_page.get("layout") != a_page.get("layout"):
            f.hard("E-PRESERVE", where_page, f"版式被改动：{b_page.get('layout')} → {a_page.get('layout')}")
        if i + 1 not in bg_touched and b_page.get("background") != a_page.get("background"):
            f.hard("E-PRESERVE", where_page,
                   f"未点名的底色被改动：{b_page.get('background')} → {a_page.get('background')}")

        b_ids = [o.get("id") for o in b_page["objects"]]
        a_ids = [o.get("id") for o in a_page["objects"]]
        if len(set(b_ids)) != len(b_ids) or len(set(a_ids)) != len(a_ids):
            f.hard("E-PRESERVE", where_page, "对象标识重复，无法逐对象核对；请先在 PowerPoint 中修复重复标识")
            continue
        b_objs = dict(zip(b_ids, b_page["objects"]))
        a_objs = dict(zip(a_ids, a_page["objects"]))
        for oid, obj in b_objs.items():
            where = f"{where_page} {obj.get('name')}"
            if oid not in a_objs:
                removed += 1
                if (i + 1, oid) not in deleted:
                    f.hard("E-PRESERVE", where, "对象消失，但没有任何 delete 操作点名它")
                continue
            _object_drift(obj, a_objs[oid], grants.get((i + 1, oid)), where, f)
        for oid, obj in a_objs.items():
            if oid not in b_objs:
                f.hard("E-PRESERVE", f"{where_page} {obj.get('name')}", "出现原稿没有的对象；定点编辑不新增任何对象")
        # 没被点名的对象，相对前后次序必须不变（order 操作只授权它移动的那一个）
        kept_b = [oid for oid in b_ids if (i + 1, oid) not in grants and oid in a_objs]
        kept_a = [oid for oid in a_ids if (i + 1, oid) not in grants and oid in b_objs]
        if kept_b != kept_a:
            f.hard("E-PRESERVE", where_page, "未点名对象的层级（前后次序）被改动")
    return {"deleted": removed, "touched": len(grants)}


def compare_parts(before: Dict[str, str], after: Dict[str, str], before_slides: List[str],
                  after_slides: List[str], touched_pages: set, f: C.Report,
                  movable: Optional[set] = None) -> Dict[str, Any]:
    """包内部件比对。slide 按页比对（页序不等于部件名，python-pptx 保存时还会重新编号）；
    其余部件按名字逐字节比对。被点名页的 slide 与其 rels、改数图表自己的部件可以变，别的不许变。"""
    import re as _re

    allowed_new = ("ppt/media/", "ppt/embeddings/")
    movable = movable or set()
    slide_re = _re.compile(r"^ppt/slides/(?:_rels/)?slide\d+\.xml(?:\.rels)?$")

    def page_bundle(parts: Dict[str, str], slide_name: str) -> Tuple[Any, Any]:
        rels = slide_name.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        return parts.get(slide_name), parts.get(rels)

    changed, lost, added, moved = [], [], [], []
    for i in range(min(len(before_slides), len(after_slides))):
        if i + 1 in touched_pages:
            continue
        if page_bundle(before, before_slides[i]) != page_bundle(after, after_slides[i]):
            changed.append(f"第 {i + 1} 页的 slide 部件")
    for name, digest in before.items():
        if slide_re.match(name) or name == "[Content_Types].xml":
            continue
        if name not in after:
            lost.append(name)
        elif after[name] != digest:
            (moved if name in movable else changed).append(name)
    for name in after:
        if name in before or slide_re.match(name) or name.startswith(allowed_new):
            continue
        added.append(name)

    if lost:
        f.hard("E-PRESERVE", "包内部件", f"部件丢失：{lost[:6]}")
    if changed:
        f.hard("E-PRESERVE", "包内部件", f"未点名的部件被改写：{changed[:6]}")
    if added:
        f.hard("E-PRESERVE", "包内部件", f"出现原稿没有的部件：{added[:6]}")
    return {"parts": len(after), "changed": len(changed), "lost": len(lost), "added": len(added),
            "moved": moved}


# ---------------------------------------------------------------- 主流程


def main(argv: Optional[List[str]] = None) -> int:
    C.utf8_stdio()
    parser = argparse.ArgumentParser(description="在既有 PPTX 上做定点修改")
    parser.add_argument("source")
    parser.add_argument("ops")
    parser.add_argument("out")
    parser.add_argument("--assets", default=None, help="素材目录；缺省为 ops 同级 assets/ 或所在目录")
    parser.add_argument("--report", default=None)
    parser.add_argument("--force", action="store_true", help="允许替换已存在的输出文件")
    args = parser.parse_args(argv)

    src = Path(args.source).resolve()
    ops_path = Path(args.ops).resolve()
    out = Path(args.out).resolve()
    report_path = Path(args.report).resolve() if args.report else None
    asset_root = Path(args.assets).resolve() if args.assets else (
        ops_path.parent / "assets" if (ops_path.parent / "assets").is_dir() else ops_path.parent)

    rep = C.Report()
    if not src.is_file():
        rep.hard("E-SRC", str(src), "找不到原稿")
    for label, target in (("输出文件", out), ("--report", report_path)):
        if target is not None and target in (src, ops_path):
            rep.hard("E-EXIST", str(target), f"{label}与输入同路径，会覆盖源文件")
    if report_path is not None and report_path == out:
        rep.hard("E-EXIST", str(out), "--report 与输出文件同路径")
    if out.suffix.lower() not in (".pptx", ".pptm"):
        rep.hard("E-EXIST", str(out), f"输出文件名应以 .pptx 结尾，收到 {out.name!r}")
    if out.exists() and not args.force:
        rep.hard("E-EXIST", str(out), "目标文件已存在，未写入；确认要替换就加 --force")
    if rep.blocked:
        for issue in rep.items:
            print(issue)
        print(f"\n阻断：{rep.count('HARD')} 个硬错误，未写出文件。")
        return 2

    from pptx import Presentation
    try:
        prs = Presentation(str(src))
        spec = json.loads(ops_path.read_text(encoding="utf-8-sig"))
        ops = spec.get("ops") or []
        if not ops:
            raise C.BadOp("edit.json 里没有 ops")
    except (OSError, json.JSONDecodeError, C.BadOp) as exc:
        print(f"[HARD] E-OPS 无法读取：{type(exc).__name__}: {exc}")
        return 2

    before = I.inspect(src)
    before_parts = I.structure(src)
    before_slides = I.slide_part_names(src)
    log: List[Dict[str, Any]] = []
    try:
        grants, bg_touched, declared_deleted, movable = run_ops(prs, ops, asset_root, log, rep)
    except C.BadOp as exc:
        print(f"[HARD] E-OPS 未写入：{exc}")
        return 2
    except Exception as exc:  # 运行期异常：不留半成品
        print(f"[HARD] E-RUNTIME {type(exc).__name__}: {exc}")
        return 3

    # 先落到临时件：保留核对不通过时，交付路径一个字节都不动。
    tmp = out.with_name(f".{out.name}.{os.getpid()}.part")
    try:
        prs.save(str(tmp))
        after = I.inspect(tmp)
        after_parts = I.structure(tmp)
        after_slides = I.slide_part_names(tmp)
    except Exception as exc:
        if tmp.exists():
            tmp.unlink()
        print(f"[HARD] E-RUNTIME {type(exc).__name__}: {exc}")
        return 3

    for label, rec in (("原稿", before), ("改动后", after)):
        for err in rec.get("errors") or []:
            rep.hard("E-PRESERVE", label, f"无法完整解析，逐对象核对不可靠：{err}")
    touched_pages = {int(op["page"]) for op in ops}
    facts = compare(before, after, grants, bg_touched, declared_deleted, rep)
    facts["parts"] = compare_parts(before_parts, after_parts, before_slides, after_slides,
                                   touched_pages, rep, movable)
    if not rep.blocked:
        post_check(after, rep)
        os.replace(tmp, out)
    elif tmp.exists():
        # 保留核对未通过：把证据留在旁边，交付路径不写。
        rejected = out.with_name(out.stem + ".preserve-failed.pptx")
        os.replace(tmp, rejected)
        facts["rejected"] = str(rejected)
    status = "BLOCK" if rep.blocked else "READY"

    print(f"\n已改写 {out}（{len(ops)} 个操作，{len(after['pages'])} 页）" if not rep.blocked
          else f"\n保留核对未通过：{out} 未写入；诊断稿 {facts.get('rejected')}（不是交付物）")
    for item in log:
        print(f"  {item['op']:>2}. {item['kind']:<10} 第 {item['page']} 页"
              f"{' · ' + item['target'] if item.get('target') else ''}")
    for issue in rep.items:
        print(issue)
    print(f"保留核对：{status}   问题 {rep.count('HARD')}")
    print(f"页数 {len(after['pages'])}｜母版 {after['masters']}｜版式 {after['layouts']}｜宏 {after['macros']}"
          f"｜包内 {facts['parts']['parts']} 部件（越界改动 {facts['parts']['changed']}／丢失 {facts['parts']['lost']}"
          f"／声明过的改写 {len(facts['parts']['moved'])}）")
    if report_path:
        report_path.write_text(json.dumps({
            "status": status,
            "source": str(src), "output": str(out),
            "sha256": hashlib.sha256(out.read_bytes()).hexdigest() if out.is_file() else None,
            "ops": log, "preservation": facts,
            "issues": rep.to_dict(),
        }, ensure_ascii=False, indent=2), encoding="utf-8")
    if not rep.blocked and not src.samefile(out):
        print("原稿未被改动；母版、动画、切换、SmartArt、备注按原样保留。")
    print("机器能证明的到此为止：改动后的版面是否更好看，仍需自己看渲染。")
    return 0 if not rep.blocked else 1


if __name__ == "__main__":
    sys.exit(main())
