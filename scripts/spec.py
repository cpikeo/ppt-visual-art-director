#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plan → render spec（L2→L3 的唯一交接物）。

所有设计计算（文字样式、几何、裁切窗、图表配色/直标/轴范围、图例布局、页码、主题色）只在这里算一次，
写成“已解析”的 JSON 交给 scripts/render.js 去画。JS 不重复实现任何设计逻辑；
JS 画不出的 OOXML 能力（渐变、蒙版、双色调、软边、分组、图表直标…）以 `ooxml` 指令随对象带到 post.py。
"""

from __future__ import annotations

import math
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional

from PIL import Image

import motion
from model import (CANVAS_EMU, EMU, LIMITS, MASK_PRESETS, TEXT_TYPES, alt_text, chart_color_map, chart_flags, chart_items,
                   chart_label_plan, chart_min_value, chart_mode, chart_non_negative, chart_style, chart_surface,
                   chart_unit, contrast_ratio, corner_adj, crop_geometry, deck_palette, explicit_box, is_dark,
                   legend_label, list_indent_for, mix_toward, page_bg, page_numbers, paragraph_runs, pn_box, pn_ink,
                   resolve_asset, resolve_points, text_style)

PRST = {"rect": "rect", "round_rect": "roundRect", "oval": "ellipse", "triangle": "triangle", "diamond": "diamond",
        "chevron": "chevron", "arrow_right": "rightArrow", "arrow_up": "upArrow"}
DASH_JS = {"solid": "solid", "dash": "dash", "dot": "sysDot", "dash_dot": "dashDot", "long_dash": "lgDash"}
SPEC_VERSION = 1


def _hex(c: Any) -> str:
    return str(c).lstrip("#").upper()


def _stroke(stroke: Optional[Dict[str, Any]], arrow: str = "none") -> Optional[Dict[str, Any]]:
    if not stroke:
        return None
    return {"color": _hex(stroke["color"]), "width_pt": float(stroke.get("width_pt") or LIMITS["geometry"]["stroke_default_pt"]),
            "dash": DASH_JS[str(stroke.get("dash") or "solid")], "cap": stroke.get("cap"), "arrow": arrow}


def _nice_floor(v: float) -> float:
    m = abs(v)
    mag = 10 ** math.floor(math.log10(m))
    for k in (1, 2, 2.5, 5, 10):
        if k * mag >= m:
            return -k * mag
    return -10 * mag


def _deemphasize(bg: str, kind: str) -> str:
    dark = chart_mode(bg) == "dark"
    base = LIMITS["chart"]["deemphasis"]["dark" if dark else "light"]
    floor = LIMITS["chart"]["min_contrast"].get(kind, 2.0)
    target = "#FFFFFF" if dark else "#000000"
    cand = base
    for step in range(25):
        cand = mix_toward(base, target, step / 24 * 0.45)
        if contrast_ratio(cand, bg) >= floor:
            return cand
    return cand


def label_format(chart: Dict[str, Any]) -> str:
    """直标带单位：读者看到“52”要能知道是百分比还是亿元。"""
    base = (chart.get("value_axis") or {}).get("number_format") or "#,##0.###"
    unit = chart_unit(chart)
    if chart.get("stacked") == "percent":
        return '0"%"'
    if not unit or "%" in base:
        return base
    if unit == "%":
        return base + '"%"'
    sep = "" if ord(unit[0]) >= 0x2E80 else " "
    return base + '"' + sep + unit + '"'


class Ctx:
    def __init__(self, plan, page, boxes, asset_root):
        self.plan, self.page, self.boxes, self.asset_root = plan, page, boxes, asset_root
        self.title_claimed = False


def _common(el: Dict[str, Any], ctx: Ctx, kind: str) -> Dict[str, Any]:
    out = {"kind": kind, "name": f"{ctx.page['id']}.{el['id']}"}
    if el.get("rotation"):
        out["rotation"] = float(el["rotation"]) % 360
    alt = alt_text(el)
    if alt:
        out["alt"] = alt
    return out


def _text(el, ctx: Ctx, top_level: bool) -> Dict[str, Any]:
    st = text_style(ctx.plan, ctx.page, el)
    out = _common(el, ctx, "text")
    out["box"] = list(explicit_box(el))
    if el["type"] == "title" and top_level and not ctx.title_claimed:
        ctx.title_claimed = True
        out["placeholder"] = "title"
    out.update(pad=list(st["pad"]), valign=st["valign"], align=st["align"], wrap=True, paras=[])
    for seg, pieces in paragraph_runs(el) or [(str(el.get("content", "")), [{"text": str(el.get("content", ""))}])]:
        para = {"align": st["align"], "line_spacing": st["line_spacing"], "runs": []}
        if el.get("space_after_pt"):
            para["space_after_pt"] = float(el["space_after_pt"])
        if el.get("list"):
            size = max([float(q.get("font_size") or st["size"]) for q in pieces] or [float(st["size"])])
            para["list"] = {"kind": str(el["list"]), "indent_in": list_indent_for(str(el["list"]), size),
                            "color": _hex((pieces[0].get("color") if pieces else None) or st["color"])}
        for piece in pieces:
            para["runs"].append({"text": piece.get("text", ""), "font": piece.get("font") or st["family"],
                                 "size": float(piece.get("font_size") or st["size"]),
                                 "color": _hex(piece.get("color") or st["color"]),
                                 "bold": bool(piece.get("bold", st["bold"])),
                                 "spc": float(piece.get("letter_spacing_pt", st["spacing"]))})
        para["fallback"] = {"font": st["family"], "size": float(st["size"]), "color": _hex(st["color"])}
        out["paras"].append(para)
    return out


def _shape(el, ctx: Ctx) -> Dict[str, Any]:
    out = _common(el, ctx, "shape")
    box = explicit_box(el)
    out.update(box=list(box), prst=PRST[el.get("shape_kind") or "rect"], line=_stroke(el.get("stroke")))
    if el.get("gradient"):
        g = el["gradient"]
        out["fill"] = {"color": _hex(g["from"]), "alpha": 1}  # 占位纯色，post.py 换成 gradFill
        out["ooxml"] = {"gradient": {"from": _hex(g["from"]), "to": _hex(g["to"]), "angle": float(g.get("angle", 0)),
                                     "from_alpha": g.get("from_alpha"), "to_alpha": g.get("to_alpha")}}
    elif el.get("color"):
        out["fill"] = {"color": _hex(el["color"]), "alpha": float(el.get("fill_alpha") or 1)}
    if el.get("corner_in") is not None:
        out.setdefault("ooxml", {})["adj"] = corner_adj(float(el["corner_in"]), box[2], box[3])
    return out


def _path(el, ctx: Ctx) -> Dict[str, Any]:
    pts = resolve_points(el, ctx.boxes)
    out = _common(el, ctx, "path")
    closed = bool(el.get("closed"))
    out.update(points=[[float(x), float(y)] for x, y in pts], closed=closed,
               line=_stroke(el.get("stroke"), str(el.get("arrow") or "none")))
    if closed and el.get("color"):
        out["fill"] = {"color": _hex(el["color"]), "alpha": float(el.get("fill_alpha") or 1)}
    return out


def _image(el, ctx: Ctx) -> Dict[str, Any]:
    path = resolve_asset(ctx.asset_root, el["image_id"])
    if path is None:
        raise FileNotFoundError(f"素材缺失 {el['image_id']}")
    with Image.open(path) as img:
        sw, sh = img.size
        if img.format not in ("PNG", "JPEG", "WEBP"):   # 按内容而非扩展名：PptxGenJS 依赖的 image-size 0.6–2.0.2 有 JXL/HEIF/ICNS 解析 DoS 公告（npm audit），不让这类内容进入引擎
            raise ValueError(f"素材 {el['image_id']} 的实际格式是 {img.format}，仅接受 PNG/JPEG/WEBP")
    box = explicit_box(el)
    l, r, t, b = crop_geometry(sw, sh, box[2], box[3], el.get("crop_focus"), el.get("crop_zoom"))
    out = _common(el, ctx, "image")
    out.update(box=list(box), path=str(path), alpha=float(el["alpha"]) if el.get("alpha") and float(el["alpha"]) < 1 else None)
    ox: Dict[str, Any] = {"crop": [l, r, t, b]}
    if el.get("duotone"):
        ox["duotone"] = {"dark": _hex(el["duotone"]["dark"]), "light": _hex(el["duotone"]["light"])}
    if el.get("mask"):
        m = el["mask"]
        ox["mask"] = MASK_PRESETS[m] if isinstance(m, str) else [[float(x), float(y)] for x, y in m]
    if el.get("soft_edge_pt"):
        ox["soft_edge_pt"] = float(el["soft_edge_pt"])
    out["ooxml"] = ox
    return out


def _table(el, ctx: Ctx, style, family) -> Dict[str, Any]:
    chart = el["chart"]
    out = _common(el, ctx, "table")
    heads, rows = chart["data"]["headers"], chart["data"]["rows"]
    weights = chart.get("column_weights") or [1] * len(heads)
    box = explicit_box(el)
    tot = float(sum(weights))
    out.update(box=list(box), col_w=[box[2] * w / tot for w in weights], row_h=box[3] / (len(rows) + 1),
               aligns=chart.get("column_align") or ["left"] * len(heads), font=family, size=LIMITS["chart"]["font_pt"]["table"],
               head_ink=_hex(style["title"]), body_ink=_hex(style["body"]), rule=_hex(style["rule"]), grid=_hex(style["grid"]),
               headers=[str(h) for h in heads], rows=[[str(v) for v in r] for r in rows])
    return out


def _chart(el, ctx: Ctx) -> Dict[str, Any]:
    chart, plan, page = el["chart"], ctx.plan, ctx.page
    kind = chart["kind"]
    bg = chart_surface(page, el)
    style = chart_style(bg)
    family = el.get("font") or (plan.get("fonts") or {}).get("body") or LIMITS["defaults"]["fonts"]["body"]
    if kind == "table":
        return _table(el, ctx, style, family)
    fp = LIMITS["chart"]["font_pt"]
    flags = chart_flags(chart)
    colors = [_hex(c) for c in chart_color_map(plan, chart)]
    cfg = chart.get("value_axis") or {}
    axis_on = chart.get("axis_labels", True)
    out = _common(el, ctx, "chart")
    series = chart["data"]["series"]
    combo = list(chart.get("combo_line_series") or []) if kind == "bar" else []
    out.update(box=list(explicit_box(el)), chart_kind=kind, family=family, font_pt=fp,
               style={k: _hex(v) for k, v in style.items()}, surface=_hex(bg), colors=colors,
               horizontal=bool(chart.get("horizontal")) and kind == "bar", stacked=chart.get("stacked") or False,
               combo=combo, legend=flags["legend"], data_labels=flags["data_labels"], axis_labels=bool(axis_on),
               title=(chart.get("title") if chart.get("show_title", True) and chart.get("title") else None),
               axis_titles=chart.get("axis_titles") or {}, gap_width=int(chart.get("gap_width") or LIMITS["chart"]["gap_width"]),
               overlap=LIMITS["chart"]["stacked_overlap"] if (chart.get("stacked") and kind in ("bar", "line", "area")) else None,
               label_format=label_format(chart), number_format=cfg.get("number_format"))
    if kind == "scatter":
        out["series"] = [{"name": str(s["name"]), "points": [[float(x), float(y)] for x, y in s["points"]]} for s in series]
    else:
        out["categories"] = [str(c) for c in chart["data"]["categories"]]
        out["series"] = [{"name": str(s["name"]), "values": [float(v) for v in s["values"]]} for s in series]
    if kind == "bar" and chart.get("highlight_category") is not None:
        hi, dim = int(chart["highlight_category"]), _hex(_deemphasize(bg, "bar"))
        out["point_colors"] = [colors[0] if j == hi else dim for j in range(len(chart["data"]["categories"]))]
    vmin = None
    if kind not in ("pie", "donut"):
        if cfg.get("min") is not None:
            vmin = float(cfg["min"])
        elif kind in ("bar", "area") and chart_non_negative(chart):
            vmin = 0.0
        elif kind == "bar" and not chart.get("stacked") and chart.get("data_labels") is not False:
            low = chart_min_value(chart)
            if low is not None and low < 0 and axis_on:
                vmin = _nice_floor(low * 1.25)
    out["val_axis"] = {"min": vmin, "max": cfg.get("max"), "major": cfg.get("major_unit")}
    if flags["data_labels"]:
        out["label_plan"] = [{**e, "ink": _hex(e["ink"]), "plate": _hex(e["plate"]) if e.get("plate") else None,
                              "points": [_hex(p) for p in e["points"]] if e.get("points") else None}
                             for e in chart_label_plan(plan, chart, bg)]
    out["combo_line_labels"] = bool(chart.get("combo_line_labels"))
    return out


def _legend(el, ctx: Ctx) -> Dict[str, Any]:
    page, plan = ctx.page, ctx.plan
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
    style = chart_style(chart_surface(page, el))
    ink = _hex(el.get("color") or style["body"])
    family = str(el.get("font") or (plan.get("fonts") or {}).get("body") or LIMITS["defaults"]["fonts"]["body"])
    size = float(el.get("font_size") or 18)
    swatch = min(float(el.get("swatch_in") or 0.16), box[2], box[3])
    gap = 0.14
    show_value = bool(el.get("show_value")) and spec.get("kind") in ("pie", "donut")
    values = list(((spec.get("data") or {}).get("series") or [{}])[0].get("values") or [])
    name = f"{page['id']}.{el['id']}"
    align = str(el.get("align") or "left")
    children: List[Dict[str, Any]] = []

    def add(i, x, y, w_lab, h_lab, sy):
        children.append({"kind": "shape", "name": f"{name}#sw{i}", "box": [x, sy, swatch, swatch], "prst": "rect",
                         "fill": {"color": _hex(colors[i]), "alpha": 1}, "line": None})
        children.append({"kind": "text", "name": f"{name}#tx{i}", "box": [x + swatch + gap, y, max(w_lab, 0.3), h_lab],
                         "pad": [0, 0], "valign": "middle", "align": align, "wrap": True,
                         "paras": [{"align": align, "line_spacing": 1.1, "fallback": {"font": family, "size": size, "color": ink},
                                    "runs": [{"text": legend_label(spec, items[i], values, i, show_value), "font": family,
                                              "size": size, "color": ink, "bold": False, "spc": 0.0}]}]})

    if str(el.get("direction") or "vertical") == "horizontal":
        col = box[2] / len(items)
        for i in range(len(items)):
            cx = box[0] + i * col
            add(i, cx, box[1], col - swatch - gap, box[3], box[1] + box[3] / 2 - swatch / 2)
    else:
        row = box[3] / len(items)
        for i in range(len(items)):
            top = box[1] + i * row
            add(i, box[0], top, box[2] - swatch - gap, row, top + max((row - swatch) / 2, 0))
    out = _common(el, ctx, "group")
    out["children"] = children
    return out


def _element(el, ctx: Ctx, top_level: bool = True) -> Dict[str, Any]:
    t = el["type"]
    if t in TEXT_TYPES:
        return _text(el, ctx, top_level)
    if t in ("shape", "divider"):
        return _shape(el, ctx)
    if t == "path":
        return _path(el, ctx)
    if t == "image":
        return _image(el, ctx)
    if t == "chart":
        return _chart(el, ctx)
    if t == "legend":
        return _legend(el, ctx)
    if t == "group":
        out = _common(el, ctx, "group")
        out["children"] = [_element(c, ctx, False) for c in el["elements"]]
        return out
    raise ValueError(f"未知元素类型：{t}")


def _page_number(plan, page, index) -> Optional[Dict[str, Any]]:
    cfg = page_numbers(plan)
    if page["id"] in cfg["hide"]:
        return None
    b = pn_box(plan, page)
    family = (plan.get("fonts") or {}).get("body") or LIMITS["defaults"]["fonts"]["body"]
    pt, ink = LIMITS["page_number"]["font_pt"], _hex(pn_ink(plan, page))
    return {"kind": "text", "name": "_page_number", "box": [b["x"], b["y"], b["w"], b["h"]], "pad": [0, 0], "valign": "middle",
            "align": "right", "wrap": False,
            "paras": [{"align": "right", "line_spacing": 1.0, "fallback": {"font": family, "size": pt, "color": ink},
                       "runs": [{"text": str(cfg["start"] + index), "font": family, "size": pt, "color": ink,
                                 "bold": False, "spc": 0.0}]}]}


def _theme(plan) -> Dict[str, Any]:
    fonts = plan.get("fonts") or {}
    pal = list(deck_palette(plan))
    for extra in LIMITS["defaults"]["palette"]:
        if len(pal) >= 6:
            break
        if extra.upper() not in [c.upper() for c in pal]:
            pal.append(extra)
    bgs = [page_bg(p).upper() for p in plan.get("pages") or []] or [LIMITS["defaults"]["background"].upper()]
    light = [b for b in bgs if not is_dark(b)]
    base = Counter(bgs).most_common(1)[0][0]
    colors = {"dk2": LIMITS["defaults"]["ink"], "lt2": (light[0] if light else "#F3F4F6"), "hlink": pal[0],
              "folHlink": pal[min(3, len(pal) - 1)]}
    for i in range(6):
        colors[f"accent{i + 1}"] = pal[i]
    return {"heading": fonts.get("heading") or LIMITS["defaults"]["fonts"]["heading"],
            "body": fonts.get("body") or LIMITS["defaults"]["fonts"]["body"],
            "colors": {k: _hex(v) for k, v in colors.items()}, "master_bg": _hex(base),
            "master_ink": _hex(LIMITS["defaults"]["ink_on_dark"] if is_dark(base) else LIMITS["defaults"]["ink"]),
            "title_pt": LIMITS["type"]["default_pt"]["title"]}


def build_spec(plan: Dict[str, Any], views: List[Dict[str, Any]], asset_root: Path, motion_mode: str = "full") -> Dict[str, Any]:
    pages = []
    for idx, page in enumerate(plan["pages"]):
        ctx = Ctx(plan, page, views[idx]["boxes"], asset_root)
        els = [_element(el, ctx) for el in page["elements"]]
        pages.append({"id": page["id"], "bg": _hex(page_bg(page)), "elements": els,
                      "page_number": _page_number(plan, page, idx), "notes": page.get("speaker_notes") or "",
                      "motion": motion.normalize(page, lambda eid, pid=page["id"]: f"{pid}.{eid}", motion_mode)})
    brief = plan.get("brief") or {}
    return {"version": SPEC_VERSION,
            "canvas": {"w_emu": CANVAS_EMU[0], "h_emu": CANVAS_EMU[1], "w_in": CANVAS_EMU[0] / EMU, "h_in": CANVAS_EMU[1] / EMU},
            "theme": _theme(plan),
            "core": {"title": plan.get("title", ""), "author": "ppt-visual-art-director",
                     "comments": (brief.get("concept") or "")[:500]},
            "pages": pages}
