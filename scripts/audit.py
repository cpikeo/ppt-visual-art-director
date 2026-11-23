#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""设计审计与结构体检。三类事实，都不下审美结论，只指出“该看哪里”：

1. plan_signals   plan 层：字号档位、主导关系、边距、对齐轴、行长、密度、字族数。
2. render_metrics 渲染层：每页的墨量、留白、重心、贴边；输出灰度总览（有效性检查用）。
3. lint_ooxml     包结构：关键 OOXML 容器的子元素次序。次序错了，PowerPoint 会提示“需要修复”，
                  LibreOffice 却照常打开——这是“预览正常、目标机报错”的典型来源。

阈值全部来自 limits.json 的 audit 段。
"""

from __future__ import annotations

import zipfile
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Tuple

from model import LIMITS, TEXT_TYPES, Report, canvas, flatten, is_dark, page_bg, page_geometry, parse_xml, text_style

CFG = LIMITS["audit"]


# ---------------------------------------------------------------- 1. plan 层


def _top_size(el: Dict[str, Any], st: Dict[str, Any]) -> float:
    sizes = [float(r.get("font_size") or st["size"]) for r in (el.get("runs") or [])]
    return max(sizes or [st["size"]])


def plan_signals(plan: Dict[str, Any]) -> Report:
    rep = Report()
    pages = plan.get("pages") or []
    W, H = canvas(plan)
    families: set = set()
    sizes: Counter = Counter()
    title_sizes: Counter = Counter()
    edge = CFG["edge_margin_in"]
    tol = LIMITS["canvas"]["bound_tolerance_in"]

    for page in pages:
        pid = str(page.get("id"))
        boxes = page_geometry(page)
        texts: List[Tuple[Dict[str, Any], Dict[str, Any], Tuple[float, float, float, float]]] = []
        mass_area = 0.0
        for item in flatten(page):
            el = item["el"]
            box = boxes.get(str(el.get("id")))
            if box is None:
                continue
            if el["type"] in ("image", "chart") and box[2] * box[3] >= CFG["dominant_area_share"] * W * H:
                mass_area = max(mass_area, box[2] * box[3])
            if el["type"] in TEXT_TYPES:
                texts.append((el, text_style(plan, page, el), box))

        top_title_done = False
        for el, st, box in texts:
            families.add(str(st["family"]))
            for run in el.get("runs") or []:
                if run.get("font"):
                    families.add(str(run["font"]))
            if el["type"] != "number":
                sizes[round(_top_size(el, st))] += 1
            if el["type"] == "title" and not top_title_done:
                title_sizes[round(st["size"])] += 1
                top_title_done = True

        # 边距：文字墨迹边缘离画布边的距离。声明了 overlap_allowed 的元素是出血承载层，不算。
        for el, st, box in texts:
            if el.get("overlap_allowed"):
                continue
            px, py = st["pad"]
            left, top = box[0] + px, box[1] + py
            right, bottom = W - (box[0] + box[2] - px), H - (box[1] + box[3] - py)
            if min(left, right) < edge - tol or min(top, bottom) < CFG["edge_margin_v_in"] - tol:
                rep.soft("W-MARGIN", f"{pid}/{el['id']}",
                         f"文字墨迹离画布边 左{left:.2f} 右{right:.2f} 上{top:.2f} 下{bottom:.2f}in，"
                         f"低于 {edge}in（竖向 {CFG['edge_margin_v_in']}in）：投影边缘与机位偏差会吃掉它")

        # 对齐轴：文字左缘归并后的轴数
        axes: List[float] = []
        for el, st, box in sorted(texts, key=lambda t: t[2][0] + t[1]["pad"][0]):
            x = box[0] + st["pad"][0]
            if not axes or x - axes[-1] > CFG["axis_cluster_in"]:
                axes.append(x)
        if len(axes) > CFG["max_axes"]:
            rep.soft("W-AXIS", pid, f"文字左缘有 {len(axes)} 条对齐轴（上限 {CFG['max_axes']}）：网格没有被执行，还是每块文字各找各的位置？")

        # 主导关系：最大字号 vs 次大字号；已有占画布 ≥ 阈值的图或表时，主导由它承担
        if len(texts) >= 2 and not mass_area and page.get("role") != "breathe":
            ranked = sorted({round(_top_size(el, st)) for el, st, _b in texts}, reverse=True)
            if len(ranked) == 1 or ranked[0] / ranked[1] < CFG["min_dominance_ratio"]:
                second = ranked[1] if len(ranked) > 1 else ranked[0]
                rep.soft("W-DOMINANCE", pid,
                         f"最大字号 {ranked[0]}pt 与次大 {second}pt 之比 {ranked[0] / second:.2f}，"
                         f"低于 {CFG['min_dominance_ratio']}：先看到什么？尺度、明度或位置要有一项明显压过其余")

        # 行长与密度
        total_chars = 0
        for el, st, box in texts:
            text = str(el.get("content") or "")
            chars = sum(1 for c in text if not c.isspace())
            if el["type"] in ("body", "text", "title", "number"):
                total_chars += chars
            if el["type"] in ("body", "text") and chars > 40:
                per_line = (box[2] - 2 * st["pad"][0]) / (st["size"] / 72.0)
                if per_line > CFG["max_measure_cjk"]:
                    rep.soft("W-MEASURE", f"{pid}/{el['id']}",
                             f"每行约 {per_line:.0f} 字（上限 {CFG['max_measure_cjk']}）：行太长，视线回扫困难；收窄栏宽或分栏")
        if total_chars > CFG["max_chars_page"]:
            rep.soft("W-DENSITY", pid, f"本页正文与标题共 {total_chars} 字（上限 {CFG['max_chars_page']}）：投影页更像读物了，删到论断＋证据，或拆页")

    if len(families) > CFG["max_families"]:
        rep.soft("W-FONTS", "整稿", f"用了 {len(families)} 个字族 {sorted(families)}（上限 {CFG['max_families']}）：体系靠角色，不靠字族数量")
    steps = sorted(sizes)
    if len(steps) > CFG["max_type_steps"]:
        rep.soft("W-SCALE", "整稿", f"正文与标题用了 {len(steps)} 档字号 {steps}（上限 {CFG['max_type_steps']}）：用字重、墨阶与位置分层，别再加一档")
    close = [(a, b) for a, b in zip(steps, steps[1:]) if b / a < CFG["min_step_ratio"]]
    if close:
        rep.soft("W-SCALE", "整稿", f"近似字号并存 {close}：差不到 {CFG['min_step_ratio']} 倍，读者看不出层级，只会觉得没对齐")
    if len(title_sizes) > CFG["max_title_sizes"]:
        rep.soft("W-SCALE", "整稿", f"页标题有 {len(title_sizes)} 种字号 {sorted(title_sizes)}：标题是系统的一部分，逐页改会让整稿散掉")

    # 明暗节奏：同一明度到底，转折与停顿就没有底色可依
    if len(pages) >= 6 and len({is_dark(page_bg(p)) for p in pages}) == 1:
        rep.info("I-TONE", "整稿", "全稿同一明暗场：转折、停顿与高潮只能靠尺度承担；确认这是有意的单色系统")
    return rep


# ---------------------------------------------------------------- 2. 渲染层


def render_metrics(pngs: List[Path], plan: Dict[str, Any], render_dir: Path) -> Tuple[List[Dict[str, Any]], Report]:
    """逐页墨量、留白、重心、贴边。输出灰度总览 contact_sheet_gray.png。"""
    import numpy as np
    from PIL import Image

    rep = Report()
    rows: List[Dict[str, Any]] = []
    pages = plan.get("pages") or []
    W, _H = canvas(plan)
    grays: List[Any] = []
    for i, path in enumerate(pngs):
        page = pages[i] if i < len(pages) else {}
        with Image.open(path) as img:
            g = img.convert("L")
            grays.append(g.copy())
        a = np.asarray(g, dtype=np.float32)
        h, w = a.shape
        bw = max(2, int(min(h, w) * 0.01))
        border = np.concatenate([a[:bw].ravel(), a[-bw:].ravel(), a[:, :bw].ravel(), a[:, -bw:].ravel()])
        bg = float(np.median(border))
        mask = np.abs(a - bg) > CFG["ink_threshold"]
        mask[:3], mask[-3:], mask[:, :3], mask[:, -3:] = False, False, False, False  # 渲染边缘 1–2px 接缝不是内容
        cover = float(mask.mean())
        ys, xs = np.nonzero(mask)
        row: Dict[str, Any] = {"page": page.get("id", f"p{i + 1}"), "luminance": round(float(a.mean()) / 255, 3),
                               "ink": round(cover, 3)}
        if len(xs):
            ppi = w / W
            row["margins_in"] = {"left": round(float(xs.min()) / ppi, 2), "right": round(float(w - 1 - xs.max()) / ppi, 2),
                                 "top": round(float(ys.min()) / ppi, 2), "bottom": round(float(h - 1 - ys.max()) / ppi, 2)}
            row["centroid"] = {"dx": round(float(xs.mean()) / w - 0.5, 3), "dy": round(float(ys.mean()) / h - 0.5, 3)}
        rows.append(row)
        pid = str(row["page"])
        has_bleed = any(it["el"].get("overlap_allowed") or it["el"].get("type") == "image" for it in flatten(page)) if page else False
        if cover < CFG["empty_ink"] and page.get("role") not in ("breathe", "hold"):
            rep.soft("R-EMPTY", pid, f"渲染墨量 {cover:.1%}：这页几乎是空的，主角有分量吗？")
        if cover > CFG["crowd_ink"] and not has_bleed:
            rep.soft("R-CROWD", pid, f"渲染墨量 {cover:.1%}：太满，没有留白可言")
        if "margins_in" in row and not has_bleed and min(row["margins_in"].values()) < CFG["render_edge_in"]:
            rep.soft("R-EDGE", pid, f"渲染内容离页边最近 {min(row['margins_in'].values()):.2f}in：贴边且没有声明出血")
        if "centroid" in row and not has_bleed and not page.get("composition_rationale") and (
                abs(row["centroid"]["dx"]) > CFG["centroid_dx"] or abs(row["centroid"]["dy"]) > CFG["centroid_dy"]):
            rep.info("R-BALANCE", pid, f"视觉重心偏离中心 dx={row['centroid']['dx']:+.2f} dy={row['centroid']['dy']:+.2f}，composition_rationale 没写偏置理由")
    if len(rows) >= 6:
        lums = [r["luminance"] for r in rows]
        if max(lums) - min(lums) < CFG["flat_tone_range"]:
            rep.info("R-TONE", "整稿", f"全稿页均明度只在 {min(lums):.2f}–{max(lums):.2f}：灰度缩略图里几乎没有节奏")

    # 灰度总览：有效性检查——层级、类别、状态转成灰度后还辨得出吗
    if grays:
        cols = 2 if len(grays) <= 4 else 3
        tw = LIMITS["preview"]["contact_sheet_width_px"] // cols
        th = int(tw * grays[0].height / grays[0].width)
        rws = (len(grays) + cols - 1) // cols
        pad = 24
        sheet = Image.new("L", (cols * tw + pad * (cols + 1), rws * th + pad * (rws + 1)), 233)
        for i, g in enumerate(grays):
            sheet.paste(g.resize((tw, th), Image.LANCZOS), (pad + (i % cols) * (tw + pad), pad + (i // cols) * (th + pad)))
        sheet.save(render_dir / "contact_sheet_gray.png")
    return rows, rep


# ---------------------------------------------------------------- 3. OOXML 次序体检

_FILL = ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")


def _rank(*groups: Tuple[str, ...]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for i, names in enumerate(groups):
        for n in names:
            out[n] = i
    return out


_RPR = _rank(("ln",), _FILL, ("effectLst", "effectDag"), ("highlight",), ("uLnTx", "uLn"), ("uFillTx", "uFill"),
             ("latin",), ("ea",), ("cs",), ("sym",), ("hlinkClick",), ("hlinkMouseOver",), ("rtl",), ("extLst",))
_AX_TAIL = (("crossAx",), ("crosses", "crossesAt"))
_ORDER: Dict[str, Dict[str, int]] = {
    "rPr": _RPR, "endParaRPr": _RPR, "defRPr": _RPR,
    "ln": _rank(_FILL, ("prstDash", "custDash"), ("round", "bevel", "miter"), ("headEnd",), ("tailEnd",), ("extLst",)),
    "spPr": _rank(("xfrm",), ("custGeom", "prstGeom"), _FILL, ("ln",), ("effectLst", "effectDag"), ("scene3d",), ("sp3d",), ("extLst",)),
    "pPr": _rank(("lnSpc",), ("spcBef",), ("spcAft",), ("buClrTx", "buClr"), ("buSzTx", "buSzPct", "buSzPts"),
                 ("buFontTx", "buFont"), ("buNone", "buAutoNum", "buChar", "buBlip"), ("tabLst",), ("defRPr",), ("extLst",)),
    "bodyPr": _rank(("prstTxWarp",), ("noAutofit", "normAutofit", "spAutoFit"), ("scene3d",), ("sp3d",), ("flatTx",), ("extLst",)),
    "gradFill": _rank(("gsLst",), ("lin", "path"), ("tileRect",)),
    "dLbls": _rank(("dLbl",), ("delete",), ("numFmt",), ("spPr",), ("txPr",), ("dLblPos",), ("showLegendKey",), ("showVal",),
                   ("showCatName",), ("showSerName",), ("showPercent",), ("showBubbleSize",), ("separator",),
                   ("showLeaderLines",), ("leaderLines",), ("extLst",)),
    "dLbl": _rank(("idx",), ("delete", "layout"), ("tx",), ("numFmt",), ("spPr",), ("txPr",), ("dLblPos",), ("showLegendKey",),
                  ("showVal",), ("showCatName",), ("showSerName",), ("showPercent",), ("showBubbleSize",), ("separator",), ("extLst",)),
    "ser": _rank(("idx",), ("order",), ("tx",), ("spPr",), ("marker", "invertIfNegative"), ("pictureOptions", "explosion"),
                 ("dPt",), ("dLbls",), ("trendline",), ("errBars",), ("cat", "xVal"), ("val", "yVal"), ("shape", "smooth"), ("extLst",)),
    "valAx": _rank(("axId",), ("scaling",), ("delete",), ("axPos",), ("majorGridlines",), ("minorGridlines",), ("title",), ("numFmt",),
                   ("majorTickMark",), ("minorTickMark",), ("tickLblPos",), ("spPr",), ("txPr",), *_AX_TAIL,
                   ("crossBetween",), ("majorUnit",), ("minorUnit",), ("dispUnits",), ("extLst",)),
    "catAx": _rank(("axId",), ("scaling",), ("delete",), ("axPos",), ("majorGridlines",), ("minorGridlines",), ("title",), ("numFmt",),
                   ("majorTickMark",), ("minorTickMark",), ("tickLblPos",), ("spPr",), ("txPr",), *_AX_TAIL,
                   ("auto",), ("lblAlgn",), ("lblOffset",), ("tickLblSkip",), ("tickMarkSkip",), ("noMultiLvlLbl",), ("extLst",)),
}
_PARTS = ("ppt/slides/", "ppt/charts/chart", "ppt/slideLayouts/", "ppt/slideMasters/")


def lint_ooxml(path: Path, limit: int = 20) -> List[str]:
    """子元素次序违例清单。次序以 ECMA-376 的 sequence 为准；未知子元素不判。"""
    problems: List[str] = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if not name.endswith(".xml") or not name.startswith(_PARTS):
                continue
            try:
                root = parse_xml(zf.read(name))
            except Exception as exc:  # noqa: BLE001
                problems.append(f"{name}: XML 无法解析：{exc}")
                continue
            for node in root.iter():
                tag = node.tag
                if not isinstance(tag, str):
                    continue
                parent = tag.rsplit("}", 1)[-1]
                table = _ORDER.get(parent)
                if not table:
                    continue
                last, last_name = -1, ""
                for child in node:
                    if not isinstance(child.tag, str):
                        continue
                    local = child.tag.rsplit("}", 1)[-1]
                    r = table.get(local)
                    if r is None:
                        continue
                    if r < last:
                        problems.append(f"{name}: <{parent}> 中 {local} 排在 {last_name} 之后（应在其前）")
                        break
                    last, last_name = r, local
                if len(problems) >= limit:
                    return problems
    return problems
