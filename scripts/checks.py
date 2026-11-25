#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""plan 校验：硬错误（E-）、软提示（W-）与整稿模式信号。只读 plan，不写文件。"""

from __future__ import annotations

import hashlib
import math
import re
from pathlib import Path
from typing import Any, Dict, List, Tuple


from model import chart_data_problems, LABEL_INK_MIN, backdrop_image, zone_contrast, zone_probe, page_bg, CARRIERS, LIMITS, Report, TEXT_TYPES, _intersection, _union, all_elements, canvas, chart_color_map, chart_flags, chart_readout_problems, gray_conflicts, chart_style, chart_surface, chart_surface_known, chart_items, chart_palette, contrast_ratio, explicit_box, fit_height_in, flatten, image_probe, page_geometry, page_numbers, paragraph_runs, pn_box, resolve_asset, resolve_bg, resolve_points, rot_box, rotated_ancestors, substantive_overlap, text_style


def validate(plan: Dict[str, Any], asset_root: Path, production: bool = False) -> Report:
    rep = Report()
    w, h = canvas(plan)
    tol = LIMITS["canvas"]["bound_tolerance_in"]
    pages = plan.get("pages") or []
    assets = plan.get("assets") or {}
    pn_cfg = page_numbers(plan)

    rhythm = plan.get("rhythm")
    if rhythm and len(rhythm) != len(pages):
        rep.soft("W-RHYTHM", "整稿", f"rhythm 有 {len(rhythm)} 条，页数 {len(pages)}；按页一一对应")
    brief = plan.get("brief") or {}
    uses_images = any(el.get("type") == "image" for p in pages for el in all_elements(p))
    if uses_images and not brief.get("art"):
        rep.info("I-ART", "brief", "本稿有图像，但没有写艺术指导五参数（光/景别/材质/色域/颗粒）：图的选法没有共同世界")
    if brief.get("art") and not uses_images:
        rep.info("I-ART", "brief", "写了图像世界但没有用图：确认这是纯排版稿，或补上素材")
    page_ids = {p.get("id") for p in pages}
    for hidden in pn_cfg["hide"]:
        if hidden not in page_ids:
            rep.soft("W-PAGE", "页码", f"hide 里的“{hidden}”不是本稿页 id")

    seen_pages: set = set()
    for pi, page in enumerate(pages):
        pid = page.get("id", f"p{pi + 1}")
        if pid in seen_pages:
            rep.hard("E-ID", pid, f"页面 id 重复：{pid}")
        seen_pages.add(pid)

        elements = page.get("elements") or []
        if not elements:
            rep.soft("W-EMPTY", pid, "本页没有任何元素：成品只有页码，确认这是有意的空页")
        ids: set = set()
        for el in all_elements(page):
            eid = el.get("id", "?")
            if eid in ids:
                rep.hard("E-ID", f"{pid}/{eid}", f"元素 id 在本页重复（含 group 子元素）：{eid}")
            ids.add(eid)

        focus = page.get("focus_element")
        if focus and focus not in ids:
            rep.hard("E-REF", pid, f"focus_element「{focus}」不是本页元素 id")
        _check_depth(rep, pid, page, ids)
        _check_design_intent(rep, pid, page, ids, production)

        boxes = page_geometry(page)
        if focus and focus in ids:
            _check_dominance(rep, pid, page, focus, boxes)

        for item in flatten(page):
            el = item["el"]
            eid = el.get("id", "?")
            eloc = f"{pid}/{eid}"
            kind = el.get("type")
            box = boxes.get(eid)
            if box is None:
                if kind == "path":
                    rep.hard("E-REF", eloc, "路径顶点无法解析（锚点指向不存在的元素，或点少于 2 个）")
                else:
                    rep.hard("E-GEOM", eloc, "缺少坐标，或 group 没有可用的子元素")
                continue
            x, y, ew, eh = box
            rbox = rot_box(box, el.get("rotation"))
            if x + ew <= 0 or y + eh <= 0 or x >= w or y >= h:
                rep.hard("E-GEOM", eloc, f"元素完全在画布外：x={x} y={y} w={ew} h={eh}")
            elif kind in TEXT_TYPES + ("chart",) and (rbox[0] < -tol or rbox[1] < -tol
                                                      or rbox[0] + rbox[2] > w + tol
                                                      or rbox[1] + rbox[3] > h + tol):
                what = "文本框" if kind in TEXT_TYPES else "图表"
                rep.soft("W-GEOM", eloc, f"{what}（含旋转）伸出画布（{w}×{h}），边缘内容会被裁掉")

            display_text = kind in ("number", "title")  # 巨型数字/标题可以是底层图形（字体即图像）
            if el.get("overlap_allowed") and kind not in CARRIERS and not display_text and not (
                    kind == "image" and el.get("image_role") in ("background", "texture", "hero")):
                rep.hard("E-OVERLAP", eloc,
                         "overlap_allowed 只能声明在承载层（背景图、整幅主视觉/蒙版/色块/闭合路径/做底的巨型数字或标题）上；正文、注释与图表被压时不能自我豁免")
            if kind == "divider" and eh > 0.2:
                rep.soft("W-GEOM", eloc, f"divider 高 {eh:.2f}in，已经是一块色块；用 shape 更清楚")
            if kind == "group":
                kids = el.get("elements") or []
                if not kids:
                    rep.hard("E-GEOM", eloc, "group 没有子元素")
                declared = explicit_box(el)
                computed = _union([boxes[str(c.get("id"))] for c in kids if str(c.get("id")) in boxes] or [])
                if declared and computed and any(abs(a - b) > 0.05 for a, b in zip(declared, computed)):
                    rep.info("I-GROUP", eloc, "group 的几何由子元素推出，声明的 x/y/w/h 未采用")

            if kind in TEXT_TYPES:
                check_text(plan, page, el, eloc, ew, eh, asset_root, rep)
            elif kind == "image":
                check_image(el, eloc, ew, eh, assets, asset_root, rep)
            elif kind == "shape" and not (el.get("color") or el.get("gradient") or el.get("stroke")):
                rep.hard("E-SHAPE", eloc, "shape 既无填充也无描边，成品里看不见")
            elif kind == "path":
                check_path(page, el, eloc, boxes, rep)
            elif kind == "chart":
                check_chart(eloc, el, plan, page, rep)
            elif kind == "legend":
                check_legend(plan, page, el, eloc, rep)

        elems = all_elements(page)
        has_legend_el = any(e.get("type") == "legend" for e in elems)
        # 手绘图例色块的识别：小方块 + 右侧标签 + 颜色不在本页任何图表的颜色映射里。
        # 出现即提示——手绘色块没有数据对应关系，是"扇区与图例错色"的高发形态。
        charts_here = [e for e in elems if e.get("type") == "chart" and e["chart"].get("kind") != "table"]
        if charts_here and not has_legend_el:
            mapped = {c for ch in charts_here for c in chart_color_map(plan, ch["chart"])}
            texts = [e for e in elems if e.get("type") in TEXT_TYPES and e.get("id")]
            for e in elems:
                if e.get("type") != "shape" or not e.get("color"):
                    continue
                b = boxes.get(str(e.get("id")))
                if b is None or b[2] > 0.6 or b[3] > 0.6 or abs(b[2] - b[3]) > 0.25:
                    continue
                color = str(e["color"]).upper()
                if color in mapped:
                    continue
                if not any(abs(t["x"] - (b[0] + b[2])) <= 0.6 and abs(t["y"] - b[1]) <= 0.35 for t in texts):
                    continue
                rep.soft("W-CHART", f"{pid}/{e.get('id')}",
                         f"疑似手绘图例色块 {color}，不在本页图表的颜色映射内：与扇区/系列无法核对；"
                         f"改用 legend 元素（色块由同一份映射生成），或给图表写 color_map")

        if pid not in pn_cfg["hide"]:
            box = pn_box(plan, page)
            for item in flatten(page):
                el = item["el"]
                if el.get("overlap_allowed"):
                    continue  # 承载层在下，页码本来就压在它上面
                gb = boxes.get(el.get("id"))
                if gb is None:
                    continue
                gb = rot_box(gb, el.get("rotation"))
                ix, iy = _intersection(gb, (box["x"], box["y"], box["w"], box["h"]))
                if ix <= 0.02 or iy <= 0.02:
                    continue
                where = f"{pid}/{el.get('id')}"
                if el["type"] in TEXT_TYPES or el["type"] == "chart":
                    rep.hard("E-PAGE", where, "文字或图表与页码保留框相交；移开它，或整稿改用另一角区")
                elif el["type"] == "image":
                    rep.info("I-PAGE", where, "图片压到页码角区：用 page_number_tone 声明实际明暗，并确认页码可读")
                elif el["type"] != "path":
                    rep.soft("W-PAGE", where, "色块压到页码角区：确认页码在它上面可读，必要时换角区")

    # 跨元素：实质重叠（叠放顺序即声明顺序）。同一 group 内部的叠放是设计许可的，不查。
    for page in pages:
        pid = page.get("id")
        items = flatten(page)
        boxes = page_geometry(page)
        for i, lower in enumerate(items):
            lbox = boxes.get(lower["el"].get("id"))
            if lbox is None or lower["el"].get("overlap_allowed") or lower["el"]["type"] == "path":
                continue
            lbox = rot_box(lbox, lower["el"].get("rotation"))
            for upper in items[i + 1:]:
                if upper["group"] is not None and upper["group"] == lower["group"]:
                    continue
                ubox = boxes.get(upper["el"].get("id"))
                if ubox is None:
                    continue
                ubox = rot_box(ubox, upper["el"].get("rotation"))
                if substantive_overlap(lbox, ubox):
                    rep.hard("E-OVERLAP", f"{pid}/{upper['el'].get('id')}",
                             f"与先声明的 {lower['el'].get('id')} 实质重叠；若后者是承载层（背景图/蒙版/色块），"
                             f"在 {lower['el'].get('id')} 上声明 overlap_allowed，否则移开")
    spaced = rep.__dict__.get("_ispace") or []
    if spaced:  # 单个空格是整稿的风格取舍：每稿只报一条，列出位置，不逐元素刷屏
        where = "、".join(l for l, _ in spaced[:8]) + (f" 等 {len(spaced)} 处" if len(spaced) > 8 else "")
        rep.info("I-SPACE", "deck", f"{where}：中西文之间手敲了单个空格（如「{spaced[0][1]}」）。这是通行的中文排法，"
                                    f"但 LibreOffice 会叠加自动间距显得更松、PowerPoint 里正常；全稿统一取一种并写进交付说明")
    return rep

# 中西文间的手敲空格：字体自带字距已够，空格宽度又随渲染器而异（Noto CJK 约 0.22–0.26 em；宋体 0.5 em）。
# 分级：单个普通 ASCII 空格是中文排版的常规做法（中西文之间加空格是通行排法，
# PowerPoint 里也不显松），只在 LibreOffice 渲染中叠加自动间距才显松——这是渲染器相关的
# 风格取舍，不是缺陷，降为 I-SPACE 提示；连续两个及以上空格、不断行空格（ ）
# 与全角空格（　）宽度不可控，在任何渲染器里都不稳，仍是必须处置的 W-SPACE。
_NBSP = chr(0x00A0)          # 不断行空格
_IDEO_SPACE = chr(0x3000)    # 全角空格
_CJK_SPACE2 = re.compile(
    r"[㐀-鿿][  " + _IDEO_SPACE + r"]+[A-Za-z0-9]"
    r"|[A-Za-z0-9%][  " + _IDEO_SPACE + r"]+[㐀-鿿]"
)


def classify_cjk_spaces(content: Any) -> Tuple[List[str], List[str]]:
    """把命中的中西文间空格分成 (须处置的 W-SPACE 匹配, 属风格取舍的 I-SPACE 匹配)。

    单个 ASCII 空格 = 中文排版常规，降为 I-SPACE；其余（多空格、不断行空格、全角空格）
    宽度不可控，仍是 W-SPACE。
    """
    hard: List[str] = []
    soft: List[str] = []
    for match in _CJK_SPACE2.finditer(str(content)):
        seg = match.group(0)
        run = re.search(r"[  " + _IDEO_SPACE + r"]+", seg)
        spaces = run.group(0) if run else " "
        if len(spaces) > 1 or _NBSP in spaces or _IDEO_SPACE in spaces:
            hard.append(seg)
        else:
            soft.append(seg)
    return hard, soft


# 中文语境里的半角标点：与汉字相邻的 , ; : ! ? 在多数字体里前后留白异常，渲染后像“字 , 字”
_ASCII_PUNCT = re.compile(r"[㐀-鿿][,;:!?]|[,;:!?](?=[㐀-鿿])")


def check_text(plan, page, el, loc, ew, eh, asset_root: Path, rep: Report) -> None:
    content = el.get("content", "")
    if not str(content).strip():
        rep.hard("E-TEXT", loc, "文本内容为空；删掉这个元素，或写入内容")
    if paragraph_runs(el) is None:
        rep.hard("E-REF", loc, "runs 拼接与 content 不一致（含 \\n）；分段要对得上")
    st = text_style(plan, page, el)
    need = fit_height_in(el, st, ew)
    if need > eh * LIMITS["type"]["fit_tolerance"]:
        rep.soft("W-FIT", loc, f"按 {st['size']:g}pt 估算需 {need:.2f}in，框高 {eh:.2f}in：先编辑文案，再加宽/拆页")
    punct = _ASCII_PUNCT.findall(str(content))
    if punct:
        rep.soft("W-PUNCT", loc, f"汉字旁用了半角标点（{'、'.join(dict.fromkeys(punct))[:20]}）：渲染时前后留白不稳，"
                                 f"改全角（，；：！？）；西文与数字之间的半角不受影响")
    hard_space, soft_space = classify_cjk_spaces(content)
    if hard_space:
        rep.soft("W-SPACE", loc, f"中文与数字/西文之间有手敲空格（{'、'.join(hard_space[:3])}）："
                                 f"多空格、不断行空格与全角空格的宽度随渲染器变，任何环境下都不稳；"
                                 f"删掉它们，需要时用 runs 调字距")
    elif soft_space:
        rep.__dict__.setdefault("_ispace", []).append((loc, soft_space[0]))  # 汇总成一条，见 validate()
    minimum = LIMITS["type"]["min_pt"].get(el["type"], 0)
    if st["size"] < minimum:
        rep.soft("W-PT", loc, f"{el['type']} {st['size']:g}pt 低于 {minimum}pt 下限")
    bg = resolve_bg(page, el)
    if bg is None:
        check_text_on_image(plan, page, el, loc, st, asset_root, rep)
    else:
        limit = LIMITS["contrast"]["large"] if st["size"] >= LIMITS["contrast"]["large_pt"] \
            else LIMITS["contrast"]["normal"]
        ratio = contrast_ratio(st["color"], bg)
        if ratio < limit:
            rep.soft("W-CONTRAST", loc, f"{st['color']} on {bg} = {ratio:.2f}:1，低于 {limit}:1")

def check_text_on_image(plan, page, el, loc, st, asset_root: Path, rep: Report) -> None:
    """文字压在图上：量文字框正下方的真实像素，而不是只提示“去看渲染”。

    量不了（图与字之间隔着蒙版、读不到图）才退回 I-CONTRAST。量得了也只是静态近似：
    达标不等于好看，不达标说明字落点不在安静区，先改构图或简报，不靠加浓遮罩。
    """
    under = backdrop_image(page, el)
    box = explicit_box(el)
    asset = (plan.get("assets") or {}).get(str(under[0].get("image_id"))) if under else None
    brief = ((asset or {}).get("generation") or {}).get("brief") or {}
    if under and brief.get("quiet_zone") == "none":
        rep.soft("W-ASSET", loc, f"{under[0].get('image_id')} 的简报声明不叠字（quiet_zone: none），但这里有文字压在图上")
    if under is None or box is None:
        rep.info("I-CONTRAST", loc, "文字下方有图片，对比需看真实渲染")
        return
    if under[1]:
        rep.info("I-CONTRAST", loc, "图与文字之间隔着色块或渐变蒙版，静态量测不适用，对比需看真实渲染")
        return
    path = resolve_asset(asset_root, str(under[0].get("image_id")))
    probe = zone_probe(path, under[0], box, page_bg(page)) if path else None
    if probe is None:
        rep.info("I-CONTRAST", loc, "文字下方有图片，但读不到像素；对比需看真实渲染")
        return
    zone = LIMITS["image"]["zone"]
    limit = LIMITS["contrast"]["large"] if st["size"] >= LIMITS["contrast"]["large_pt"] else LIMITS["contrast"]["normal"]
    share, worst = zone_contrast(st["color"], probe, limit)
    if share < zone["min_ok_share"]:
        rep.soft("W-CONTRAST", loc, f"{st['color']} 压在图上：仅 {share:.0%} 的落点像素达到 {limit}:1"
                                    f"（最不利 5% 约 {worst:.1f}:1）。字没落在安静区：移动文字、改裁切 crop_focus，或重做图的简报")
    elif probe["gray_std"] > zone["max_gray_std"]:
        rep.soft("W-BUSY", loc, f"文字下方图像纹理偏强（灰度标准差 {probe['gray_std']:.0f} > {zone['max_gray_std']}）："
                                f"对比达标但读感会碎；落字区应连续、低纹理")

def check_image(el, loc, ew, eh, assets, asset_root: Path, rep: Report) -> None:
    image_id = el.get("image_id", "")
    asset = assets.get(image_id)
    if asset is None:
        rep.hard("E-REF", loc, f"image_id 未在 assets 登记：{image_id}")
    else:
        check_asset(loc, image_id, asset, asset_root, rep)
    # 信息职责（info_duty）与资产角色（image_role）是两个维度：只有"证据"这一职责带来源约束。
    duty = el.get("info_duty") or ("evidence" if el.get("image_role") == "evidence" else None)
    if duty == "evidence" and asset is not None and asset.get("generation") is not None:
        rep.hard("E-ASSET", loc, f"{image_id} 是生成图，信息职责不能是 evidence；证据只能是真实素材")
    elif duty is None:
        rep.soft("W-ASSET", loc, "未声明信息职责 info_duty：这张图承担什么信息、删掉损失什么？"
                                "（evidence/structure/identity/atmosphere/mood/material/none）")
    elif duty == "none":
        rep.soft("W-ASSET", loc, "info_duty: none = 纯装饰：确认删掉它不损失信息，或改写 asset_reason 说明它留下的理由")
    path = resolve_asset(asset_root, image_id)
    if path is None:
        rep.hard("E-IMG", loc, f"素材文件不存在或不在素材根内：{asset_root / image_id}")
        return
    try:  # 早拦：损坏、SVG 等不可解码或超大像素的素材，否则要到写包末段才以 E-RUNTIME（退出码 3）失败
        from PIL import Image
        with Image.open(path) as probe_img:
            probe_img.verify()
    except Exception as exc:
        rep.hard("E-IMG", loc, f"素材无法解码（仅支持 PNG/JPEG/WebP，且像素数在安全上限内）：{path.name}（{type(exc).__name__}）")
        return
    ppi, loss = image_probe(path, ew, eh, el.get("crop_focus"))
    if ppi is not None and ppi < LIMITS["image"]["min_ppi"]:
        rep.soft("W-PPI", loc, f"有效 {ppi:.0f} ppi < {LIMITS['image']['min_ppi']}，投影会发虚")
    if loss is not None and loss > LIMITS["image"]["max_crop_loss"]:
        rep.soft("W-CROP", loc, f"最终裁切丢掉了原图 {loss:.0%}，画幅或落位框大概选错了")

def check_path(page: Dict[str, Any], el, loc, boxes, rep: Report) -> None:
    pts = resolve_points(el, boxes)
    raw = el.get("points") or []
    if len(pts) != len(raw):
        rep.hard("E-REF", loc, "有顶点无法解析：被引用的元素不存在，或引用形成环")
        return
    if len({(round(a, 4), round(b, 4)) for a, b in pts}) < 2:
        rep.hard("E-GEOM", loc, "路径长度为零：至少要有两个不同的点")
    if el.get("closed") and el.get("arrow") and el["arrow"] != "none":
        rep.soft("W-SHAPE", loc, "闭合路径的箭头没有意义，会被忽略")
    chains = rotated_ancestors(page)
    own = set(chains.get(str(el.get("id")), ()))
    for point in raw:
        if not isinstance(point, dict):
            continue
        target_id = str(point.get("el"))
        target = boxes.get(target_id)
        if target is not None and (target[2] == 0 or target[3] == 0):
            rep.soft("W-REF", loc, "锚点目标宽度或高度为 0，锚点位置可能不稳定")
        stray = set(chains.get(target_id, ())) - own
        if stray:
            rep.soft("W-ANCHOR", loc,
                     f"锚点指向旋转组 {sorted(stray)} 内的 {target_id}：锚点按未旋转几何解析，"
                     f"连线会错位；把连线放进同一个组，或改用显式坐标")

def check_legend(plan: Dict[str, Any], page: Dict[str, Any], el: Dict[str, Any], loc: str, rep: Report) -> None:
    """自定义图例：条目来自被引用图表的颜色映射，这里只查引用与装载，不查配色（配色由映射保证）。"""
    ref = str(el.get("for_chart") or "")
    target = next((e for e in page.get("elements") or [] if str(e.get("id")) == ref), None)
    if target is None or target.get("type") != "chart":
        rep.hard("E-REF", loc, f"for_chart「{ref}」不是本页顶层的 chart 元素")
        return
    spec = target["chart"]
    if spec.get("kind") == "table":
        rep.hard("E-CHART", loc, "表格没有图例；用表头与细线分组表达结构")
        return
    if chart_flags(spec)["legend"]:
        rep.soft("W-CHART", loc, f"{ref} 已开原生图例，再放自定义图例会重复表达同一批类别")
    items = chart_items(spec)
    if spec.get("kind") in ("pie", "donut") and not items:
        rep.hard("E-CHART", loc, "被引用的饼/环没有类别，图例无从生成")
    box = explicit_box(el)
    if box and items:
        vertical = str(el.get("direction") or "vertical") != "horizontal"
        pitch = (box[3] / len(items)) if vertical else (box[2] / len(items))
        need = (float(el.get("font_size") or 18) * 1.45 + 4) / 72.0
        if pitch < need:
            rep.soft("W-FIT", loc, f"图例每{'行' if vertical else '列'} {pitch:.2f} in < 文字所需 {need:.2f} in"
                                  f"（{len(items)} 条）：加高/加宽图例框，或减少条目")

def check_chart(loc: str, el: Dict[str, Any], plan: Dict[str, Any], page: Dict[str, Any], rep: Report) -> None:
    chart = el["chart"]
    kind = chart["kind"]
    data = chart["data"]
    if kind == "table":
        if "headers" not in data:
            rep.hard("E-CHART", loc, "table 的 data 需为 headers/rows")
            return
        headers = data["headers"]
        weights = chart.get("column_weights")
        if weights and len(weights) != len(headers):
            rep.hard("E-CHART", loc, f"column_weights 数量 {len(weights)} ≠ 列数 {len(headers)}")
        aligns = chart.get("column_align")
        if aligns and len(aligns) != len(headers):
            rep.hard("E-CHART", loc, f"column_align 数量 {len(aligns)} ≠ 列数 {len(headers)}")
        for r, row in enumerate(data["rows"]):
            if len(row) != len(headers):
                rep.hard("E-CHART", loc, f"第 {r + 1} 行 {len(row)} 列 ≠ 表头 {len(headers)} 列")
        if chart.get("title"):
            rep.soft("W-CHART", loc, "表格不画图题；用页面 title 或 caption 承担结论")
        return

    series = data.get("series") or []
    # 数据契约（重名、非有限、值数、饼/环、百分比堆叠）与既有图表改数共用 model.chart_data_problems
    for severity, code, message in chart_data_problems(
            kind, data.get("categories"), series, str(chart.get("unit") or ""),
            chart.get("stacked"), chart_flags(chart)["data_labels"]):
        (rep.hard if severity == "HARD" else rep.soft)(code, loc, message)
    if kind == "scatter":
        if "points" not in (series[0] if series else {}):
            rep.hard("E-CHART", loc, "scatter 的 data 需为 series[].points=[[x,y],…]")
            return
        if len(series) > 6:
            rep.soft("W-CHART", loc, f"散点有 {len(series)} 个系列，靠颜色区分会失效；考虑合并或标注重点")
        if not (chart.get("axis_titles") or {}):
            rep.soft("W-CHART", loc, "散点没有轴名：两个轴各是什么量、什么单位，读者无从判断")
    else:
        if "categories" not in data:
            rep.hard("E-CHART", loc, f"{kind} 的 data 需为 categories/series")
            return
        if kind in ("pie", "donut"):
            # 读数闭环用有效开关（model.chart_flags）判断，与编译器、核验器同一份缺省；
            # legend 元素（for_chart 指向本图）同样算类别来源。
            has_legend_el = any(e.get("type") == "legend" and str(e.get("for_chart")) == str(el.get("id"))
                                for e in all_elements(page))
            for severity, message in chart_readout_problems(chart, has_legend_el):
                (rep.hard if severity == "HARD" else rep.soft)("E-CHART" if severity == "HARD" else "W-CHART", loc, message)

    # 单位与期间：不一致就不能共轴。口径不同也不能，必须统一措辞或拆图。
    units = {str(s["unit"]).strip() for s in series if s.get("unit")}
    periods = {str(s["period"]).strip() for s in series if s.get("period")}
    bases = {str(s["basis"]).strip() for s in series if s.get("basis")}
    if len(units) > 1:
        rep.hard("E-CHART", loc, f"多系列单位不一致 {sorted(units)}；同一张图只画同一单位，或换算后改 unit")
    if len(periods) > 1:
        rep.hard("E-CHART", loc, f"多系列期间不一致 {sorted(periods)}；不能在同一坐标轴上比较")
    if len(bases) > 1:
        rep.hard("E-CHART", loc, f"多系列口径不一致 {sorted(bases)}；同一张图只比较同一口径，写清或拆图")
    if chart.get("unit") and units and str(chart["unit"]).strip() not in units:
        rep.hard("E-CHART", loc, f"图表级 unit「{chart['unit']}」与系列单位 {sorted(units)} 矛盾")
    if chart.get("period") and periods and str(chart["period"]).strip() not in periods:
        rep.hard("E-CHART", loc, f"图表级 period「{chart['period']}」与系列期间 {sorted(periods)} 矛盾")
    if chart.get("basis") and bases and str(chart["basis"]).strip() not in bases:
        rep.hard("E-CHART", loc, f"图表级 basis「{chart['basis']}」与系列口径 {sorted(bases)} 矛盾")
    if kind != "scatter":
        if not chart.get("unit") and not units and chart.get("stacked") != "percent":
            rep.soft("W-CHART", loc, "没有声明单位；读者无法判断数值量纲")
        if not chart.get("period") and not periods:
            rep.soft("W-CHART", loc, "没有声明期间；数值图的时效无从判断")
    if not chart.get("basis") and not bases:
        rep.soft("W-CHART", loc, "没有声明口径（分母/样本/统计方式）；读者无法核对一致性")

    stacked = chart.get("stacked")
    if stacked:
        if kind not in ("bar", "area"):
            rep.hard("E-CHART", loc, f"{kind} 不支持 stacked")
        if stacked == "percent" and kind != "bar":
            rep.hard("E-CHART", loc, "百分比堆叠只有 bar 支持")
        if len(series) < 2:
            rep.hard("E-CHART", loc, "堆叠至少需要两个系列")
    if chart.get("horizontal") and kind != "bar":
        rep.hard("E-CHART", loc, "horizontal 只用于 bar")
    combo = chart.get("combo_line_series") or []
    if combo:
        if kind != "bar":
            rep.hard("E-CHART", loc, "combo_line_series 只用于 bar")
        names = [s["name"] for s in series]
        unknown = [n for n in combo if n not in names]
        if unknown:
            rep.hard("E-CHART", loc, f"combo_line_series 里有不存在的系列 {unknown}")
        if len(set(combo)) != len(combo):
            rep.hard("E-CHART", loc, "combo_line_series 有重复系列")
        if len(combo) >= len(names):
            rep.hard("E-CHART", loc, "combo_line_series 不能把全部系列都改成折线")
        if stacked:
            rep.soft("W-CHART", loc, "堆叠柱上的折线读不出量级；确认这是有意为之")
        if not chart.get("combo_line_labels"):
            rep.soft("W-CHART", loc, "折线层没有直标：折线的量只能对着柱的轴目测；要读数就开 combo_line_labels")

    # 颜色映射：与编译器、核验器、自定义图例读同一份 model.chart_color_map。
    items = chart_items(chart)
    colors = chart_color_map(plan, chart)
    base = chart_palette(plan, chart)
    unknown = [k for k in (chart.get("color_map") or {}) if k not in items]
    if unknown:
        rep.hard("E-CHART", loc, f"color_map 里有对不上的{'类别' if kind in ('pie', 'donut') else '系列'} {unknown}："
                                 f"指派会静默失效、颜色退回序列取色，这正是扇区与图例错色的根因")
    if kind in ("pie", "donut"):
        seen: Dict[str, str] = {}
        for name, color in zip(items, colors):
            first = seen.setdefault(color, name)
            if first != name:
                rep.hard("E-CHART", loc, f"「{first}」与「{name}」分到同色 {color}：扇区同色即数据失真，"
                                         f"补足 palette/color_map，或合并类别")
                break
    if len(base) < len(items) and kind not in ("pie", "donut", "table"):
        rep.soft("W-CHART", loc, f"palette 只有 {len(base)} 色，需要 {len(items)}；将循环取色")
    # 灰度自检：系列只靠颜色区分时，转成灰度后明度要拉得开（投影与黑白打印都会丢掉色相）。
    # 全部系列两两比较：相邻拉开不代表隔一个的两条线分得清。
    if kind not in ("pie", "donut", "table") and len(series) >= 2:
        pairs = gray_conflicts(colors)
        if pairs:
            shown = "；".join(f"「{items[i]}」{colors[i]} 与「{items[j]}」{colors[j]}（ΔL* {d:.1f}）" for i, j, d in pairs[:3])
            more = f"，另有 {len(pairs) - 3} 对" if len(pairs) > 3 else ""
            rep.soft("W-GRAY", loc, f"{len(pairs)} 对系列转灰度后明度差 < {LIMITS['chart']['min_series_delta_lstar']}：{shown}{more}。"
                                    "拉开明度，或改为直接标注系列名、不同线型与标记，不要只靠色相")
    if chart.get("axis_labels") is False and not chart_flags(chart)["data_labels"]:
        rep.hard("E-CHART", loc, "axis_labels=false 时必须 data_labels=true，否则无读数")
    hi = chart.get("highlight_category")
    # 承载面：轴、图例、直标的墨色与系列色都相对它取，不是相对页面底色。
    surface = chart_surface(page, el)
    if not chart_surface_known(page, el):
        rep.soft("W-CHART", loc, "图表压在图片或渐变上，承载面无法静态判定，按页面底色取色；"
                                 "换成纯色承载面，或在渲染里逐项看轴、图例与直标")
    body = chart_style(surface)["body"]
    if contrast_ratio(body, surface) < LABEL_INK_MIN:
        rep.soft("W-CHART", loc, f"承载面 {surface} 明度居中，轴/图例墨色最多 {contrast_ratio(body, surface):.2f}:1（< {LABEL_INK_MIN}）；"
                                 "改深或改浅，不要放在中灰上")
    floor_kind = LIMITS["chart"]["min_contrast"].get(kind, 2.0)
    for name, color in zip(items, colors):
        ratio = contrast_ratio(color, surface)
        if ratio < INVISIBLE_RATIO:
            rep.hard("E-CHART", loc, f"「{name}」色 {color} 对承载面 {surface} 仅 {ratio:.2f}:1：几乎不可见；"
                                     "换色、换承载面，或在 color_map 里给这一项单独指派")
        elif ratio < floor_kind:
            rep.soft("W-CHART", loc, f"「{name}」色 {color} 对承载面 {surface} 仅 {ratio:.2f}:1"
                                     f"（{kind} 建议 ≥ {floor_kind}）：这一项在投影上会发虚")
    if hi is not None and chart_flags(chart)["legend"]:
        rep.hard("E-CHART", loc, "highlight_category 的图不能开原生图例：图例色块是系列色，而退隐的柱是灰，两边对不上；"
                                 "强调靠直标或图题说清")
    if kind == "area" and not stacked and len(series) > 1 and chart_flags(chart)["data_labels"]:
        rep.hard("E-CHART", loc, "多系列非堆叠面积相互遮挡，直标落在哪块面上无法保证；改折线、改堆叠，或去掉直标")
    if hi is not None:
        if kind != "bar" or len(series) != 1:
            rep.hard("E-CHART", loc, "highlight_category 只用于单系列柱/条图")
        elif stacked:
            rep.hard("E-CHART", loc, "堆叠柱不能逐类别强调")
        elif hi >= len(data["categories"]):
            rep.hard("E-CHART", loc, f"highlight_category {hi} 超出类别范围")
    axis = chart.get("value_axis") or {}
    if axis.get("min") is not None and axis.get("max") is not None and axis["max"] <= axis["min"]:
        rep.hard("E-CHART", loc, "value_axis 的 max 不大于 min")
    if axis.get("min") is not None and axis["min"] > 0 and not axis.get("note"):
        rep.hard("E-CHART", loc, "数值轴不从零起时必须用 value_axis.note 写明起点，否则读者会把截轴当全量")
    if axis.get("note"):
        rep.soft("W-NOTE", loc,
                 f"value_axis.note「{str(axis['note'])[:40]}」不会画进图里：截轴起点要读者看得见，"
                 f"必须落到页面的 caption、轴名或图题上")

INVISIBLE_RATIO = 1.5   # 系列色对承载面低于此值：肉眼几乎看不见，直接阻断

CAPS_LABEL = re.compile(r"^[A-Z][A-Z0-9 &/·.\-–—]{2,23}$")

def _check_depth(rep: Report, pid: str, page: Dict[str, Any], ids: set) -> None:
    """`depth` 声明前/中/背景各是哪些元素；叠放顺序就是声明顺序，所以意图必须与顺序一致。"""
    depth = page.get("depth")
    if not depth:
        return
    order = [e.get("id") for e in page.get("elements") or []]
    rank = {"background": 0, "midground": 1, "foreground": 2}
    seen: Dict[str, str] = {}
    pos: Dict[str, int] = {}
    for layer in rank:
        for eid in depth.get(layer) or []:
            if eid not in ids:
                rep.hard("E-REF", pid, f"depth.{layer} 里的「{eid}」不是本页元素 id")
            elif eid in seen:
                rep.hard("E-LAYER", f"{pid}/{eid}", f"同一元素同时声明为 {seen[eid]} 与 {layer}")
            else:
                seen[eid] = layer
                if eid in order:
                    pos[eid] = order.index(eid)
    # 只比顶层元素；组内顺序不参与
    ranked = sorted(pos.items(), key=lambda kv: kv[1])
    for (a, _), (b, _) in zip(ranked, ranked[1:]):
        if rank[seen[a]] > rank[seen[b]]:
            rep.hard("E-LAYER", f"{pid}/{a}",
                     f"「{a}」声明为 {seen[a]}，却叠在 {seen[b]} 的「{b}」之下：叠放就是声明顺序，把它写到「{b}」之后")
            break


def _visual_weight(el: Dict[str, Any], box: Tuple[float, float, float, float]) -> float:
    """视觉重量的粗估（平方英寸）：文字取字面“墨量”（字号² × 字数），其余取占面积。
    只用于比较同页元素的尺度反差，不是美学评分。"""
    kind = el.get("type")
    _x, _y, w, h = box
    if kind in ("title", "body", "text", "number", "caption", "footer"):
        pt = float(el.get("font_size") or 0) or 20.0
        chars = len(str(el.get("content") or "").replace("\n", "")) or 1
        return min(pt * pt * chars * 0.9 / 5184.0, w * h)
    if kind in ("image", "chart"):
        return w * h
    if kind == "shape":
        return w * h * (1.0 if el.get("color") else 0.0)
    return 0.0


def _check_design_intent(rep: Report, pid: str, page: Dict[str, Any], ids: set, production: bool) -> None:
    """设计意图的“记录一致性”：判断是否已做、声明与元素是否自相矛盾。不评价构图好不好。
    草稿模式：待判断只给 W-PENDING；生产模式（--production）：待判断即阻断 E-PENDING。"""
    pending = M_pending(page)
    if pending:
        msg = f"设计判断未完成：{'、'.join(pending)}。先由 AI 按真实渲染补 focus 与 composition_rationale（并把 design_state 改为 judged）"
        if production:
            rep.hard("E-PENDING", pid, msg)
        else:
            rep.soft("W-PENDING", pid, msg + "（草稿可继续；--production 会阻断）")
    comp = page.get("composition") or {}
    for ref in comp.get("reading_path") or []:
        if ref not in ids:
            rep.hard("E-REF", pid, f"composition.reading_path 里的「{ref}」不是本页元素 id")
    mode, focus = comp.get("focus_mode"), page.get("focus_element")
    if mode == "none" and focus:
        rep.hard("E-REF", pid, f"composition.focus_mode=none 与 focus_element「{focus}」自相矛盾：二者留一个")
    if mode == "single" and not focus:
        rep.soft("W-FOCUS", pid, "focus_mode=single 但没有 focus_element：写出主角元素 id，或改成 dual / multi / none")
    if focus and focus in ids:
        kind = next((e.get("type") for e in all_elements(page) if e.get("id") == focus), None)
        if kind in ("shape", "divider", "path", "legend"):
            rep.soft("W-FOCUS", f"{pid}/{focus}", f"focus_element 指向装饰性元素（{kind}）：主角通常是文字、数字、图像或图表；确认这是有意的")
    if comp.get("exception") is None and mode in (None, "single") and comp.get("grid") == "none":
        rep.info("I-GRID", pid, "本页声明不用网格（grid=none）但没有写 composition.exception：说明秩序由什么建立")


def M_pending(page: Dict[str, Any]) -> List[str]:
    from model import pending_fields
    return pending_fields(page)


def _check_dominance(rep: Report, pid: str, page: Dict[str, Any], focus: str, boxes: Dict[str, Any]) -> None:
    """focus_element 说它是主角，那尺度上它要真的压得住：与同页最重的非承载元素比较。"""
    items = flatten(page)
    weights: Dict[str, float] = {}
    for it in items:
        el = it["el"]
        eid = el.get("id")
        if el.get("overlap_allowed") and el.get("type") in ("image", "shape"):
            continue   # 承载层（底图、色场）不与主角争主次
        if eid in boxes and eid != "bar":
            weights[eid] = _visual_weight(el, boxes[eid])
    if focus not in weights or len(weights) < 2:
        return
    others = {k: v for k, v in weights.items() if k != focus and v > 0}
    if not others:
        return
    top_id, top_w = max(others.items(), key=lambda kv: kv[1])
    mine = weights[focus]
    if mine < top_w * 1.15:
        rep.soft("W-DOMINANT", f"{pid}/{focus}",
                 f"声明的主角「{focus}」视觉重量（≈{mine:.2f} in²）没有明显超过「{top_id}」（≈{top_w:.2f} in²）："
                 f"放大主角、缩小或灰化配角，或改 focus_element；若本页确为并列，删掉 focus_element")


def page_space(plan: Dict[str, Any], page: Dict[str, Any]) -> Dict[str, Any]:
    """一页的空间事实（只读几何，不评审美）：占位率（24×14 格被元素盒占到的比例）、
    视觉重心（按 _visual_weight 加权，0–1 归一）、主体尺度比（最重元素 ÷ 第二重）。
    整页承载层（满幅图/色场）与发丝线不计占位：它们是环境，不是内容。"""
    w, h = canvas(plan)
    boxes = page_geometry(page)
    cells, wsum, cx, cy, weights = set(), 0.0, 0.0, 0.0, []
    for el in all_elements(page):
        box = boxes.get(str(el.get("id")))
        if box is None or el.get("id") == "_page_number":
            continue
        x, y, bw, bh = box
        if min(bw, bh) <= 0.08 or (el.get("overlap_allowed") and bw * bh >= 0.7 * w * h):
            continue
        for i in range(max(0, int(x / w * 24)), min(24, int(math.ceil((x + bw) / w * 24)))):
            for j in range(max(0, int(y / h * 14)), min(14, int(math.ceil((y + bh) / h * 14)))):
                cells.add((i, j))
        wt = _visual_weight(el, box)
        if wt > 0:
            weights.append(wt)
            wsum += wt
            cx += wt * (x + bw / 2) / w
            cy += wt * (y + bh / 2) / h
    weights.sort(reverse=True)
    return {"density": len(cells) / 336.0,
            "centroid": (cx / wsum, cy / wsum) if wsum else (0.5, 0.5),
            "hero": weights[0] if weights else 0.0,
            "ratio": (weights[0] / weights[1]) if len(weights) > 1 and weights[1] > 0 else 9.9}


def spatial_signals(plan: Dict[str, Any], rep: Report) -> None:
    """整稿空间节奏的事实线索（疏密、重心、气质、留白是否有主体撑住）；只指出该看的页，不下审美结论。"""
    pages = plan.get("pages") or []
    sp = [page_space(plan, p) for p in pages]
    for p, s in zip(pages, sp):
        if s["density"] < 0.22 and s["hero"] < 0.6 and p.get("role") not in ("breathe",):
            rep.soft("W-VOID", str(p.get("id")),
                     f"占位率仅 {s['density']:.0%} 且最大元素的视觉重量 {s['hero']:.2f} in²：大面积空白没有主体撑住。"
                     "放大主体、改成出血或收拢到一个尺度更大的整体；确为停顿页就把 role 设为 breathe")
    if len(pages) < 5:
        return
    dens = [s["density"] for s in sp]
    if max(dens) - min(dens) < 0.12:
        rep.soft("W-RHYTHM", "整稿", f"{len(pages)} 页占位率只在 {min(dens):.0%}–{max(dens):.0%} 之间：整套没有疏密落差，"
                 "选一页极空、一页极密，或确认这是同轴比较")
    def sd(v: List[float]) -> float:
        m = sum(v) / len(v)
        return (sum((x - m) ** 2 for x in v) / len(v)) ** 0.5
    sx, sy = sd([s["centroid"][0] for s in sp]), sd([s["centroid"][1] for s in sp])
    if max(sx, sy) < 0.08:
        rep.soft("W-CENTROID", "整稿", f"视觉重心跨页几乎不动（σx={sx:.2f}，σy={sy:.2f}，画布宽高为 1）：构图重心每页相同，整稿会平；"
                 "用偏心、上下分层、对角或出血让重心换位置")
    atm = [(p.get("composition") or {}).get("atmosphere") for p in pages]
    got = [a for a in atm if a]
    if len(got) >= 5 and len(set(got)) == 1:
        rep.soft("W-ATMOS", "整稿", f"{len(got)} 页声明了同一种空间气质「{got[0]}」：封面、数据页、结尾通常需要不同的空间性格")



def design_signals(plan: Dict[str, Any]) -> Report:
    """整稿层面的模式信号：指出"这一处该由你看一眼"，不是评分，也不下审美结论。

    只报能说出位置与可核对事实的东西：连续同构、标题下装饰线、无功能英文标签、
    素材复用、待主张页、满幅图开场。判为可接受就在交付说明里写理由。
    """
    rep = Report()
    pages = plan.get("pages") or []
    if len(pages) < 2:
        return rep
    spatial_signals(plan, rep)
    w, h = canvas(plan)
    canvas_area = max(w * h, 1e-6)

    vague = [p.get("id") for p in pages if p.get("role") not in ("open", "breathe")
             and not str(p.get("claim") or "").strip()]
    if vague:
        rep.soft("W-CLAIM", "整稿", f"{vague} 没有写 claim：页题是话题还是论断？说不清就是未决")

    prints: List[Tuple] = []
    rule_pages: List[str] = []
    label_pages: List[str] = []
    image_use: Dict[str, List[str]] = {}
    for page in pages:
        boxes = page_geometry(page)
        items = flatten(page)
        sig = []
        for item in items:
            el, box = item["el"], boxes.get(str(item["el"].get("id")))
            if box is None:
                continue
            sig.append((el.get("type"), round(box[0] / 0.5), round(box[1] / 0.5),
                        round(box[2] / 0.5), round(box[3] / 0.5)))
        prints.append(tuple(sorted(sig)))

        titles = [(el, boxes.get(str(el.get("id")))) for el in all_elements(page)
                  if el.get("type") == "title" and boxes.get(str(el.get("id")))]
        for el in all_elements(page):
            box = boxes.get(str(el.get("id")))
            if box is None:
                continue
            thin = el.get("type") == "divider" or (el.get("type") == "shape" and box[3] <= 0.08)
            if thin:
                for _t, tb in titles:
                    if tb is None:
                        continue
                    if (tb[1] + tb[3] - 0.15 <= box[1] <= tb[1] + tb[3] + 0.5
                            and abs(box[0] - tb[0]) <= 0.8 and box[2] <= 3.2):
                        rule_pages.append(str(page.get("id")))
                        break
            if el.get("type") in TEXT_TYPES:
                text = str(el.get("content") or "").strip()
                if CAPS_LABEL.match(text):
                    label_pages.append(str(page.get("id")))
            if el.get("type") == "image":
                image_use.setdefault(str(el.get("image_id")), []).append(str(page.get("id")))

    span = 1
    for i in range(1, len(prints)):
        span = span + 1 if prints[i] == prints[i - 1] else 1
        if span == 3:
            rep.soft("W-MONOTONY", f"{pages[i - 2].get('id')}–{pages[i].get('id')}",
                     "连续三页的元素类型与位置几乎相同：这是同轴比较，还是同一骨架套了三遍？")

    if len(set(rule_pages)) >= max(3, int(len(pages) * 0.6)):
        rep.soft("W-RULE", "整稿", f"{len(set(rule_pages))} 页在标题下同位置放了细线：它在承担分隔，还是逐页复制的装饰？")

    if len(set(label_pages)) >= 3:
        rep.soft("W-LABEL", "整稿", f"{len(set(label_pages))} 页出现全大写英文短标签（如 OVERVIEW）：是导航需要，还是模板气质？")

    reused = {k: v for k, v in image_use.items() if len(v) >= 3}
    if reused:
        detail = "；".join(f"{k} 用于 {len(v)} 页" for k, v in reused.items())
        rep.soft("W-IMG-REUSE", "整稿", f"{detail}：同一素材重复出现，读者会以为是在填充版面")

    first = pages[0]
    first_full = [el for el in all_elements(first) if el.get("type") == "image"]
    for el in first_full:
        box = page_geometry(first).get(str(el.get("id")))
        centered = [t for t in all_elements(first) if t.get("type") == "title" and t.get("align") == "center"]
        if box and box[2] * box[3] >= 0.7 * canvas_area and centered:
            rep.info("I-OPEN", str(first.get("id")), "开场是满幅图＋居中标题：这是默认做法，先确认它不是习惯")
            break
    return rep

def check_asset(loc: str, image_id: str, asset: Dict[str, Any], asset_root: Path, rep: Report) -> None:
    """素材来源与证据边界。唯一事实来源是 generation 记录：有它就是生成图，没有就是真实素材。

    generated 字段只能是对这一事实的声明；两者矛盾（例如 generated: false 却登记了生成记录）直接阻断，
    不允许靠字段组合绕过证据限制。
    """
    if not asset.get("source"):
        rep.hard("E-ASSET", loc, f"{image_id} 缺少 source（来源）")
    gen = asset.get("generation")
    is_generated = gen is not None
    declared = asset.get("generated")
    if declared is not None and bool(declared) != is_generated:
        rep.hard("E-ASSET", loc, f"{image_id} 的 generated={declared} 与 generation "
                                 f"{'有记录' if is_generated else '无记录'}矛盾；生成图必须登记 generation，真实素材不得登记")
    rights = asset.get("rights", "pending")
    if rights == "blocked":
        rep.hard("E-ASSET", loc, f"{image_id} 权利状态 blocked，不得使用")
    elif rights == "pending":
        rep.soft("W-ASSET", loc, f"{image_id} 权利未确认（pending）；交付前确认或改用已清权素材")
    elif rights == "cleared" and not asset.get("reference"):
        rep.soft("W-ASSET", loc, f"{image_id} 声明已清权但没有 reference（授权页/合同/编号）")
    if is_generated and asset.get("used_as_evidence"):
        rep.hard("E-ASSET", loc, f"{image_id} 是生成图，不能作证据")
    if not is_generated:
        if asset.get("used_as_evidence") and rights != "cleared":
            rep.soft("W-ASSET", loc, f"{image_id} 作证据但权利未确认")
        return
    if not isinstance(gen, dict):
        rep.hard("E-ASSET", loc, f"{image_id} 的 generation 必须是对象")
        return
    gbrief = gen.get("brief")
    if not (isinstance(gbrief, dict) and gbrief.get("intent") and gbrief.get("quiet_zone") and gbrief.get("reject")):
        rep.soft("W-ASSET", loc, f"{image_id} 生成记录没有完整简报（brief.intent / quiet_zone / reject）："
                                 f"没有写明图的职责、落字区与淘汰条件，就无法判断这张图是否合格")
    if gen.get("status") != "generated":
        rep.hard("E-ASSET", loc, f"{image_id} 生成记录 status={gen.get('status')}，不得作为正式素材")
    for key in ("tool", "prompt", "file", "sha256"):
        if not gen.get(key):
            rep.hard("E-ASSET", loc, f"{image_id} 生成记录缺少 {key}")
    path = resolve_asset(asset_root, image_id)
    gfile = resolve_asset(asset_root, str(gen.get("file", "")))
    if path is None:
        rep.hard("E-ASSET", loc, f"{image_id} 找不到素材文件")
        return
    if gfile is None or path.resolve() != gfile.resolve():
        rep.hard("E-ASSET", loc, f"{image_id} 与 generation.file={gen.get('file')} 不是同一物理文件")
    elif str(gen.get("sha256", "")).lower() != hashlib.sha256(path.read_bytes()).hexdigest():
        rep.hard("E-ASSET", loc, f"{image_id} 实际 SHA-256 与登记不符")
