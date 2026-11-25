#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""空间构成的渲染量测：把“空白是否有功能、整稿是否机械重复”变成可核对的数字。

输入是真实渲染出的页面 PNG（verify.py 已有），不依赖任何新库。
只量事实（占用率、最大连续空白矩形、几何位置、相邻页签名），不评价好看与否；
是否接受由 composition.temperament / composition.negative_space 的声明决定。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

GRID_W, GRID_H = 96, 54          # 量测网格（16:9 约 0.14 in/格）
INK_DELTA = 40                   # 与页底色的通道差超过它算“有内容”
CELL_MIN = 20                    # 格内内容覆盖 ≥ 约 8% 才算占用：发丝线、页码、脚注线不切碎空白


def page_metrics(png: Path, page_w_in: float, page_h_in: float) -> Dict[str, Any]:
    from PIL import Image, ImageChops

    with Image.open(png) as img:
        rgb = img.convert("RGB")
        small = rgb.resize((GRID_W * 4, GRID_H * 4), Image.BOX)
        bg = max(small.getcolors(maxcolors=GRID_W * GRID_H * 16), key=lambda c: c[0])[1]   # 众数色=页底色
        diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, bg)).convert("L")
        mask = diff.point(lambda v: 255 if v > INK_DELTA else 0)
        cells = mask.resize((GRID_W, GRID_H), Image.BOX).tobytes()
    occ = [[cells[r * GRID_W + c] > CELL_MIN for c in range(GRID_W)] for r in range(GRID_H)]
    total = GRID_W * GRID_H
    ink = sum(sum(row) for row in occ) / total
    area, (c0, r0, cw, rh) = _largest_empty(occ)
    sx, sy = page_w_in / GRID_W, page_h_in / GRID_H
    return {"ink_share": round(ink, 3),
            "layout": _layout(occ),
            "empty_share": round(area / total, 3),
            "empty_box": [round(c0 * sx, 2), round(r0 * sy, 2), round(cw * sx, 2), round(rh * sy, 2)]}


def _layout(occ: List[List[bool]], cols: int = 4, rows: int = 3, floor: float = 0.06) -> str:
    """版面签名：4×3 分区里哪些区有内容（占用 ≥ 6%），行主序的 0/1 串，用于跨页比较构图。"""
    cw, rh = GRID_W // cols, GRID_H // rows
    out = []
    for r in range(rows):
        for c in range(cols):
            n = sum(occ[y][x] for y in range(r * rh, (r + 1) * rh) for x in range(c * cw, (c + 1) * cw))
            out.append("1" if n / (cw * rh) >= floor else "0")
    return "".join(out)


def _largest_empty(occ: List[List[bool]]) -> Tuple[int, Tuple[int, int, int, int]]:
    """最大全空矩形（直方图栈法，O(行×列)）。返回 (格数, (列, 行, 宽, 高))。"""
    heights = [0] * GRID_W
    best, box = 0, (0, 0, 0, 0)
    for r in range(GRID_H):
        for c in range(GRID_W):
            heights[c] = 0 if occ[r][c] else heights[c] + 1
        stack: List[int] = []
        for c in range(GRID_W + 1):
            cur = heights[c] if c < GRID_W else 0
            while stack and heights[stack[-1]] >= cur:
                h = heights[stack.pop()]
                left = stack[-1] + 1 if stack else 0
                width = c - left
                if h * width > best:
                    best, box = h * width, (left, r - h + 1, width, h)
            stack.append(c)
    return best, box


def _gap(a: List[float], b: Tuple[float, float, float, float]) -> float:
    dx = max(b[0] - (a[0] + a[2]), a[0] - (b[0] + b[2]), 0.0)
    dy = max(b[1] - (a[1] + a[3]), a[1] - (b[1] + b[3]), 0.0)
    return (dx * dx + dy * dy) ** 0.5


def zone(box: List[float], w: float, h: float) -> str:
    cx, cy = box[0] + box[2] / 2, box[1] + box[3] / 2
    return f"{min(int(cx / w * 3), 2)}{min(int(cy / h * 3), 2)}"


def check_page(pid: str, page: Dict[str, Any], m: Dict[str, Any], focus_box: Optional[Tuple[float, float, float, float]],
               limits: Dict[str, Any], add) -> None:
    """add(level, code, where, message)。声明了什么就核对什么；没声明的大块空白给 SOFT。"""
    comp = page.get("composition") or {}
    temper = comp.get("temperament")
    neg = comp.get("negative_space") or {}
    env = limits.get("temperament", {}).get(temper or "", {})
    ink, empty = m["ink_share"], m["empty_share"]
    if temper:
        if "min_empty" in env and empty < env["min_empty"]:
            add("SOFT", "V-SPACE", pid, f"声明空间气质「{temper}」，实际最大连续空白仅占页面 {empty:.0%}（要求 ≥ {env['min_empty']:.0%}）："
                "主体周围的空是否真的被让出来了？放大主体、删内容或改声明")
        if "max_empty" in env and empty > env["max_empty"]:
            add("SOFT", "V-SPACE", pid, f"声明空间气质「{temper}」，实际最大连续空白占 {empty:.0%}（要求 ≤ {env['max_empty']:.0%}）：疏密与声明不符")
        if "min_ink" in env and ink < env["min_ink"]:
            add("SOFT", "V-SPACE", pid, f"声明空间气质「{temper}」，实际内容占用 {ink:.0%}（要求 ≥ {env['min_ink']:.0%}）：尺度或出血不足以包围")
    elif (not neg and page.get("role") not in ("open", "breathe", "hold", "close") and ink < limits.get("sparse_ink_share", 0.12)
          and not any(e.get("type") in ("chart", "legend") for e in page.get("elements") or [])):   # 细线图表的墨量天然低，不按占用率判
        add("SOFT", "V-SPACE", pid, f"整页内容占用仅 {ink:.0%}（< {limits.get('sparse_ink_share', 0.12):.0%}）且没有声明空白的功能："
            "主体是否被无理由缩小？放大主体与关键数字，或写 composition.temperament / negative_space 说明这是有意的疏")
    elif not neg and empty > limits.get("dead_zone_share", 0.4) and page.get("role") not in ("open", "breathe", "hold", "close"):
        add("SOFT", "V-SPACE", pid, f"最大连续空白占页面 {empty:.0%}（位置 {m['empty_box']}），页面没有声明它的功能："
            "它在突出主体、分组或停顿吗？是，写 composition.negative_space；否，放大主体或重排")
    if neg.get("role") == "focus" and focus_box:
        gap = _gap(m["empty_box"], focus_box)
        if gap > limits.get("focus_gap_in", 1.0):
            add("SOFT", "V-SPACE", pid, f"negative_space.role=focus，但最大空白离主角 {gap:.1f} in（> {limits.get('focus_gap_in', 1.0)}）：空白没有围着主角")


def rhythm(pages: List[Tuple[str, Dict[str, Any]]], w: float, h: float, limits: Dict[str, Any], add) -> None:
    """连续 N 页版面签名相同或只差一个分区、且占用率同档 = 机械重复。"""
    run = limits.get("rhythm_run", 3)

    def same(a: Dict[str, Any], b: Dict[str, Any]) -> bool:
        diff = sum(x != y for x, y in zip(a["layout"], b["layout"]))
        return diff <= 1 and abs(a["ink_share"] - b["ink_share"]) <= 0.08

    i = 0
    while i < len(pages):
        j = i
        while j + 1 < len(pages) and same(pages[i][1], pages[j + 1][1]):
            j += 1
        if j - i + 1 >= run:
            ids = [p for p, _m in pages[i:j + 1]]
            add("SOFT", "V-RHYTHM", "deck", f"{'、'.join(ids)} 连续 {len(ids)} 页版面分区占用几乎相同（签名 {pages[i][1]['layout']}）："
                "改变其中至少一页的主体尺度、占用或空白位置；确需同一讲法，在 rhythm 里写明理由")
        i = j + 1
