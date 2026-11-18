#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
plan.json → 原生可编辑 PPTX。

用法：
    python scripts/compile.py plan.json out.pptx [--assets DIR] [--report report.json] [--force]

只做确定性执行：校验契约、按 plan 建对象、写出文件。不做审美判断，不自动缩字、
换色、补数据或联网取图。硬错误（E-）在写文件前阻断；目标文件已存在时默认不动它，
确认要替换才加 --force。先写临时文件再原子替换，失败不留半成品；运行期异常另写
<名称>.error.pptx 供排查，该文件不是交付物。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
LIMITS: Dict[str, Any] = json.loads((HERE / "limits.json").read_text(encoding="utf-8-sig"))
SCHEMA: Dict[str, Any] = json.loads((HERE / "plan.schema.json").read_text(encoding="utf-8-sig"))

TEXT_TYPES = ("title", "body", "text", "number", "caption", "footer")
DASH = {"solid": "solid", "dash": "dash", "dot": "sysDot", "dash_dot": "dashDot", "long_dash": "lgDash"}
EMU = 914400.0
PT_EMU = 12700.0

Box = Tuple[float, float, float, float]


# ---------------------------------------------------------------- 诊断


@dataclass
class Issue:
    level: str          # HARD | SOFT | INFO
    code: str
    where: str
    message: str

    def __str__(self) -> str:
        return f"[{self.level}] {self.code} {self.where}: {self.message}"


class Report:
    def __init__(self) -> None:
        self.items: List[Issue] = []

    def add(self, level: str, code: str, where: str, message: str) -> None:
        self.items.append(Issue(level, code, where, message))

    def hard(self, code: str, where: str, message: str) -> None:
        self.add("HARD", code, where, message)

    def soft(self, code: str, where: str, message: str) -> None:
        self.add("SOFT", code, where, message)

    def info(self, code: str, where: str, message: str) -> None:
        self.add("INFO", code, where, message)

    def extend(self, other: "Report") -> None:
        self.items.extend(other.items)

    @property
    def blocked(self) -> bool:
        return any(i.level == "HARD" for i in self.items)

    def count(self, level: str) -> int:
        return sum(1 for i in self.items if i.level == level)

    def to_dict(self) -> List[Dict[str, str]]:
        return [{"level": i.level, "code": i.code, "where": i.where, "message": i.message} for i in self.items]


# ---------------------------------------------------------------- 载入


def load_plan(path: Path, rep: Report) -> Optional[Dict[str, Any]]:
    try:
        raw = path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        rep.hard("E-SCHEMA", str(path), f"无法读取：{exc}")
        return None
    try:
        plan = json.loads(raw)
    except json.JSONDecodeError as exc:
        rep.hard("E-SCHEMA", str(path), f"JSON 语法错误：{exc}")
        return None
    try:
        from jsonschema import Draft202012Validator
    except ImportError:
        rep.hard("E-SCHEMA", "环境", "缺少 jsonschema：pip install -r requirements.txt")
        return plan
    validator = Draft202012Validator(SCHEMA)
    errors = sorted(validator.iter_errors(plan), key=lambda e: list(map(str, e.absolute_path)))
    for err in errors:
        where = "/".join(str(p) for p in err.absolute_path) or "(根)"
        rep.hard("E-SCHEMA", where, err.message[:300])
    return plan if isinstance(plan, dict) else None


# ---------------------------------------------------------------- 度量模型


def canvas(plan: Dict[str, Any]) -> Tuple[float, float]:
    return tuple(LIMITS["canvas"][plan.get("aspect_ratio") or LIMITS["canvas"]["default"]])  # type: ignore


def _em_width(ch: str) -> float:
    if ch in " \u00a0\t":
        return 0.30
    code = ord(ch)
    if code < 128:
        return 0.55 if ch.isalnum() else 0.34
    if ch in "，。、；：？！（）【】《》「」『』“”‘’—…·～":
        return 1.0
    if code >= 0x2E80:
        return 1.0
    return 0.6


def fit_lines(text: str, pt: float, box_w: float) -> int:
    usable = max(box_w - 2 * LIMITS["type"]["inset_in"]["x"], 0.05)
    cap = usable / (pt / 72.0)
    total = 0
    for seg in text.split("\n"):
        width = sum(_em_width(c) for c in seg)
        total += max(1, math.ceil(width / cap - 1e-9))
    return max(1, total)


def fit_height_in(text: str, pt: float, box_w: float, line_spacing: float) -> float:
    pitch = LIMITS["type"]["line_pitch_em"]
    return fit_lines(text, pt, box_w) * pt * pitch * line_spacing / 72.0


def hex_rgb(hex_color: str) -> Tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def rel_luminance(hex_color: str) -> float:
    def channel(v: int) -> float:
        s = v / 255.0
        return s / 12.92 if s <= 0.03928 else ((s + 0.055) / 1.055) ** 2.4
    r, g, b = hex_rgb(hex_color)
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def contrast_ratio(a: str, b: str) -> float:
    la, lb = rel_luminance(a), rel_luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def mix_toward(fg: str, bg: str, t: float) -> str:
    """把 fg 向 bg 混合 t（0=原色，1=bg），只用于图表退隐色。"""
    fa, ba = hex_rgb(fg), hex_rgb(bg)
    out = [round(a + (b - a) * t) for a, b in zip(fa, ba)]
    return "#" + "".join(f"{v:02X}" for v in out)


def is_dark(hex_color: str) -> bool:
    return rel_luminance(hex_color) < LIMITS["contrast"]["dark_background_luminance"]


def page_bg(page: Dict[str, Any]) -> str:
    return (page.get("background") or LIMITS["defaults"]["background"]).upper()


def deck_palette(plan: Dict[str, Any]) -> List[str]:
    return [c.upper() for c in (plan.get("palette") or LIMITS["defaults"]["palette"])]


def page_numbers(plan: Dict[str, Any]) -> Dict[str, Any]:
    cfg = dict(plan.get("page_numbers") or {})
    cfg.setdefault("corner", LIMITS["page_number"]["corner"])
    cfg.setdefault("start", LIMITS["page_number"]["start"])
    cfg.setdefault("hide", [])
    return cfg


def pn_box(plan: Dict[str, Any], page: Dict[str, Any]) -> Dict[str, float]:
    w, h = canvas(plan)
    box = LIMITS["page_number"]["box_in"]
    mx, my = LIMITS["page_number"]["margin_in"]["x"], LIMITS["page_number"]["margin_in"]["y"]
    corner = page_numbers(plan)["corner"]
    return {"x": round(w - mx - box["w"], 4),
            "y": round(my if corner == "top_right" else h - my - box["h"], 4),
            "w": box["w"], "h": box["h"]}


# ---------------------------------------------------------------- 几何


def explicit_box(el: Dict[str, Any]) -> Optional[Box]:
    if all(k in el for k in ("x", "y", "w", "h")):
        return (float(el["x"]), float(el["y"]), float(el["w"]), float(el["h"]))
    return None


def anchor_point(box: Box, anchor: str, at: float = 0.5) -> Tuple[float, float]:
    x, y, w, h = box
    t = min(max(at, 0.0), 1.0)
    if anchor == "left":
        return x, y + h * t
    if anchor == "right":
        return x + w, y + h * t
    if anchor == "top":
        return x + w * t, y
    if anchor == "bottom":
        return x + w * t, y + h
    return x + w / 2.0, y + h / 2.0


def rotated_ancestors(page: Dict[str, Any]) -> Dict[str, Tuple[str, ...]]:
    """每个元素身上带旋转的祖先组，用于判断锚点会不会被组旋转带走。"""
    out: Dict[str, Tuple[str, ...]] = {}

    def walk(items: List[Dict[str, Any]], chain: Tuple[str, ...]) -> None:
        for el in items:
            eid = str(el.get("id"))
            out[eid] = chain
            if el.get("type") == "group":
                nxt = chain + (eid,) if float(el.get("rotation") or 0) else chain
                walk(el.get("elements") or [], nxt)

    walk(page.get("elements") or [], ())
    return out


def flatten(page: Dict[str, Any]) -> List[Dict[str, Any]]:
    """声明顺序展开：group 子元素就地展开并带上 group id，供校验与查背景用。"""
    out: List[Dict[str, Any]] = []

    def walk(items: List[Dict[str, Any]], gid: Optional[str]) -> None:
        for el in items:
            if el.get("type") == "group":
                walk(el.get("elements") or [], el.get("id"))
            else:
                out.append({"el": el, "group": gid})

    walk(page.get("elements") or [], None)
    return out


def _union(boxes: List[Box]) -> Optional[Box]:
    if not boxes:
        return None
    x0 = min(b[0] for b in boxes)
    y0 = min(b[1] for b in boxes)
    x1 = max(b[0] + b[2] for b in boxes)
    y1 = max(b[1] + b[3] for b in boxes)
    return (x0, y0, x1 - x0, y1 - y0)


def all_elements(page: Dict[str, Any]) -> List[Dict[str, Any]]:
    """页内全部元素，含 group 自身与其子元素（id 引用与锚点解析用这一份）。"""
    out: List[Dict[str, Any]] = []

    def walk(items: List[Dict[str, Any]]) -> None:
        for el in items:
            out.append(el)
            if el.get("type") == "group":
                walk(el.get("elements") or [])

    walk(page.get("elements") or [])
    return out


def page_geometry(page: Dict[str, Any]) -> Dict[str, Box]:
    """解析每页全部元素的盒子：显式坐标、path 点集、group 子元素并集。锚点在这里解析。"""
    by_id: Dict[str, Dict[str, Any]] = {str(el.get("id")): el for el in all_elements(page)}
    boxes: Dict[str, Box] = {}

    def resolve(eid: str, stack: Tuple[str, ...]) -> Optional[Box]:
        if eid in boxes:
            return boxes[eid]
        el = by_id.get(eid)
        if el is None or eid in stack:
            return None
        kind = el.get("type")
        box: Optional[Box] = None
        if kind == "path":
            pts = [point_xy(p, resolve, stack + (eid,)) for p in el.get("points") or []]
            if any(p is None for p in pts) or len(pts) < 2:
                box = None
            else:
                xs = [p[0] for p in pts]  # type: ignore[index]
                ys = [p[1] for p in pts]  # type: ignore[index]
                box = (min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
        elif kind == "group":
            kids = [resolve(str(c.get("id")), stack + (eid,)) for c in el.get("elements") or []]
            box = _union([b for b in kids if b is not None])
        else:
            box = explicit_box(el)
        if box is not None:
            boxes[eid] = box
        return box

    for eid in by_id:
        resolve(eid, ())
    return boxes


def point_xy(point: Any, resolve, stack: Tuple[str, ...]) -> Optional[Tuple[float, float]]:
    if isinstance(point, (list, tuple)) and len(point) == 2:
        return float(point[0]), float(point[1])
    if isinstance(point, dict):
        target = resolve(str(point.get("el")), stack)
        if target is None:
            return None
        return anchor_point(target, str(point.get("anchor", "center")), float(point.get("at", 0.5)))
    return None


def resolve_points(el: Dict[str, Any], boxes: Dict[str, Box]) -> List[Tuple[float, float]]:
    out: List[Tuple[float, float]] = []
    for point in el.get("points") or []:
        xy = point_xy(point, lambda eid, _s: boxes.get(eid), ())
        if xy is not None:
            out.append(xy)
    return out


def _intersection(a: Box, b: Box) -> Tuple[float, float]:
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    ix = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    return ix, iy


def substantive_overlap(a: Box, b: Box) -> bool:
    ix, iy = _intersection(a, b)
    if min(ix, iy) < LIMITS["geometry"]["overlap_min_edge_in"]:
        return False
    area = ix * iy
    smaller = min(a[2] * a[3], b[2] * b[3])
    if smaller <= 0:
        return False
    return area >= LIMITS["geometry"]["overlap_area_sq_in"] and area / smaller >= LIMITS["geometry"]["overlap_ratio"]


def resolve_asset(asset_root: Path, image_id: str) -> Optional[Path]:
    direct = asset_root / image_id
    if direct.is_file():
        return direct
    for ext in (".png", ".jpg", ".jpeg", ".webp", ".PNG", ".JPG", ".JPEG", ".WEBP"):
        candidate = asset_root / (image_id + ext)
        if candidate.is_file():
            return candidate
    stem = Path(image_id).stem
    if asset_root.is_dir():
        for f in sorted(asset_root.iterdir()):
            if f.is_file() and f.stem == stem:
                return f
    return None


# ---------------------------------------------------------------- 文本式样


def text_style(plan: Dict[str, Any], page: Dict[str, Any], el: Dict[str, Any]) -> Dict[str, Any]:
    kind = el["type"]
    size = float(el.get("font_size") or LIMITS["type"]["default_pt"].get(kind, 18))
    fonts = plan.get("fonts") or {}
    family = el.get("font") or fonts.get("heading" if kind in ("title", "number") else "body") \
        or LIMITS["defaults"]["fonts"]["heading" if kind in ("title", "number") else "body"]
    color = el.get("color")
    if not color:
        color = LIMITS["defaults"]["ink_on_dark"] if is_dark(page_bg(page)) else LIMITS["defaults"]["ink"]
    return {
        "size": size,
        "family": family,
        "color": (color or LIMITS["defaults"]["ink"]).upper(),
        "bold": bool(el.get("bold", LIMITS["type"]["bold"].get(kind, False))),
        "line_spacing": float(el.get("line_spacing") or LIMITS["type"]["line_spacing"].get(kind, 1.25)),
        "align": el.get("align", "left"),
        "valign": el.get("valign", "top"),
        "spacing": float(el.get("letter_spacing_pt") or 0),
    }


# ---------------------------------------------------------------- 校验


CARRIERS = ("shape", "divider", "path")


def validate(plan: Dict[str, Any], asset_root: Path) -> Report:
    rep = Report()
    w, h = canvas(plan)
    tol = LIMITS["canvas"]["bound_tolerance_in"]
    pages = plan.get("pages") or []
    assets = plan.get("assets") or {}
    palette = deck_palette(plan)
    pn_cfg = page_numbers(plan)

    rhythm = plan.get("rhythm")
    if rhythm and len(rhythm) != len(pages):
        rep.soft("W-RHYTHM", "整稿", f"rhythm 有 {len(rhythm)} 条，页数 {len(pages)}；按页一一对应")
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
        ids: set = set()
        for el in all_elements(page):
            eid = el.get("id", "?")
            if eid in ids:
                rep.hard("E-ID", f"{pid}/{eid}", f"元素 id 在本页重复（含 group 子元素）：{eid}")
            ids.add(eid)

        focus = page.get("focus_element")
        if focus and focus not in ids:
            rep.hard("E-REF", pid, f"focus_element「{focus}」不是本页元素 id")

        boxes = page_geometry(page)

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
            if x + ew <= 0 or y + eh <= 0 or x >= w or y >= h:
                rep.hard("E-GEOM", eloc, f"元素完全在画布外：x={x} y={y} w={ew} h={eh}")
            elif kind in TEXT_TYPES + ("chart",) and (x < -tol or y < -tol or x + ew > w + tol or y + eh > h + tol):
                what = "文本框" if kind in TEXT_TYPES else "图表"
                rep.soft("E-GEOM", eloc, f"{what}伸出画布（{w}×{h}），边缘内容会被裁掉")

            if el.get("overlap_allowed") and kind not in CARRIERS and not (
                    kind == "image" and el.get("image_role") in ("background", "texture")):
                rep.hard("E-OVERLAP", eloc,
                         "overlap_allowed 只能声明在承载层（背景图/蒙版/色块/闭合路径）上；被压的元素不能自我豁免")
            if kind == "divider" and eh > 0.2:
                rep.soft("E-GEOM", eloc, f"divider 高 {eh:.2f}in，已经是一块色块；用 shape 更清楚")
            if kind == "group":
                kids = el.get("elements") or []
                if not kids:
                    rep.hard("E-GEOM", eloc, "group 没有子元素")
                declared = explicit_box(el)
                computed = _union([boxes[str(c.get("id"))] for c in kids if str(c.get("id")) in boxes] or [])
                if declared and computed and any(abs(a - b) > 0.05 for a, b in zip(declared, computed)):
                    rep.info("I-GROUP", eloc, "group 的几何由子元素推出，声明的 x/y/w/h 未采用")

            if kind in TEXT_TYPES:
                check_text(plan, page, el, eloc, ew, eh, rep)
            elif kind == "image":
                check_image(el, eloc, ew, eh, assets, asset_root, rep)
            elif kind == "shape" and not (el.get("color") or el.get("gradient") or el.get("stroke")):
                rep.hard("E-SHAPE", eloc, "shape 既无填充也无描边，成品里看不见")
            elif kind == "path":
                check_path(page, el, eloc, boxes, rep)
            elif kind == "chart":
                check_chart(eloc, el, palette, rep)

        if pid not in pn_cfg["hide"]:
            box = pn_box(plan, page)
            for item in flatten(page):
                el = item["el"]
                if el.get("overlap_allowed"):
                    continue  # 承载层在下，页码本来就压在它上面
                gb = boxes.get(el.get("id"))
                if gb is None:
                    continue
                ix, iy = _intersection(gb, (box["x"], box["y"], box["w"], box["h"]))
                if ix <= 0.02 or iy <= 0.02:
                    continue
                where = f"{pid}/{el.get('id')}"
                if el["type"] in TEXT_TYPES or el["type"] == "chart":
                    rep.hard("E-PAGE", where, "文字或图表与页码保留框相交；移开它，或整稿改用另一角区")
                elif el["type"] == "image":
                    rep.info("I-PAGE", where, "图片压到页码角区：用 page_number_tone 声明实际明暗，并确认页码可读")
                elif el["type"] != "path":
                    rep.soft("E-PAGE", where, "色块压到页码角区：确认页码在它上面可读，必要时换角区")

    # 跨元素：实质重叠（叠放顺序即声明顺序）。同一 group 内部的叠放是设计许可的，不查。
    for page in pages:
        pid = page.get("id")
        items = flatten(page)
        boxes = page_geometry(page)
        for i, lower in enumerate(items):
            lbox = boxes.get(lower["el"].get("id"))
            if lbox is None or lower["el"].get("overlap_allowed") or lower["el"]["type"] == "path":
                continue
            for upper in items[i + 1:]:
                if upper["group"] is not None and upper["group"] == lower["group"]:
                    continue
                ubox = boxes.get(upper["el"].get("id"))
                if ubox is None:
                    continue
                if substantive_overlap(lbox, ubox):
                    rep.hard("E-OVERLAP", f"{pid}/{upper['el'].get('id')}",
                             f"与先声明的 {lower['el'].get('id')} 实质重叠；若后者是承载层（背景图/蒙版/色块），"
                             f"在 {lower['el'].get('id')} 上声明 overlap_allowed，否则移开")
    return rep


def check_text(plan, page, el, loc, ew, eh, rep: Report) -> None:
    content = el.get("content", "")
    if not str(content).strip():
        rep.hard("E-TEXT", loc, "文本内容为空；删掉这个元素，或写入内容")
    runs = el.get("runs")
    if runs:
        if "\n" in content:
            rep.hard("E-REF", loc, "runs 与 content 并存时 content 不能含换行")
        if "".join(r.get("text", "") for r in runs) != content:
            rep.hard("E-REF", loc, "runs 拼接与 content 不一致")
    st = text_style(plan, page, el)
    need = fit_height_in(str(content), st["size"], ew, st["line_spacing"])
    if need > eh * LIMITS["type"]["fit_tolerance"]:
        rep.soft("W-FIT", loc, f"按 {st['size']:g}pt 估算需 {need:.2f}in，框高 {eh:.2f}in：先编辑文案，再加宽/拆页")
    minimum = LIMITS["type"]["min_pt"].get(el["type"], 0)
    if st["size"] < minimum:
        rep.soft("W-PT", loc, f"{el['type']} {st['size']:g}pt 低于 {minimum}pt 下限")
    bg = resolve_bg(page, el)
    if bg is None:
        rep.info("I-CONTRAST", loc, "文字下方有图片，对比需看真实渲染")
    else:
        limit = LIMITS["contrast"]["large"] if st["size"] >= LIMITS["contrast"]["large_pt"] \
            else LIMITS["contrast"]["normal"]
        ratio = contrast_ratio(st["color"], bg)
        if ratio < limit:
            rep.soft("W-CONTRAST", loc, f"{st['color']} on {bg} = {ratio:.2f}:1，低于 {limit}:1")


def check_image(el, loc, ew, eh, assets, asset_root: Path, rep: Report) -> None:
    image_id = el.get("image_id", "")
    asset = assets.get(image_id)
    if asset is None:
        rep.hard("E-REF", loc, f"image_id 未在 assets 登记：{image_id}")
    else:
        check_asset(loc, image_id, asset, asset_root, rep)
    path = resolve_asset(asset_root, image_id)
    if path is None:
        rep.hard("E-IMG", loc, f"素材文件不存在：{asset_root / image_id}")
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
        rep.soft("E-SHAPE", loc, "闭合路径的箭头没有意义，会被忽略")
    chains = rotated_ancestors(page)
    own = set(chains.get(str(el.get("id")), ()))
    for point in raw:
        if not isinstance(point, dict):
            continue
        target_id = str(point.get("el"))
        target = boxes.get(target_id)
        if target is not None and (target[2] == 0 or target[3] == 0):
            rep.soft("E-REF", loc, "锚点目标宽度或高度为 0，锚点位置可能不稳定")
        stray = set(chains.get(target_id, ())) - own
        if stray:
            rep.soft("W-ANCHOR", loc,
                     f"锚点指向旋转组 {sorted(stray)} 内的 {target_id}：锚点按未旋转几何解析，"
                     f"连线会错位；把连线放进同一个组，或改用显式坐标")


def check_chart(loc: str, el: Dict[str, Any], palette: List[str], rep: Report) -> None:
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
            rep.soft("E-CHART", loc, "表格不画图题；用页面 title 或 caption 承担结论")
        return

    series = data.get("series") or []
    if kind == "scatter":
        if "points" not in (series[0] if series else {}):
            rep.hard("E-CHART", loc, "scatter 的 data 需为 series[].points=[[x,y],…]")
            return
        if len(series) > 6:
            rep.soft("E-CHART", loc, f"散点有 {len(series)} 个系列，靠颜色区分会失效；考虑合并或标注重点")
        if not (chart.get("axis_titles") or {}):
            rep.soft("W-CHART", loc, "散点没有轴名：两个轴各是什么量、什么单位，读者无从判断")
    else:
        if "categories" not in data:
            rep.hard("E-CHART", loc, f"{kind} 的 data 需为 categories/series")
            return
        cats = data["categories"]
        for s in series:
            if len(s["values"]) != len(cats):
                rep.hard("E-CHART", loc, f"系列「{s['name']}」有 {len(s['values'])} 个值，类别 {len(cats)} 个")
        if kind in ("pie", "donut"):
            if len(series) != 1:
                rep.hard("E-CHART", loc, f"{kind} 只能单系列，现有 {len(series)} 个")
            for s in series:
                if any(v < 0 for v in s["values"]):
                    rep.hard("E-CHART", loc, "饼/环不能有负值")

    # 单位与期间：不一致就不能共轴。shared_basis 只确认口径，不豁免单位与期间。
    units = {str(s["unit"]).strip() for s in series if s.get("unit")}
    periods = {str(s["period"]).strip() for s in series if s.get("period")}
    bases = {str(s["basis"]).strip() for s in series if s.get("basis")}
    if len(units) > 1:
        rep.hard("E-CHART", loc, f"多系列单位不一致 {sorted(units)}；同一张图只画同一单位，或换算后改 unit")
    if len(periods) > 1:
        rep.hard("E-CHART", loc, f"多系列期间不一致 {sorted(periods)}；不能在同一坐标轴上比较")
    if len(bases) > 1 and not chart.get("shared_basis"):
        rep.hard("E-CHART", loc, f"多系列口径不一致 {sorted(bases)}；写清统一口径，或拆图")
    if chart.get("shared_basis") and not bases and not chart.get("basis"):
        rep.soft("E-CHART", loc, "shared_basis 没有任何口径文字支撑，等于没有说清")
    if chart.get("unit") and units and str(chart["unit"]).strip() not in units:
        rep.hard("E-CHART", loc, f"图表级 unit「{chart['unit']}」与系列单位 {sorted(units)} 矛盾")
    if chart.get("period") and periods and str(chart["period"]).strip() not in periods:
        rep.hard("E-CHART", loc, f"图表级 period「{chart['period']}」与系列期间 {sorted(periods)} 矛盾")
    if chart.get("basis") and bases and str(chart["basis"]).strip() not in bases:
        rep.hard("E-CHART", loc, f"图表级 basis「{chart['basis']}」与系列口径 {sorted(bases)} 矛盾")
    if kind != "scatter":
        if not chart.get("unit") and not units and chart.get("stacked") != "percent":
            rep.soft("E-CHART", loc, "没有声明单位；读者无法判断数值量纲")
        if not chart.get("period") and not periods:
            rep.soft("E-CHART", loc, "没有声明期间；数值图的时效无从判断")
    if not chart.get("basis") and not bases:
        rep.soft("E-CHART", loc, "没有声明口径（分母/样本/统计方式）；读者无法核对一致性")

    stacked = chart.get("stacked")
    if stacked:
        if kind not in ("bar", "area"):
            rep.hard("E-CHART", loc, f"{kind} 不支持 stacked")
        if stacked == "percent" and kind != "bar":
            rep.hard("E-CHART", loc, "百分比堆叠只有 bar 支持")
        if len(series) < 2:
            rep.hard("E-CHART", loc, "堆叠至少需要两个系列")
        if stacked == "percent" and kind == "bar":
            bad = []
            for j in range(len(data["categories"])):
                total = sum(float(s["values"][j]) for s in series)
                if abs(total - 100) > 0.5:
                    bad.append(f"{data['categories'][j]}={total:g}")
            if bad:
                level = "硬" if chart.get("data_labels") else "提示"
                reporter = rep.hard if chart.get("data_labels") else rep.soft
                reporter("E-CHART", loc,
                         f"百分比堆叠要求每列合计 100，实际 {bad}；图形会归一化而直标仍是原值（{level}）")
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
            rep.soft("E-CHART", loc, "堆叠柱上的折线读不出量级；确认这是有意为之")

    need = len(data.get("categories") or []) if kind in ("pie", "donut") else len(series)
    if len(palette) < need:
        rep.soft("E-CHART", loc, f"palette 只有 {len(palette)} 色，需要 {need}；将循环取色")
    if chart.get("axis_labels") is False and not chart.get("data_labels"):
        rep.hard("E-CHART", loc, "axis_labels=false 时必须 data_labels=true，否则无读数")
    hi = chart.get("highlight_category")
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


def resolve_bg(page: Dict[str, Any], target: Dict[str, Any]) -> Optional[str]:
    """文字实际落在什么颜色上：页面底色 + 其下覆盖 ≥90% 的色块；其下有图则无法静态判定。"""
    base = page_bg(page)
    box = explicit_box(target)
    if box is None:
        return base
    area = max(box[2] * box[3], 1e-6)
    for item in flatten(page):
        el = item["el"]
        if el is target:
            break
        below = explicit_box(el)
        if below is None or el.get("type") not in ("shape", "divider", "image"):
            continue
        ix, iy = _intersection(below, box)
        cover = ix * iy / area
        if el["type"] == "image" and cover >= 0.6:
            return None
        if el["type"] not in ("shape", "divider") or cover < 0.9:
            continue
        if el.get("gradient"):
            return None
        alpha = float(el.get("fill_alpha", 1.0))
        color = str(el.get("color") or LIMITS["defaults"]["background"]).upper()
        base = color if alpha >= 0.999 else mix_toward(base, color, alpha)
    return base


def image_probe(path: Path, w: float, h: float, focus: Optional[Dict[str, float]]):
    """返回（有效 ppi, 被裁掉的原图比例）。裁掉太多说明画幅选错，不是"裁切很紧"。"""
    try:
        from PIL import Image
        with Image.open(path) as img:
            px_w, px_h = img.size
    except Exception:
        return None, None
    if w <= 0 or h <= 0 or not px_w or not px_h:
        return None, None
    left, right, top, bottom = _crop_geometry(px_w, px_h, w, h, focus)
    loss = 1 - (1 - left - right) * (1 - top - bottom)
    return min(px_w / w, px_h / h), loss


def check_asset(loc: str, image_id: str, asset: Dict[str, Any], asset_root: Path, rep: Report) -> None:
    rights = asset.get("rights", "pending")
    generated = bool(asset.get("generated"))
    if rights == "blocked":
        rep.hard("E-ASSET", loc, f"{image_id} 权利状态 blocked，不得使用")
    elif rights == "pending":
        rep.soft("E-ASSET", loc, f"{image_id} 权利未确认（pending）；交付前确认或改用已清权素材")
    elif rights == "cleared" and not asset.get("reference"):
        rep.soft("E-ASSET", loc, f"{image_id} 声明已清权但没有 reference（授权页/合同/编号）")
    if generated and asset.get("used_as_evidence"):
        rep.hard("E-ASSET", loc, f"{image_id} 是生成图，不能作证据")
    gen = asset.get("generation")
    if not gen:
        return
    if gen.get("status") != "generated":
        rep.hard("E-ASSET", loc, f"{image_id} 生成记录 status={gen.get('status')}，不得作为正式素材")
    path = resolve_asset(asset_root, image_id)
    gfile = resolve_asset(asset_root, gen.get("file", ""))
    if path and gfile and path.resolve() != gfile.resolve():
        rep.hard("E-ASSET", loc, f"{image_id} 与 generation.file={gen.get('file')} 不是同一物理文件")
    if path and not gen.get("sha256"):
        rep.soft("E-ASSET", loc, f"{image_id} 是生成图但没有登记 sha256，无法证明用的是哪一次结果")
    if path and gen.get("sha256"):
        actual = hashlib.sha256(path.read_bytes()).hexdigest()
        if actual.lower() != str(gen["sha256"]).lower():
            rep.hard("E-ASSET", loc, f"{image_id} 实际 SHA-256 与登记不符")


# ---------------------------------------------------------------- 绘图底件


def _strip_style(shape) -> None:
    """删掉 p:style：主题的 effectRef 会带来投影，LibreOffice 与 PowerPoint 表现还不一致。
    形状的填充与描边全部由 plan 显式给出，不需要主题引用。"""
    from pptx.oxml.ns import qn
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)


def _alpha_on(color_node, alpha: float) -> None:
    from pptx.oxml.ns import qn
    from lxml import etree
    for old in color_node.findall(qn("a:alpha")):
        color_node.remove(old)
    etree.SubElement(color_node, qn("a:alpha")).set("val", str(int(round(alpha * 100000))))


def _apply_fill_alpha(shape, alpha: float) -> None:
    from pptx.oxml.ns import qn
    solid = shape.fill._xPr.find(qn("a:solidFill"))
    if solid is None:
        return
    _alpha_on(solid[0], alpha)


def _apply_gradient(shape, grad: Dict[str, Any]) -> None:
    """两停线性渐变：两端各自可声明透明度。用于压图蒙版、光带与大面积色场过渡。"""
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    from lxml import etree

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


def _apply_stroke(shape, stroke: Optional[Dict[str, Any]], arrow: str = "none") -> None:
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    from lxml import etree

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
    from pptx.oxml.ns import qn
    from lxml import etree
    font.name = family
    rPr = getattr(font, "_rPr", None)
    if rPr is None:
        return
    for tag in ("a:ea", "a:cs"):
        node = rPr.find(qn(tag))
        if node is None:
            node = etree.SubElement(rPr, qn(tag))
        node.set("typeface", family)


def _set_run_font(run, family: str, size: float, color: str, bold: bool, spacing: float = 0.0) -> None:
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    from lxml import etree
    font = run.font
    font.size = Pt(size)
    font.bold = bold
    font.name = family
    font.color.rgb = RGBColor.from_string(color.lstrip("#").upper())
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        node = rPr.find(qn(tag))
        if node is None:
            node = etree.SubElement(rPr, qn(tag))
        node.set("typeface", family)
    if spacing:
        rPr.set("spc", str(int(round(spacing * 100))))


# ---------------------------------------------------------------- 建对象


def add_text(slide, plan, page, el):
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

    st = text_style(plan, page, el)
    box = explicit_box(el)
    frame = slide.shapes.add_textbox(Inches(box[0]), Inches(box[1]), Inches(box[2]), Inches(box[3]))
    frame.name = f"{page['id']}.{el['id']}"
    tf = frame.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    inset = LIMITS["type"]["inset_in"]
    tf.margin_left = tf.margin_right = Inches(inset["x"])
    tf.margin_top = tf.margin_bottom = Inches(inset["y"])
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[st["valign"]]
    align = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}[st["align"]]

    paragraphs = el["content"].split("\n")
    runs_override = el.get("runs")
    for idx, seg in enumerate(paragraphs):
        para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = st["line_spacing"]
        if el.get("space_after_pt"):
            para.space_after = Pt(float(el["space_after_pt"]))
        pieces = runs_override if (runs_override and idx == 0) else [{"text": seg}]
        for piece in pieces:
            run = para.add_run()
            run.text = piece.get("text", "")
            _set_run_font(
                run,
                piece.get("font") or st["family"],
                float(piece.get("font_size") or st["size"]),
                str(piece.get("color") or st["color"]).upper(),
                bool(piece.get("bold", st["bold"])),
                float(piece.get("letter_spacing_pt", st["spacing"])),
            )
    return frame


def add_shape(slide, page, el):
    from pptx.util import Inches
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor

    kinds = {"rect": MSO_SHAPE.RECTANGLE, "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE, "oval": MSO_SHAPE.OVAL,
             "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE, "diamond": MSO_SHAPE.DIAMOND, "chevron": MSO_SHAPE.CHEVRON,
             "arrow_right": MSO_SHAPE.RIGHT_ARROW, "arrow_up": MSO_SHAPE.UP_ARROW}
    box = explicit_box(el)
    shape = slide.shapes.add_shape(kinds[el.get("shape_kind") or "rect"],
                                   Inches(box[0]), Inches(box[1]), Inches(box[2]), Inches(box[3]))
    shape.name = f"{page['id']}.{el['id']}"
    _strip_style(shape)
    shape.shadow.inherit = False
    if el.get("gradient"):
        _apply_gradient(shape, el["gradient"])
    elif el.get("color"):
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(str(el["color"]).lstrip("#").upper())
        if el.get("fill_alpha") and float(el["fill_alpha"]) < 1:
            _apply_fill_alpha(shape, float(el["fill_alpha"]))
    else:
        shape.fill.background()
    _apply_stroke(shape, el.get("stroke"))
    if el.get("rotation"):
        shape.rotation = float(el["rotation"])
    if shape.has_text_frame:
        shape.text_frame.text = ""
    return shape


def add_path(slide, page, el, boxes):
    from pptx.util import Emu, Pt
    from pptx.enum.shapes import MSO_CONNECTOR
    from pptx.dml.color import RGBColor

    pts = resolve_points(el, boxes)
    name = f"{page['id']}.{el['id']}"
    if len(pts) == 2:
        shape = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                           Emu(int(round(pts[0][0] * EMU))), Emu(int(round(pts[0][1] * EMU))),
                                           Emu(int(round(pts[1][0] * EMU))), Emu(int(round(pts[1][1] * EMU))))
    else:
        builder = slide.shapes.build_freeform(Emu(int(round(pts[0][0] * EMU))), Emu(int(round(pts[0][1] * EMU))), scale=1.0)
        builder.add_line_segments([(Emu(int(round(x * EMU))), Emu(int(round(y * EMU)))) for x, y in pts[1:]],
                                  close=bool(el.get("closed")))
        shape = builder.convert_to_shape()
    shape.name = name
    _strip_style(shape)
    shape.shadow.inherit = False
    has_fill = hasattr(shape, "fill")
    if el.get("closed") and el.get("color") and has_fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(str(el["color"]).lstrip("#").upper())
        if el.get("fill_alpha") and float(el["fill_alpha"]) < 1:
            _apply_fill_alpha(shape, float(el["fill_alpha"]))
    elif has_fill:
        shape.fill.background()
    _apply_stroke(shape, el.get("stroke"), str(el.get("arrow") or "none"))
    return shape


def add_image(slide, page, el, asset_root: Path):
    from pptx.util import Inches
    from pptx.oxml.ns import qn
    from lxml import etree
    from PIL import Image

    path = resolve_asset(asset_root, el["image_id"])
    if path is None:
        raise FileNotFoundError(f"素材缺失 {el['image_id']}")
    with Image.open(path) as img:
        src_w, src_h = img.size
    box = explicit_box(el)
    pic = slide.shapes.add_picture(str(path), Inches(box[0]), Inches(box[1]), Inches(box[2]), Inches(box[3]))
    pic.name = f"{page['id']}.{el['id']}"
    l, r, t, b = _crop_geometry(src_w, src_h, box[2], box[3], el.get("crop_focus"))
    if any(v > 0 for v in (l, r, t, b)):
        pic.crop_left, pic.crop_right, pic.crop_top, pic.crop_bottom = l, r, t, b
    if el.get("alpha") and float(el["alpha"]) < 1:
        blip = pic._element.blipFill.find(qn("a:blip"))
        if blip is not None:
            etree.SubElement(blip, qn("a:alphaModFix")).set("amt", str(int(round(float(el["alpha"]) * 100000))))
    return pic


def _crop_geometry(src_w: float, src_h: float, w: float, h: float, focus: Optional[Dict[str, float]]):
    fx = 0.5 if not focus else float(focus.get("x", 0.5))
    fy = 0.5 if not focus else float(focus.get("y", 0.5))
    src_ratio, dst_ratio = src_w / src_h, w / h
    l = r = t = b = 0.0
    if src_ratio > dst_ratio:
        trim = 1 - (dst_ratio / src_ratio)
        l, r = trim * fx, trim * (1 - fx)
    elif src_ratio < dst_ratio:
        trim = 1 - (src_ratio / dst_ratio)
        t, b = trim * fy, trim * (1 - fy)
    return l, r, t, b


def _series_colors(palette: List[str], count: int) -> List[str]:
    if not palette:
        palette = LIMITS["defaults"]["palette"]
    return [palette[i % len(palette)] for i in range(count)]


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
    from pptx.oxml.ns import qn
    from lxml import etree
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
    from pptx.oxml.ns import qn
    from lxml import etree
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


def add_chart(slide, plan, page, el, palette: List[str]):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.chart.data import CategoryChartData, XyChartData
    from pptx.enum.chart import XL_LABEL_POSITION

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
    colors = _series_colors(palette, len(data.categories) if kind in ("pie", "donut") else len(series))
    plot = obj.plots[0]
    combo = set(chart.get("combo_line_series") or [])
    stacked = bool(chart.get("stacked"))

    if kind in ("pie", "donut"):
        obj.has_legend = True
        _style_legend(obj, style, family)
        for j, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = RGBColor.from_string(colors[j % len(colors)].lstrip("#"))
            point.format.line.color.rgb = RGBColor.from_string(bg.lstrip("#"))
    elif kind == "scatter":
        from pptx.enum.chart import XL_MARKER_STYLE
        obj.has_legend = len(series) > 1
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
        obj.has_legend = len(series) > 1
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
    if chart.get("data_labels") and kind == "scatter":
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
        _force_show_values(plot)

    if combo:
        line_plot = _split_combo(obj, combo, names, colors)
        if line_plot is not None and chart.get("data_labels"):
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
    from pptx.enum.chart import XL_TICK_LABEL_POSITION
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
        try:
            ca.tick_labels.position = XL_TICK_LABEL_POSITION.LOW
        except Exception:
            pass

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
    from pptx.oxml.ns import qn
    from lxml import etree

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
    from pptx.oxml.ns import qn
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
    if len(unit) <= 3:
        return base + '" ' + unit + '"'
    return base


def _force_show_values(plot) -> None:
    """python-pptx 给环图写出的 dLbls 里 showVal=0：声明了直标却一个数都不显示。这里补正。"""
    from pptx.oxml.ns import qn
    from lxml import etree

    dlbls = plot._element.find(qn("c:dLbls"))
    if dlbls is None:
        return
    node = dlbls.find(qn("c:showVal"))
    if node is not None:
        node.set("val", "1")
        return
    node = etree.Element(qn("c:showVal"))
    node.set("val", "1")
    anchor = None
    for tag in ("c:showCatName", "c:showSerName", "c:showPercent", "c:showBubbleSize",
                "c:separator", "c:showLeaderLines"):
        anchor = dlbls.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(node)
    else:
        dlbls.append(node)


def _add_dlbls(plot_element, style: Dict[str, str], number_format: str, position: Optional[str] = None) -> None:
    """python-pptx 没有散点/折线混排的直标入口，直接写 XML；顺序必须符合 schema。"""
    from pptx.oxml.ns import qn
    from lxml import etree

    for old in plot_element.findall(qn("c:dLbls")):
        plot_element.remove(old)
    labels = etree.Element(qn("c:dLbls"))
    fmt = etree.SubElement(labels, qn("c:numFmt"))
    fmt.set("formatCode", number_format)
    fmt.set("sourceLinked", "0")
    txPr = etree.SubElement(labels, qn("c:txPr"))
    body = etree.SubElement(txPr, qn("a:bodyPr"))
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
    from pptx.oxml.ns import qn
    from lxml import etree
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


def add_table(slide, plan, page, el, style, family):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.oxml.ns import qn
    from lxml import etree

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


def add_page_number(slide, plan, page, index):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    cfg = page_numbers(plan)
    if page["id"] in cfg["hide"]:
        return None
    box = pn_box(plan, page)
    tone = page.get("page_number_tone") or ("dark" if is_dark(page_bg(page)) else "light")
    ink = LIMITS["page_number"]["ink"][tone]
    shape = slide.shapes.add_textbox(Inches(box["x"]), Inches(box["y"]), Inches(box["w"]), Inches(box["h"]))
    shape.name = "_page_number"
    tf = shape.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.RIGHT
    run = para.add_run()
    run.text = str(cfg["start"] + index)
    _set_run_font(run, (plan.get("fonts") or {}).get("body") or LIMITS["defaults"]["fonts"]["body"],
                  LIMITS["page_number"]["font_pt"], ink, False)
    return shape


# ---------------------------------------------------------------- 主流程


def build_element(slide, plan, page, el, boxes, asset_root: Path):
    kind = el["type"]
    if kind in TEXT_TYPES:
        return add_text(slide, plan, page, el)
    if kind in ("shape", "divider"):
        return add_shape(slide, page, el)
    if kind == "path":
        return add_path(slide, page, el, boxes)
    if kind == "image":
        return add_image(slide, page, el, asset_root)
    if kind == "chart":
        return add_chart(slide, plan, page, el, deck_palette(plan))
    if kind == "group":
        children = [build_element(slide, plan, page, child, boxes, asset_root) for child in el["elements"]]
        group = slide.shapes.add_group_shape([c for c in children if c is not None])
        group.name = f"{page['id']}.{el['id']}"
        if el.get("rotation"):
            group.rotation = float(el["rotation"])
        return group
    raise ValueError(f"未知元素类型：{kind}")


def _reorder(slide, built: List[Any]) -> None:
    """叠放顺序就是声明顺序。python-pptx 把组合插入到末尾，这里按声明重排一次。"""
    from pptx.oxml.ns import qn
    tree = slide.shapes._spTree
    keep_tags = {qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:grpSp"), qn("p:cxnSp")}
    kids = [c for c in tree if c.tag in keep_tags]
    index = {id(shape._element): i for i, shape in enumerate(built)}
    ordered = sorted(kids, key=lambda c: index.get(id(c), len(built)))
    for c in kids:
        tree.remove(c)
    for c in ordered:
        tree.append(c)


def build(plan: Dict[str, Any], asset_root: Path, out: Path) -> None:
    from pptx import Presentation
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    w, h = canvas(plan)
    prs = Presentation()
    prs.slide_width = Inches(w)
    prs.slide_height = Inches(h)
    blank = prs.slide_layouts[6]

    for idx, page in enumerate(plan["pages"]):
        slide = prs.slides.add_slide(blank)
        background = slide.background.fill
        background.solid()
        background.fore_color.rgb = RGBColor.from_string(page_bg(page).lstrip("#"))
        boxes = page_geometry(page)
        built = [build_element(slide, plan, page, el, boxes, asset_root) for el in page["elements"]]
        _reorder(slide, [b for b in built if b is not None])
        add_page_number(slide, plan, page, idx)

    prs.core_properties.title = plan.get("title", "")
    prs.core_properties.author = "ppt-visual-art-director"
    prs.core_properties.comments = (plan.get("visual_concept") or "")[:500]
    # 临时文件名带进程号：既不覆盖用户同名文件，也不会两次并发编译互相踩踏。
    tmp = out.with_name(f".{out.name}.{os.getpid()}.part")
    try:
        prs.save(str(tmp))
        os.replace(tmp, out)
    finally:
        if tmp.exists():
            tmp.unlink()


def write_error_deck(out: Path, message: str) -> None:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(5.9))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = "编译失败，此文件不是交付物。\n\n" + message
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.size = Pt(14 if run.text.startswith("编译失败") else 11)
    prs.save(str(out))


def utf8_stdio() -> None:
    """控制台代码页不是 UTF-8 时（Windows 中文终端很常见），中文与 ⏎ 会抛 UnicodeEncodeError。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass


def main(argv: Optional[List[str]] = None) -> int:
    utf8_stdio()
    parser = argparse.ArgumentParser(description="plan.json → 原生可编辑 PPTX")
    parser.add_argument("plan")
    parser.add_argument("out")
    parser.add_argument("--assets", default=None, help="素材目录；缺省为 plan 同级 assets/ 或 plan 所在目录")
    parser.add_argument("--report", default=None, help="把诊断写成 JSON")
    parser.add_argument("--force", action="store_true", help="允许替换已存在的输出文件")
    args = parser.parse_args(argv)

    plan_path = Path(args.plan).resolve()
    out_path = Path(args.out).resolve()
    asset_root = Path(args.assets).resolve() if args.assets else (
        plan_path.parent / "assets" if (plan_path.parent / "assets").is_dir() else plan_path.parent)

    rep = Report()
    if out_path.exists() and not args.force:
        rep.hard("E-EXIST", str(out_path), "目标文件已存在，未写入；确认要替换就加 --force")
    if out_path.is_dir():
        rep.hard("E-EXIST", str(out_path), "输出路径是目录")
    elif not out_path.parent.is_dir():
        rep.hard("E-EXIST", str(out_path.parent), "输出目录不存在，先建目录再编译")
    if args.report and Path(args.report).resolve() == out_path:
        rep.hard("E-EXIST", str(out_path), "--report 与输出文件同路径，会互相覆盖")
    plan = load_plan(plan_path, rep)
    if plan is None or rep.blocked:
        for issue in rep.items:
            print(issue)
        print(f"\n阻断：{rep.count('HARD')} 个硬错误，未写出文件。")
        return 2

    rep.extend(validate(plan, asset_root))
    for issue in rep.items:
        print(issue)
    if rep.blocked:
        print(f"\n阻断：{rep.count('HARD')} 个硬错误，未写出文件。")
        if args.report:
            Path(args.report).write_text(json.dumps({"blocked": True, "issues": rep.to_dict()},
                                                    ensure_ascii=False, indent=2), encoding="utf-8")
        return 2

    try:
        build(plan, asset_root, out_path)
    except Exception as exc:  # 运行期异常：不覆盖正式输出
        error_path = out_path.with_name(out_path.stem + ".error.pptx")
        write_error_deck(error_path, f"{type(exc).__name__}: {exc}")
        print(f"[HARD] E-RUNTIME 运行期失败：{type(exc).__name__}: {exc}\n诊断文件：{error_path}（不是交付物）")
        return 3

    print(f"\n已写出 {out_path}（{len(plan['pages'])} 页，{canvas(plan)[0]}×{canvas(plan)[1]} in）"
          f"  警告 {rep.count('SOFT')}，提示 {rep.count('INFO')}")
    if args.report:
        Path(args.report).write_text(json.dumps({
            "blocked": False,
            "output": str(out_path),
            "pages": len(plan["pages"]),
            "sha256": hashlib.sha256(out_path.read_bytes()).hexdigest(),
            "issues": rep.to_dict(),
        }, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
