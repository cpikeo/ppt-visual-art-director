#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
plan.json → 原生可编辑 PPTX。

用法：
    python scripts/compile.py plan.json out.pptx [--assets DIR] [--report report.json] [--force]

只做确定性执行：校验契约、按 plan 建对象、写出文件。不自动缩字、换色、补数据或联网取图。
除对象级检查外，还会报整稿层面的模式信号（连续同构、标题下细线、无功能标签、素材复用、
待主张页、满幅图开场）——它们只指出位置，是否接受由人判断。

硬错误（E-）在写文件前阻断；输出与报告都不得指向 plan 自己。目标文件已存在时默认不动，
确认要替换才加 --force。先写临时文件再原子替换，失败不留半成品；运行期异常另写
<名称>.error.pptx 供排查，该文件不是交付物。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree

HERE = Path(__file__).resolve().parent
LIMITS: Dict[str, Any] = json.loads((HERE / "limits.json").read_text(encoding="utf-8-sig"))
SCHEMA: Dict[str, Any] = json.loads((HERE / "plan.schema.json").read_text(encoding="utf-8-sig"))

TEXT_TYPES = ("title", "body", "text", "number", "caption", "footer")
DASH = {"solid": "solid", "dash": "dash", "dot": "sysDot", "dash_dot": "dashDot", "long_dash": "lgDash"}
EMU = 914400.0
PT_EMU = 12700.0

Box = Tuple[float, float, float, float]

# OOXML 命名空间。自己拼簇名而不用 pptx.oxml.ns.qn：那会把 pptx 提升为导入期依赖，
# 而本模块要先能在"只校验 plan"的场景下单独运行。
_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return "{%s}%s" % (_NS[prefix], local)


# ---------------------------------------------------------------- 诊断


class BadOp(Exception):
    """revise.py 的操作无法执行：定位不到对象、参数不合契约。不猜，直接停。"""


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


def _para_width_em(seg: str, pieces: List[Dict[str, Any]], pt: float) -> float:
    """一段的宽度，以"基准字号 1 em"计。逐 run 按各自字号折算：
    标题里放一个 96pt 数字加 24pt 单位，按整段取同一个字号估高会差出好几行。"""
    if not pieces:
        return sum(_em_width(c) for c in seg)
    total = 0.0
    for piece in pieces:
        size = float(piece.get("font_size") or pt)
        total += sum(_em_width(c) for c in str(piece.get("text", ""))) * size / pt
    return total


def fit_lines(el: Dict[str, Any], st: Dict[str, Any], box_w: float) -> List[int]:
    """逐段行数。按每段最大字号与逐 run 折算的宽度算，不用元素级字号一刀切。"""
    inset = st["pad"][0]
    usable = max(box_w - 2 * inset, 0.05)
    paras = paragraph_runs(el)
    if paras is None:
        paras = [(str(el.get("content", "")), [])]
    out: List[int] = []
    for seg, pieces in paras:
        base = max([float(p.get("font_size") or st["size"]) for p in pieces] or [st["size"]])
        cap = usable / (base / 72.0)
        out.append(max(1, math.ceil(_para_width_em(seg, pieces, base) / cap - 1e-9)))
    return out or [1]


def fit_height_in(el: Dict[str, Any], st: Dict[str, Any], box_w: float) -> float:
    """框高需求（英寸）。行数、字号、行距、段后距都算进去；估不准时以真实渲染为准。"""
    pitch = LIMITS["type"]["line_pitch_em"]
    paras = paragraph_runs(el) or [(str(el.get("content", "")), [])]
    lines = fit_lines(el, st, box_w)
    total = 0.0
    for (seg, pieces), count in zip(paras, lines):
        base = max([float(p.get("font_size") or st["size"]) for p in pieces] or [st["size"]])
        total += count * base * pitch * st["line_spacing"] / 72.0
    space_after = float(el.get("space_after_pt") or 0)
    if space_after:
        total += space_after * max(len(paras) - 1, 1) / 72.0
    return total


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
    """image_id → 素材文件：先按原名，再补常见扩展名。只认这两个规则。

    不做"同名不同扩展名"的兜底：那会在 hero.png 与 hero-2.jpg 并存时悄悄取到另一张图，
    而核验器用的是同一个解析函数，两边一起错，谁也发现不了。
    """
    direct = asset_root / image_id
    if direct.is_file():
        return direct
    known = {".png", ".jpg", ".jpeg", ".webp"}
    if direct.suffix.lower() in known:
        return None
    for ext in (".png", ".jpg", ".jpeg", ".webp", ".PNG", ".JPG", ".JPEG", ".WEBP"):
        candidate = asset_root / (image_id + ext)
        if candidate.is_file():
            return candidate
    return None


# ---------------------------------------------------------------- 文本式样


def text_pad(el: Dict[str, Any]) -> Tuple[float, float]:
    """文本框内边距：缺省 0.05/0.03 in。贴网格轴线的文字写 pad_in: 0，否则光学边缘会偏 0.05in。"""
    if el.get("pad_in") is not None:
        pad = float(el["pad_in"])
        return pad, pad
    return LIMITS["type"]["inset_in"]["x"], LIMITS["type"]["inset_in"]["y"]


def text_style(plan: Dict[str, Any], page: Dict[str, Any], el: Dict[str, Any]) -> Dict[str, Any]:
    kind = el["type"]
    size = float(el.get("font_size") or LIMITS["type"]["default_pt"].get(kind, 18))
    fonts = plan.get("fonts") or {}
    if kind == "number":
        family = el.get("font") or fonts.get("number") or fonts.get("heading")
    else:
        family = el.get("font") or fonts.get("heading" if kind == "title" else "body")
    family = family or LIMITS["defaults"]["fonts"]["heading" if kind in ("title", "number") else "body"]
    color = el.get("color")
    if not color:
        color = LIMITS["defaults"]["ink_on_dark"] if is_dark(page_bg(page)) else LIMITS["defaults"]["ink"]
    pad_x, pad_y = text_pad(el)
    return {
        "size": size,
        "family": family,
        "color": (color or LIMITS["defaults"]["ink"]).upper(),
        "bold": bool(el.get("bold", LIMITS["type"]["bold"].get(kind, False))),
        "line_spacing": float(el.get("line_spacing") or LIMITS["type"]["line_spacing"].get(kind, 1.25)),
        "align": el.get("align", "left"),
        "valign": el.get("valign", "top"),
        "spacing": float(el.get("letter_spacing_pt") or 0),
        "pad": (pad_x, pad_y),
    }


def paragraph_runs(el: Dict[str, Any]) -> Optional[List[Tuple[str, List[Dict[str, Any]]]]]:
    """把 content 与 runs 归一成 [(段文本, [run…]), …]。

    runs 里的 \\n 就是分段：内联样式可以跨段，标题第一行深墨、第二行强调色不必拆成两个文本框。
    拼接与 content 不等时返回 None，由校验报错，不猜。
    """
    content = str(el.get("content", ""))
    paras = content.split("\n")
    runs = el.get("runs")
    if not runs:
        return [(seg, [{"text": seg}] if seg else []) for seg in paras]
    if "".join(str(r.get("text", "")) for r in runs) != content:
        return None
    buckets: List[List[Dict[str, Any]]] = [[]]
    for spec in runs:
        style = {k: v for k, v in spec.items() if k != "text"}
        for i, piece in enumerate(str(spec.get("text", "")).split("\n")):
            if i:
                buckets.append([])
            if piece:
                buckets[-1].append({**style, "text": piece})
    if len(buckets) != len(paras):
        return None
    return [(paras[i], buckets[i]) for i in range(len(paras))]


# ---------------------------------------------------------------- 校验


CARRIERS = ("shape", "divider", "path")

# a:blip 里排在 duotone 之后的元素（ECMA-376 CT_Blip 的 EG_ColorTransform 次序）。
_BLIP_AFTER_DUOTONE = ("fillOverlay", "grayscale", "hsl", "lum", "tint")


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
            rbox = rot_box(box, el.get("rotation"))
            if x + ew <= 0 or y + eh <= 0 or x >= w or y >= h:
                rep.hard("E-GEOM", eloc, f"元素完全在画布外：x={x} y={y} w={ew} h={eh}")
            elif kind in TEXT_TYPES + ("chart",) and (rbox[0] < -tol or rbox[1] < -tol
                                                      or rbox[0] + rbox[2] > w + tol
                                                      or rbox[1] + rbox[3] > h + tol):
                what = "文本框" if kind in TEXT_TYPES else "图表"
                rep.soft("W-GEOM", eloc, f"{what}（含旋转）伸出画布（{w}×{h}），边缘内容会被裁掉")

            if el.get("overlap_allowed") and kind not in CARRIERS and not (
                    kind == "image" and el.get("image_role") in ("background", "texture")):
                rep.hard("E-OVERLAP", eloc,
                         "overlap_allowed 只能声明在承载层（背景图/蒙版/色块/闭合路径）上；被压的元素不能自我豁免")
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
    return rep


def check_text(plan, page, el, loc, ew, eh, rep: Report) -> None:
    content = el.get("content", "")
    if not str(content).strip():
        rep.hard("E-TEXT", loc, "文本内容为空；删掉这个元素，或写入内容")
    if paragraph_runs(el) is None:
        rep.hard("E-REF", loc, "runs 拼接与 content 不一致（含 \\n）；分段要对得上")
    st = text_style(plan, page, el)
    need = fit_height_in(el, st, ew)
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
            rep.soft("W-CHART", loc, "表格不画图题；用页面 title 或 caption 承担结论")
        return

    series = data.get("series") or []
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
            if not chart.get("data_labels") and chart.get("legend") is False:
                rep.hard("E-CHART", loc, "饼/环既没有直标也没有图例：读者不知道每一块是什么")

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
        if stacked == "percent" and kind == "bar":
            bad = []
            for j in range(len(data["categories"])):
                total = sum(float(s["values"][j]) for s in series)
                if abs(total - 100) > 0.5:
                    bad.append(f"{data['categories'][j]}={total:g}")
            if bad:
                message = f"百分比堆叠要求每列合计 100，实际 {bad}；图形会归一化而直标仍是原值"
                if chart.get("data_labels"):
                    rep.hard("E-CHART", loc, message)
                else:
                    rep.soft("W-CHART", loc, message)
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

    need = len(data.get("categories") or []) if kind in ("pie", "donut") else len(series)
    if len(palette) < need:
        rep.soft("W-CHART", loc, f"palette 只有 {len(palette)} 色，需要 {need}；将循环取色")
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
    if axis.get("note"):
        rep.soft("W-NOTE", loc,
                 f"value_axis.note「{str(axis['note'])[:40]}」不会画进图里：截轴起点要读者看得见，"
                 f"必须落到页面的 caption、轴名或图题上")


CAPS_LABEL = re.compile(r"^[A-Z][A-Z0-9 &/·.\-–—]{2,23}$")


def design_signals(plan: Dict[str, Any]) -> Report:
    """整稿层面的模式信号：指出"这一处该由你看一眼"，不是评分，也不下审美结论。

    只报能说出位置与可核对事实的东西：连续同构、标题下装饰线、无功能英文标签、
    素材复用、待主张页、满幅图开场。判为可接受就在交付说明里写理由。
    """
    rep = Report()
    pages = plan.get("pages") or []
    if len(pages) < 2:
        return rep
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
    # 有效分辨率按"最终留在画面里的那部分像素"算：裁得越紧，同样的框里 ppi 越高。
    return min(px_w * (1 - left - right) / w, px_h * (1 - top - bottom) / h), loss


def check_asset(loc: str, image_id: str, asset: Dict[str, Any], asset_root: Path, rep: Report) -> None:
    rights = asset.get("rights", "pending")
    generated = bool(asset.get("generated"))
    if rights == "blocked":
        rep.hard("E-ASSET", loc, f"{image_id} 权利状态 blocked，不得使用")
    elif rights == "pending":
        rep.soft("W-ASSET", loc, f"{image_id} 权利未确认（pending）；交付前确认或改用已清权素材")
    elif rights == "cleared" and not asset.get("reference"):
        rep.soft("W-ASSET", loc, f"{image_id} 声明已清权但没有 reference（授权页/合同/编号）")
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
        rep.soft("W-ASSET", loc, f"{image_id} 是生成图但没有登记 sha256，无法证明用的是哪一次结果")
    if path and gen.get("sha256"):
        actual = hashlib.sha256(path.read_bytes()).hexdigest()
        if actual.lower() != str(gen["sha256"]).lower():
            rep.hard("E-ASSET", loc, f"{image_id} 实际 SHA-256 与登记不符")


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


def corner_adj(radius_in: float, w: float, h: float) -> int:
    """圆角半径 → adj（短边的百分比，0.5 即胶囊）。编译与核验共用这一个换算。"""
    short = max(min(w, h), 1e-6)
    return int(round(min(max(radius_in / short, 0.0), 0.5) * 100000))


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
    """新建对象一律连东亚字族一起写：只写 latin 时中文会被渲染器换掉，整稿字族就断了。"""
    set_run_attrs(run, font=family, font_ea=family, size=size, color=color, bold=bold, spacing=spacing)


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
    pad_x, pad_y = st["pad"]
    tf.margin_left = tf.margin_right = Inches(pad_x)
    tf.margin_top = tf.margin_bottom = Inches(pad_y)
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[st["valign"]]
    align = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}[st["align"]]

    paragraphs = paragraph_runs(el) or [(str(el.get("content", "")), [{"text": str(el.get("content", ""))}])]
    for idx, (seg, pieces) in enumerate(paragraphs):
        para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = st["line_spacing"]
        if el.get("space_after_pt"):
            para.space_after = Pt(float(el["space_after_pt"]))
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
    if el.get("corner_in") is not None:
        _set_corner_radius(shape, float(el["corner_in"]), box)
    if shape.has_text_frame:
        shape.text_frame.text = ""
    return shape


def add_path(slide, page, el, boxes):
    from pptx.util import Emu
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
    # 先写透明度，再写双色调：a:blip 的子元素有固定次序，反过来插会得到非法顺序。
    if el.get("alpha") and float(el["alpha"]) < 1:
        blip = pic._element.blipFill.find(qn("a:blip"))
        if blip is not None:
            etree.SubElement(blip, qn("a:alphaModFix")).set("amt", str(int(round(float(el["alpha"]) * 100000))))
    if el.get("duotone"):
        _apply_duotone(pic, el["duotone"])
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


def add_chart(slide, plan, page, el, palette: List[str]):
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
    colors = _series_colors(palette, len(data.categories) if kind in ("pie", "donut") else len(series))
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
    dark, light = style["title"], "#FFFFFF"
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


def rot_box(box: Box, rotation: Any) -> Box:
    """旋转后的外接框。声明的框是没转过的，页角与遮挡要按真实占位算。

    外接框的两边是 |w·cos| + |h·sin| 与 |w·sin| + |h·cos|：只用 cos/sin 会把
    100°–170° 这一类角度算成负数宽度。
    """
    import math

    if not rotation:
        return box
    x, y, w, h = box
    a = math.radians(float(rotation) % 180)
    c, s = abs(math.cos(a)), abs(math.sin(a))
    nw, nh = w * c + h * s, w * s + h * c
    return (x + w / 2 - nw / 2, y + h / 2 - nh / 2, nw, nh)


def pn_surface(plan: Dict[str, Any], page: Dict[str, Any], box: Dict[str, float]) -> str:
    """页码实际落在什么底色上：取最后绘制、且盖住页码中心的不透明色块或页面底色。

    图片与图表上无法靠一个色值判读，遇到就停在页面底色，由 page_number_tone 声明。
    """
    cx, cy = box["x"] + box["w"] / 2, box["y"] + box["h"] / 2
    boxes = page_geometry(page)
    surface = page_bg(page)
    for item in flatten(page):
        el = item["el"]
        if el["type"] in TEXT_TYPES or el["type"] in ("image", "chart", "path"):
            continue
        color = el.get("color")
        if not color:
            continue
        gb = boxes.get(str(el.get("id")))
        if gb is None:
            continue
        gb = rot_box(gb, el.get("rotation"))
        if gb[0] <= cx <= gb[0] + gb[2] and gb[1] <= cy <= gb[1] + gb[3]:
            surface = str(color).upper()
    return surface


def pn_ink(plan: Dict[str, Any], page: Dict[str, Any]) -> str:
    """页码墨色：声明优先；否则按页码实际落点，在两档墨里选能读的那个。

    只看 page.background 会漏掉"底色其实由一张铺满的色块给出"这种页——
    深蓝底上写深墨页码，等于没写。
    """
    box = pn_box(plan, page)
    tone = page.get("page_number_tone")
    if tone:
        return LIMITS["page_number"]["ink"][tone]
    surface = pn_surface(plan, page, box)
    base = "dark" if is_dark(page_bg(page)) else "light"
    ink = LIMITS["page_number"]["ink"][base]
    other = LIMITS["page_number"]["ink"]["light" if base == "dark" else "dark"]
    return ink if contrast_ratio(ink, surface) >= contrast_ratio(other, surface) else other


def add_page_number(slide, plan, page, index):
    from pptx.util import Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    cfg = page_numbers(plan)
    if page["id"] in cfg["hide"]:
        return None
    box = pn_box(plan, page)
    ink = pn_ink(plan, page)
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
    obj = None
    if kind in TEXT_TYPES:
        obj = add_text(slide, plan, page, el)
    elif kind in ("shape", "divider"):
        obj = add_shape(slide, page, el)
    elif kind == "path":
        obj = add_path(slide, page, el, boxes)
    elif kind == "image":
        obj = add_image(slide, page, el, asset_root)
    elif kind == "chart":
        obj = add_chart(slide, plan, page, el, deck_palette(plan))
    elif kind == "group":
        children = [build_element(slide, plan, page, child, boxes, asset_root) for child in el["elements"]]
        obj = slide.shapes.add_group_shape([c for c in children if c is not None])
        obj.name = f"{page['id']}.{el['id']}"
    else:
        raise ValueError(f"未知元素类型：{kind}")
    # 旋转在出口处统一应用：只有一处写 rot，也就不可能出现"某种类型忘了转"。
    # 角度先归一到 [0,360)：270 与 -90 是同一件事，不该因为写法不同被 schema 拦下。
    if obj is not None and el.get("rotation"):
        obj.rotation = float(el["rotation"]) % 360
    return obj


def _reorder(slide, built: List[Any]) -> None:
    """叠放顺序就是声明顺序。python-pptx 把组合插入到末尾，这里按声明重排一次。"""
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
    brief = plan.get("brief") or {}
    prs.core_properties.comments = (brief.get("concept") or "")[:500]
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
    report_path = Path(args.report).resolve() if args.report else None
    # 源文件与交付物分家：输出、报告都不得落在 plan 上，报告也不得落在成品上。
    for label, target in (("输出文件", out_path), ("--report", report_path)):
        if target is not None and target == plan_path:
            rep.hard("E-EXIST", str(target), f"{label}与 plan 同路径，会覆盖源文件")
    if report_path is not None and report_path == out_path:
        rep.hard("E-EXIST", str(out_path), "--report 与输出文件同路径，会互相覆盖")
    if out_path.suffix.lower() not in (".pptx", ".pptm"):
        rep.hard("E-EXIST", str(out_path), f"输出文件名应以 .pptx 结尾，收到 {out_path.name!r}")
    if out_path.exists() and not args.force:
        rep.hard("E-EXIST", str(out_path), "目标文件已存在，未写入；确认要替换就加 --force")
    if out_path.is_dir():
        rep.hard("E-EXIST", str(out_path), "输出路径是目录")
    elif not out_path.parent.is_dir():
        rep.hard("E-EXIST", str(out_path.parent), "输出目录不存在，先建目录再编译")
    plan = load_plan(plan_path, rep)
    if plan is None or rep.blocked:
        for issue in rep.items:
            print(issue)
        print(f"\n阻断：{rep.count('HARD')} 个硬错误，未写出文件。")
        return 2

    rep.extend(validate(plan, asset_root))
    rep.extend(design_signals(plan))
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
