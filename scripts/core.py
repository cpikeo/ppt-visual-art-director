#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
共享契约、度量、几何、色彩、文本样式与包校验基础模块。

由 compile.py、verify.py、revise.py 共用，确保数据契约、阈值与 OOXML 常量只有一份实现。
"""

from __future__ import annotations

import json
import math
import re
import sys
import zipfile
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
PPTX_MAIN_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"

Box = Tuple[float, float, float, float]

_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return "{%s}%s" % (_NS[prefix], local)


def inch(value: Optional[int]) -> Optional[float]:
    return None if value is None else round(value / EMU, 2)


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


# ---------------------------------------------------------------- 度量与色彩


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
    if not pieces:
        return sum(_em_width(c) for c in seg)
    total = 0.0
    for piece in pieces:
        size = float(piece.get("font_size") or pt)
        total += sum(_em_width(c) for c in str(piece.get("text", ""))) * size / pt
    return total


def fit_lines(el: Dict[str, Any], st: Dict[str, Any], box_w: float) -> List[int]:
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
    pitch = LIMITS["type"]["line_pitch_em"]
    paras = paragraph_runs(el) or [(str(el.get("content", "")), [])]
    lines = fit_lines(el, st, box_w)
    total = 0.0
    for (_seg, pieces), count in zip(paras, lines):
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
    out: List[Dict[str, Any]] = []

    def walk(items: List[Dict[str, Any]]) -> None:
        for el in items:
            out.append(el)
            if el.get("type") == "group":
                walk(el.get("elements") or [])

    walk(page.get("elements") or [])
    return out


def page_geometry(page: Dict[str, Any]) -> Dict[str, Box]:
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


def rot_box(box: Box, deg: Optional[float]) -> Box:
    if not deg or deg % 360 == 0:
        return box
    x, y, w, h = box
    cx, cy = x + w / 2, y + h / 2
    rad = math.radians(deg)
    c, s = abs(math.cos(rad)), abs(math.sin(rad))
    nw, nh = w * c + h * s, w * s + h * c
    return (cx - nw / 2, cy - nh / 2, nw, nh)


# ---------------------------------------------------------------- 素材与文本样式


def resolve_asset(asset_root: Path, image_id: str) -> Optional[Path]:
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


def text_pad(el: Dict[str, Any]) -> Tuple[float, float]:
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


def corner_adj(radius_in: float, w: float, h: float) -> int:
    """圆角半径 → adj（短边的百分比，0.5 即胶囊）。编译与核验共用这一个换算。"""
    short = max(min(w, h), 1e-6)
    return int(round(min(max(radius_in / short, 0.0), 0.5) * 100000))


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


# ---------------------------------------------------------------- 包校验与终端


def package_problems(path: Path, check_suffix: bool = True) -> List[str]:
    path = Path(path)
    problems: List[str] = []
    if check_suffix and path.suffix.lower() != ".pptx":
        problems.append(f"扩展名 {path.suffix or '(无)'} 不是 .pptx")
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            types = etree.fromstring(zf.read("[Content_Types].xml"))
    except (OSError, KeyError, zipfile.BadZipFile, etree.XMLSyntaxError) as exc:
        return problems + [f"不是可读的 OOXML 包：{type(exc).__name__}"]
    overrides = {o.get("PartName"): o.get("ContentType") or "" for o in types if o.get("PartName")}
    main_ct = overrides.get("/ppt/presentation.xml")
    if main_ct != PPTX_MAIN_CT:
        problems.append(f"主部件内容类型为 {main_ct}，不是 .pptx 格式（可能是宏启用或其他格式）")
    if any("vbaProject" in n for n in names) or any("macroEnabled" in v for v in overrides.values()):
        problems.append("含宏部件（vbaProject）或宏启用内容类型")
    return problems


def utf8_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass


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
