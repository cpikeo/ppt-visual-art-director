#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""plan 模型与共用工具：诊断、载入、度量、颜色、几何、成品包校验。

compile / verify / revise 共用这一份实现；这里不依赖 python-pptx，可单独校验 plan。
"""

from __future__ import annotations

import json
import math
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

# 解析来自外部的 PPTX 部件一律走这个解析器：不展开实体、不联网，防 XXE 与实体膨胀。
SAFE_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False, huge_tree=False)

def parse_xml(data: bytes):
    return etree.fromstring(data, SAFE_PARSER)

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
        if ch.isdigit():
            return 0.58          # 粗体数字偏宽；按 0.55 估，100pt 的“41%”会被低估到刚好放下
        if ch in "%@&":
            return 0.90          # 百分号、@、& 接近一个字宽；按标点的 0.34 估会漏掉大字号数字的折行
        return 0.55 if ch.isalnum() else 0.34
    if ch in "×÷±":
        return 0.90
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

def chart_palette(plan: Dict[str, Any], chart: Dict[str, Any]) -> List[str]:
    """图表取色：图自带 palette 优先，否则用整稿 palette。compile、checks、verify 共用这一处。"""
    return [c.upper() for c in (chart.get("palette") or [])] or deck_palette(plan)

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

def default_asset_root(anchor: Path, explicit: Optional[str] = None) -> Path:
    """素材根目录：--assets 优先，否则取 anchor（plan 或 edit.json）同级 assets/，再否则取其所在目录。

    compile / verify / revise 共用这一个解析，三者不会各看各的目录。
    """
    if explicit:
        return Path(explicit).resolve()
    anchor = Path(anchor).resolve()
    sub = anchor.parent / "assets"
    return sub if sub.is_dir() else anchor.parent

def resolve_asset(asset_root: Path, image_id: str) -> Optional[Path]:
    """image_id → 素材文件：先按原名，再补常见扩展名。只认这两个规则，且结果必须留在素材根内。

    不做"同名不同扩展名"的兜底：那会在 hero.png 与 hero-2.jpg 并存时悄悄取到另一张图，
    而核验器用的是同一个解析函数，两边一起错，谁也发现不了。

    边界：绝对路径、含 .. 的路径、经符号链接指向素材根之外的文件一律视为找不到。
    plan 与 edit.json 来自外部，image_id 不能成为读取任意本机文件的入口。
    """
    name = str(image_id or "")
    if not name or "\x00" in name or Path(name).is_absolute() or ".." in Path(name.replace("\\", "/")).parts:
        return None
    root = Path(asset_root).resolve()

    def inside(candidate: Path) -> Optional[Path]:
        if not candidate.is_file():
            return None
        real = candidate.resolve()
        try:
            real.relative_to(root)
        except ValueError:
            return None
        return real

    direct = root / name
    hit = inside(direct)
    if hit:
        return hit
    if direct.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}:
        return None
    for ext in (".png", ".jpg", ".jpeg", ".webp", ".PNG", ".JPG", ".JPEG", ".WEBP"):
        hit = inside(root / (name + ext))
        if hit:
            return hit
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

def backdrop_image(page: Dict[str, Any], target: Dict[str, Any]) -> Optional[Tuple[Dict[str, Any], bool]]:
    """文字落在哪张图上：返回（图片元素，图与文字之间是否隔着色块/渐变蒙版）；没有图则 None。

    判定与 resolve_bg 一致：图覆盖文字框 ≥ 60%。取声明顺序里最靠上的那张。
    """
    box = explicit_box(target)
    if box is None:
        return None
    area = max(box[2] * box[3], 1e-6)
    found: Optional[Dict[str, Any]] = None
    scrim = False
    for item in flatten(page):
        el = item["el"]
        if el is target:
            break
        below = explicit_box(el)
        if below is None:
            continue
        ix, iy = _intersection(below, box)
        cover = ix * iy / area
        if el.get("type") == "image" and cover >= 0.6:
            found, scrim = el, False
        elif found is not None and el.get("type") in ("shape", "divider") and cover >= 0.3:
            scrim = True
    return (found, scrim) if found is not None else None

_LIN = [((v / 255.0) / 12.92 if v / 255.0 <= 0.03928 else (((v / 255.0) + 0.055) / 1.055) ** 2.4) for v in range(256)]

def _lum_rgb(rgb: Tuple[int, int, int]) -> float:
    return 0.2126 * _LIN[rgb[0]] + 0.7152 * _LIN[rgb[1]] + 0.0722 * _LIN[rgb[2]]

def zone_probe(path: Path, img_el: Dict[str, Any], text_box: Box, page_background: str) -> Optional[Dict[str, Any]]:
    """文字框正下方的真实像素：按裁切、duotone、透明度还原最终落点，返回相对亮度样本与纹理强度。

    这是静态近似（不含渲染引擎的色彩管理，也不含图上其他元素），用来在编译期发现“字落在亮暗交界或高纹理区”，
    结论仍以真实渲染为准。读不了图返回 None。
    """
    img_box = explicit_box(img_el)
    if img_box is None:
        return None
    try:
        from PIL import Image
        with Image.open(path) as src:
            src = src.convert("RGB")
            pw, ph = src.size
            l, r, t, b = _crop_geometry(pw, ph, img_box[2], img_box[3], img_el.get("crop_focus"))
            x0, y0 = max(img_box[0], text_box[0]), max(img_box[1], text_box[1])
            x1 = min(img_box[0] + img_box[2], text_box[0] + text_box[2])
            y1 = min(img_box[1] + img_box[3], text_box[1] + text_box[3])
            if x1 - x0 <= 0.02 or y1 - y0 <= 0.02:
                return None

            def src_px(fx: float, fy: float, total_w: float, total_h: float, lo_x: float, lo_y: float):
                return ((l + ((fx - lo_x) / total_w) * (1 - l - r)) * pw,
                        (t + ((fy - lo_y) / total_h) * (1 - t - b)) * ph)

            ax, ay = src_px(x0, y0, img_box[2], img_box[3], img_box[0], img_box[1])
            bx, by = src_px(x1, y1, img_box[2], img_box[3], img_box[0], img_box[1])
            n = int(LIMITS["image"]["zone"]["samples"])
            region = src.crop((int(max(ax, 0)), int(max(ay, 0)), max(int(bx), int(ax) + 1), max(int(by), int(ay) + 1)))
            region = region.resize((n, n), Image.BOX)
            pixels = list(region.get_flattened_data() if hasattr(region, "get_flattened_data") else region.getdata())
    except Exception:
        return None
    duo = img_el.get("duotone")
    alpha = float(img_el.get("alpha", 1.0))
    bg = hex_rgb(page_background)
    lums: List[float] = []
    grays: List[float] = []
    for px in pixels:
        rgb = px
        if duo:
            g = (0.2126 * px[0] + 0.7152 * px[1] + 0.0722 * px[2]) / 255.0
            d, lt = hex_rgb(duo["dark"]), hex_rgb(duo["light"])
            rgb = tuple(round(d[i] + (lt[i] - d[i]) * g) for i in range(3))
        if alpha < 0.999:
            rgb = tuple(round(bg[i] + (rgb[i] - bg[i]) * alpha) for i in range(3))
        lums.append(_lum_rgb(rgb))
        grays.append((0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]))
    mean = sum(grays) / len(grays)
    std = (sum((g - mean) ** 2 for g in grays) / len(grays)) ** 0.5
    return {"lums": lums, "gray_std": std, "gray_mean": mean}

def zone_contrast(ink_hex: str, probe: Dict[str, Any], limit: float) -> Tuple[float, float]:
    """（达标像素占比，最不利 5% 像素的对比度）。"""
    li = rel_luminance(ink_hex)
    ratios = sorted((max(li, v) + 0.05) / (min(li, v) + 0.05) for v in probe["lums"])
    ok = sum(1 for r in ratios if r >= limit) / len(ratios)
    return ok, ratios[max(int(len(ratios) * 0.05) - 1, 0)]

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

PPTX_MAIN_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"

def package_problems(path: Path, check_suffix: bool = True) -> List[str]:
    """文件是否真的是无宏 .pptx：扩展名、主部件内容类型与包内部件必须一致。

    PPTM（宏启用）与扩展名不符的包都在这里拦下；返回空列表才算合格。
    """
    import zipfile

    path = Path(path)
    problems: List[str] = []
    if check_suffix and path.suffix.lower() != ".pptx":
        problems.append(f"扩展名 {path.suffix or '(无)'} 不是 .pptx")
    try:
        with zipfile.ZipFile(path) as zf:
            infos = zf.infolist()
            names = [i.filename for i in infos]
            unpacked = sum(i.file_size for i in infos)
            types = parse_xml(zf.read("[Content_Types].xml"))
    except (OSError, KeyError, zipfile.BadZipFile, etree.XMLSyntaxError) as exc:
        return problems + [f"不是可读的 OOXML 包：{type(exc).__name__}"]
    if any(n.startswith("/") or ".." in n.split("/") for n in names):
        problems.append("包内含越界部件名（绝对路径或 ..）")
    if unpacked > LIMITS["package"]["max_unpacked_mb"] * 1024 * 1024:
        problems.append(f"解压后超过 {LIMITS['package']['max_unpacked_mb']} MB，拒绝处理（疑似压缩炸弹）")
    overrides = {o.get("PartName"): o.get("ContentType") or "" for o in types if o.get("PartName")}
    main_ct = overrides.get("/ppt/presentation.xml")
    if main_ct != PPTX_MAIN_CT:
        problems.append(f"主部件内容类型为 {main_ct}，不是 .pptx 格式（可能是宏启用或其他格式）")
    if any("vbaProject" in n for n in names) or any("macroEnabled" in v for v in overrides.values()):
        problems.append("含宏部件（vbaProject）或宏启用内容类型")
    return problems

def utf8_stdio() -> None:
    """控制台代码页不是 UTF-8 时（Windows 中文终端很常见），中文与 ⏎ 会抛 UnicodeEncodeError。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass


def alt_text(el: Dict[str, Any]) -> str:
    """元素的替代文字：显式 alt 优先；图片缺省取 asset_reason，图表缺省取图题。装饰元素返回空串。"""
    if el.get("alt"):
        return str(el["alt"])
    if el.get("type") == "image":
        return str(el.get("asset_reason") or "")
    if el.get("type") == "chart":
        return str((el.get("chart") or {}).get("title") or "")
    return ""
