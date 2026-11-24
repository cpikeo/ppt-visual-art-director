#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
设计系统与版式配方：outline.json → plan.json 初稿。

判断仍归艺术总监（概念、取舍、文案、哪页用哪种讲法）；这里只做**算术**：
  · 色彩系统：由一个种子色按 CIELAB 推出环境/结构/文字/强调/数据色，并保证对比度与灰度可辨；
  · 字号阶梯：按观看场景（hall / meeting / screen）取一整套；
  · 版式配方：12 栏网格 + 同一条左轴，框高用 model.fit_height_in 算出，不靠猜。

用法：
    python scripts/scaffold.py system --seed "#1F3A5F" [--mood warm|cool|auto] [--scene meeting]
    python scripts/scaffold.py init outline.json plan.json [--force]

生成的 plan 是起点不是终点：之后照常 compile → verify → 看图，再手改几何与文案。
配方清单与 outline 写法见 references/build.md「scaffold」一节。
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import model as M

# ---------------------------------------------------------------- 色彩系统（CIELAB）

_WHITE = (0.95047, 1.0, 1.08883)


def _lin(c: float) -> float:
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _gam(c: float) -> float:
    return 12.92 * c if c <= 0.0031308 else 1.055 * (c ** (1 / 2.4)) - 0.055


def hex_to_lch(hex_color: str) -> Tuple[float, float, float]:
    r, g, b = (_lin(v / 255.0) for v in M.hex_rgb(hex_color))
    x = 0.4124564 * r + 0.3575761 * g + 0.1804375 * b
    y = 0.2126729 * r + 0.7151522 * g + 0.0721750 * b
    z = 0.0193339 * r + 0.1191920 * g + 0.9503041 * b
    f = lambda t: t ** (1 / 3) if t > 216 / 24389 else (24389 / 27 * t + 16) / 116
    fx, fy, fz = f(x / _WHITE[0]), f(y / _WHITE[1]), f(z / _WHITE[2])
    lab_l, a, b2 = 116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz)
    return lab_l, math.hypot(a, b2), math.degrees(math.atan2(b2, a)) % 360


def _lch_rgb(lab_l: float, c: float, h: float) -> Optional[Tuple[float, float, float]]:
    a, b = c * math.cos(math.radians(h)), c * math.sin(math.radians(h))
    fy = (lab_l + 16) / 116
    fx, fz = fy + a / 500, fy - b / 200
    inv = lambda t: t ** 3 if t ** 3 > 216 / 24389 else (116 * t - 16) / (24389 / 27)
    x, y, z = inv(fx) * _WHITE[0], inv(fy) * _WHITE[1], inv(fz) * _WHITE[2]
    lin = (3.2404542 * x - 1.5371385 * y - 0.4985314 * z,
           -0.9692660 * x + 1.8760108 * y + 0.0415560 * z,
           0.0556434 * x - 0.2040259 * y + 1.0572252 * z)
    if any(v < -1e-4 or v > 1 + 1e-4 for v in lin):
        return None
    return tuple(_gam(min(max(v, 0.0), 1.0)) for v in lin)  # type: ignore


def lch_to_hex(lab_l: float, c: float, h: float) -> str:
    """超出色域时降低彩度而不是截断通道，色相才不会漂。"""
    lab_l = min(max(lab_l, 0.0), 100.0)
    rgb = None
    while c >= 0:
        rgb = _lch_rgb(lab_l, c, h)
        if rgb:
            break
        c -= 1.0
    rgb = rgb or (0.0, 0.0, 0.0)
    return "#" + "".join(f"{int(round(min(max(v, 0), 1) * 255)):02X}" for v in rgb)


def _until(contrast_to: str, minimum: float, lab_l: float, c: float, h: float, step: float) -> str:
    """沿明度方向走，直到对 contrast_to 的对比度达标。"""
    for _ in range(100):
        color = lch_to_hex(lab_l, c, h)
        if M.contrast_ratio(color, contrast_to) >= minimum:
            return color
        lab_l += step
    return lch_to_hex(max(0.0, min(100.0, lab_l)), c, h)


MOOD_HUE = {"warm": 75.0, "cool": 255.0}
SCENES = {  # 论断标题 / 正文 / 关键数字 / 注释 / 封面大标题（pt）——design.md 字号阶梯的默认档
    "hall": {"title": 52, "body": 26, "number": 150, "caption": 16, "hero": 72},
    "meeting": {"title": 36, "body": 20, "number": 120, "caption": 14, "hero": 60},
    "screen": {"title": 28, "body": 16, "number": 80, "caption": 13, "hero": 44},
}
L_STEPS = [44, 63, 25, 72.5, 53.5, 34.5]        # 浅底数据色明度：首色贴近强调色，两两 ≥ 9.5 L*，浅底对比 ≥ 2.2
L_STEPS_DARK = [63, 82, 44, 72.5, 53.5, 91.5]   # 深底数据色明度：同样两两 ≥ 9.5 L*


def design_system(seed: str, mood: str = "auto", scene: str = "meeting") -> Dict[str, Any]:
    seed_l, seed_c, seed_h = hex_to_lch(seed)
    tint_h = MOOD_HUE.get(mood, seed_h)
    tint_c = 3.0 if mood in MOOD_HUE else 2.0

    def neutral(lab_l: float, c_mul: float = 1.0) -> str:
        return lch_to_hex(lab_l, tint_c * c_mul, tint_h)

    light_bg = neutral(97.5)
    dark_bg = neutral(12.0, 1.5)
    light = {
        "bg": light_bg, "bg_alt": neutral(93.5),
        "ink": _until(light_bg, 12.0, 14, tint_c, tint_h, -1),
        "ink2": _until(light_bg, 7.0, 36, tint_c, tint_h, -1),
        "muted": _until(light_bg, 4.8, 52, tint_c, tint_h, -1),
        "rule": neutral(84),
        "accent": _until(light_bg, 4.5, min(seed_l, 55), max(seed_c, 30), seed_h, -1),
    }
    dark = {
        "bg": dark_bg, "bg_alt": neutral(18.0, 1.5),
        "ink": _until(dark_bg, 12.0, 95, tint_c, tint_h, 1),
        "ink2": _until(dark_bg, 7.0, 78, tint_c, tint_h, 1),
        "muted": _until(dark_bg, 4.8, 62, tint_c, tint_h, 1),
        "rule": neutral(30, 1.5),
        "accent": _until(dark_bg, 4.5, max(seed_l, 62), max(seed_c, 30), seed_h, 1),
    }
    hues = [seed_h, seed_h + 180, seed_h + 35, seed_h + 215, seed_h + 70, seed_h + 300]
    chroma = [max(min(seed_c, 70), 35)] + [32, 38, 30, 34, 28]  # 只有首色是饱和的，其余收彩度，整套才耐看
    data = [lch_to_hex(L_STEPS[i], chroma[i], hues[i] % 360) for i in range(6)]
    data_dark = [lch_to_hex(L_STEPS_DARK[i], chroma[i], hues[i] % 360) for i in range(6)]
    return {
        "scene": scene, "sizes": SCENES[scene], "light": light, "dark": dark,
        "data": data, "data_dark": data_dark,
        "checks": {
            "light_ink_on_bg": round(M.contrast_ratio(light["ink"], light["bg"]), 1),
            "light_muted_on_bg": round(M.contrast_ratio(light["muted"], light["bg"]), 1),
            "light_accent_on_bg": round(M.contrast_ratio(light["accent"], light["bg"]), 1),
            "dark_ink_on_bg": round(M.contrast_ratio(dark["ink"], dark["bg"]), 1),
            "dark_accent_on_bg": round(M.contrast_ratio(dark["accent"], dark["bg"]), 1),
            "data_gray_conflicts": len(M.gray_conflicts(data)),
            "data_dark_gray_conflicts": len(M.gray_conflicts(data_dark)),
        },
    }


# ---------------------------------------------------------------- 网格与度量

W, H = 13.333, 7.5
MX, MY, GUT, COLS = 0.8, 0.7, 0.25, 12
COL_W = (W - 2 * MX - GUT * (COLS - 1)) / COLS


def col(start: int, span: int) -> Tuple[float, float]:
    """第 start 列（从 1 起）起、跨 span 列的 (x, w)。"""
    return MX + (start - 1) * (COL_W + GUT), span * COL_W + (span - 1) * GUT


class Ctx:
    def __init__(self, outline: Dict[str, Any]):
        design = outline.get("design") or {}
        self.system = design_system(design.get("seed", "#1F3A5F"), design.get("mood", "auto"),
                                    design.get("scene", "meeting"))
        self.sz = self.system["sizes"]
        self.fonts = design.get("fonts") or {"heading": "微软雅黑", "body": "微软雅黑"}
        self.stub = {"fonts": self.fonts}

    def pal(self, dark: bool) -> Dict[str, str]:
        return self.system["dark" if dark else "light"]

    def need(self, el: Dict[str, Any], w: float, bg: str) -> float:
        """框高 = 估算需求 × 1.08 + 0.02，向上取 0.05。折行按 94% 的宽度估：
        渲染器在中西文交界处加的间距各家不同，宁可框高一点，也不让文字落出框外。"""
        el = dict(el, x=0, y=0, w=w, h=1)
        st = M.text_style(self.stub, {"background": bg}, el)
        return math.ceil((M.fit_height_in(el, st, w * 0.94) * 1.08 + 0.02) * 20) / 20

    def text(self, eid: str, kind: str, x: float, y: float, w: float, content: str, bg: str,
             h: Optional[float] = None, **kw: Any) -> Dict[str, Any]:
        el: Dict[str, Any] = {"id": eid, "type": kind, "content": content, "pad_in": 0}
        el.update({k: v for k, v in kw.items() if v is not None})
        height = h if h is not None else self.need(el, w, bg)
        el.update(x=round(x, 3), y=round(y, 3), w=round(w, 3), h=round(height, 3))
        return el


# ---------------------------------------------------------------- 版式配方

def _image_el(spec: Dict[str, Any], x: float, y: float, w: float, h: float) -> Dict[str, Any]:
    el = {"id": "img", "type": "image", "image_id": spec["id"], "x": round(x, 3), "y": round(y, 3),
          "w": round(w, 3), "h": round(h, 3),
          "image_role": spec.get("role", "background"), "info_duty": spec.get("info_duty", "atmosphere"),
          "asset_reason": spec.get("reason", "氛围图，承担景深与情绪"),
          "crop_focus": spec.get("focus", {"x": 0.5, "y": 0.5})}
    if el["image_role"] in ("background", "texture"):
        el["overlap_allowed"] = True  # 只有承载层可以被压；其余角色的图不与文字相交
    if spec.get("duotone"):
        el["duotone"] = spec["duotone"]
    return el


def _title(c: Ctx, p: Dict[str, Any], dark: bool, span: int = 10) -> Dict[str, Any]:
    pal = c.pal(dark)
    x, w = col(1, span)
    return c.text("t", "title", x, MY, w, p["title"], pal["bg"], font_size=c.sz["title"], color=pal["ink"],
                  line_spacing=1.12)


def _note(c: Ctx, p: Dict[str, Any], dark: bool) -> List[Dict[str, Any]]:
    if not p.get("note"):
        return []
    pal = c.pal(dark)
    x, w = col(1, 9)
    return [c.text("note", "caption", x, H - 0.85, w, p["note"], pal["bg"], font_size=c.sz["caption"],
                   color=pal["muted"], h=0.3)]


def r_cover(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", True))
    pal = c.pal(dark)
    els: List[Dict[str, Any]] = []
    if p.get("image"):
        els.append(_image_el(p["image"], 0, 0, W, H))
    x, w = col(1, 8)
    t = c.text("t", "title", x, 0, w, p["title"], pal["bg"], font_size=c.sz["hero"], color=pal["ink"],
               line_spacing=1.08, letter_spacing_pt=-0.5)
    t["y"] = round(H * 0.62 - t["h"], 3)  # 标题底边落在 62% 水平线，上方留出整块安静区
    els.append(t)
    if p.get("sub"):
        els.append(c.text("sub", "caption", x, H * 0.62 + 0.3, w, p["sub"], pal["bg"], font_size=c.sz["body"] - 2,
                          color=pal["ink2"]))
    if p.get("kicker"):
        els.append(c.text("kicker", "caption", x, MY, w, p["kicker"], pal["bg"], font_size=c.sz["caption"],
                          color=pal["accent"], letter_spacing_pt=1.5, bold=True))
    return els, dark


def r_statement(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    x, w = col(1, 10)
    t = c.text("t", "title", x, 0, w, p["title"], pal["bg"], font_size=c.sz["hero"], color=pal["ink"],
               line_spacing=1.08)
    t["y"] = round((H - t["h"]) / 2 - 0.2, 3)
    els = [t]
    if p.get("sub"):
        els.append(c.text("sub", "caption", x, t["y"] + t["h"] + 0.3, w, p["sub"], pal["bg"],
                          font_size=c.sz["body"] - 2, color=pal["muted"]))
    return els, dark


def r_big_number(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    els = [_title(c, p, dark)]
    size = c.sz["number"] + 40
    unit = p.get("unit", "")
    num = str(p["number"])
    runs = [{"text": num}, {"text": unit, "font_size": round(size * float(p.get("unit_scale", 0.4)))}] if unit else None
    x, w = col(1, 6)
    n = c.text("n", "number", x, 0, w, num + unit, pal["bg"], font_size=size, color=pal["accent"], line_spacing=1.0,
               letter_spacing_pt=-2, runs=runs)
    n["y"] = round(H * 0.5 - n["h"] * 0.5 + 0.2, 3)
    els.append(n)
    bx, bw = col(8, 5)
    body = c.text("b", "body", bx, 0, bw, p["body"], pal["bg"], font_size=c.sz["body"], color=pal["ink2"])
    body["y"] = round(n["y"] + n["h"] * 0.5 - body["h"] * 0.5, 3)
    els.append(body)
    return els + _note(c, p, dark), dark


def r_chart(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    els = [_title(c, p, dark)]
    top = MY + els[0]["h"] + 0.35
    bottom = H - 1.0
    aside = p.get("aside")
    x, w = col(1, 8 if aside else 12)
    chart = dict(p["chart"])
    chart.setdefault("palette", c.system["data_dark" if dark else "data"])
    els.append({"id": "ch", "type": "chart", "x": round(x, 3), "y": round(top, 3), "w": round(w, 3),
                "h": round(bottom - top, 3), "chart": chart})
    if aside:
        ax, aw = col(10, 3)
        els.append(c.text("aside", "body", ax, top + 0.4, aw, aside, pal["bg"], font_size=c.sz["body"],
                          color=pal["ink2"]))
    return els + _note(c, p, dark), dark


def r_compare(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    title = _title(c, p, dark)
    els: List[Dict[str, Any]] = []
    block = 0.0
    for key, start, accent in (("left", 1, False), ("right", 7, True)):
        side = p[key]
        x, w = col(start, 5)
        label = c.text(f"{key}_label", "caption", x, 0, w, side["label"], pal["bg"], font_size=c.sz["caption"] + 1,
                       color=pal["muted"], bold=True, letter_spacing_pt=1)
        n = c.text(f"{key}_n", "number", x, label["h"] + 0.15, w, side["number"], pal["bg"], font_size=c.sz["number"],
                   color=pal["accent"] if accent else pal["ink2"], line_spacing=1.0, letter_spacing_pt=-2)
        b = c.text(f"{key}_b", "body", x, n["y"] + n["h"] + 0.1, w, side["body"], pal["bg"], font_size=c.sz["body"],
                   color=pal["ink2"])
        els += [label, n, b]
        block = max(block, b["y"] + b["h"])
    # 两栏同基线；整块在“标题下沿 → 页脚上沿”之间垂直居中，而不是贴着标题往下堆
    free_top, free_bottom = MY + title["h"] + 0.2, H - 1.0
    shift = free_top + max((free_bottom - free_top - block) / 2, 0.0)
    for el in els:
        el["y"] = round(el["y"] + shift, 3)
    return [title] + els + _note(c, p, dark), dark


def r_list(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    x, w = col(1, 5)
    t = c.text("t", "title", x, MY, w, p["title"], pal["bg"], font_size=c.sz["title"], color=pal["ink"],
               line_spacing=1.12)
    lx, lw = col(7, 6)
    lst = c.text("items", "body", lx, 0, lw, "\n".join(p["items"]), pal["bg"], font_size=c.sz["body"] + 4,
                 color=pal["ink"], list=p.get("list", "number"), space_after_pt=16)
    lst["y"] = round(max(MY, (H - lst["h"]) / 2 - 0.1), 3)
    return [t, lst] + _note(c, p, dark), dark


def r_table(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", False))
    els = [_title(c, p, dark)]
    top = MY + els[0]["h"] + 0.4
    x, w = col(1, 12)
    rows = len(p["table"]["rows"]) + 1
    els.append({"id": "tb", "type": "chart", "x": round(x, 3), "y": round(top, 3), "w": round(w, 3),
                "h": round(min(H - 1.2 - top, rows * 0.8), 3), "chart": {"kind": "table", "data": p["table"]}})
    return els + _note(c, p, dark), dark


def r_split_image(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    side = p.get("side", "right")
    img_x = W * 0.5 if side == "right" else 0.0
    els = [_image_el(dict(p["image"], role=p["image"].get("role", "subject")), img_x, 0, W * 0.5, H)]
    x, w = col(1 if side == "right" else 8, 5)
    t = c.text("t", "title", x, MY, w, p["title"], pal["bg"], font_size=c.sz["title"], color=pal["ink"],
               line_spacing=1.12)
    els.append(t)
    if p.get("body"):
        body = c.text("b", "body", x, 0, w, p["body"], pal["bg"], font_size=c.sz["body"] + 2, color=pal["ink2"])
        body["y"] = round(H - MY - body["h"], 3)  # 标题在上、正文压底：两端拉开，留出整块安静的中段
        els.append(body)
    return els, dark


def r_close(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", True))
    pal = c.pal(dark)
    x, w = col(1, 9)
    t = c.text("t", "title", x, MY + 0.6, w, p["title"], pal["bg"], font_size=c.sz["title"] + 6, color=pal["ink"],
               line_spacing=1.1)
    els = [t]
    if p.get("items"):
        els.append(c.text("items", "body", x, MY + 0.6 + t["h"] + 0.6, w, "\n".join(p["items"]), pal["bg"],
                          font_size=c.sz["body"] + 2, color=pal["ink2"], list=p.get("list", "number"),
                          space_after_pt=14))
    return els, dark


RECIPES = {
    "cover": (r_cover, "open", "标题落在 62% 水平线，上方留整块安静区"),
    "statement": (r_statement, "hold", "一句话居中偏上，其余全空，让人停下来"),
    "big_number": (r_big_number, "build", "关键数字压住页面，正文只写结论与限定"),
    "chart": (r_chart, "build", "图承担读数，旁注只写意义"),
    "compare": (r_compare, "turn", "两个数字同基线同尺度，落差就是论点"),
    "list": (r_list, "build", "标题在左固定，条目在右逐项展开"),
    "table": (r_table, "build", "表格承担精确读数"),
    "split_image": (r_split_image, "build", "半幅图承载空间，文字留在另一侧"),
    "close": (r_close, "close", "收在决定与下一步"),
}


def build_plan(outline: Dict[str, Any]) -> Dict[str, Any]:
    ctx = Ctx(outline)
    plan: Dict[str, Any] = {k: v for k, v in outline.items() if k not in ("design", "pages")}
    plan["fonts"] = ctx.fonts
    plan["palette"] = ctx.system["data"]
    pages, rhythm = [], []
    for i, p in enumerate(outline["pages"], 1):
        fn, role, grammar = RECIPES[p["recipe"]]
        elements, dark = fn(ctx, p)
        page: Dict[str, Any] = {
            "id": p.get("id", f"p{i}"), "role": p.get("role", role), "claim": p.get("claim", p["title"]),
            "focus": p.get("focus", p["title"].split("\n")[0]), "elements": elements,
            "composition_rationale": p.get("rationale", grammar),
            "background": ctx.pal(dark)["bg"],
        }
        for key in ("evidence", "open_questions", "speaker_notes", "page_number_tone"):
            if p.get(key):
                page[key] = p[key]
        pages.append(page)
        rhythm.append(p.get("beat") or f"{page['role']}：{grammar}")
    plan["pages"] = pages
    plan["rhythm"] = rhythm
    return plan


def main(argv: Optional[List[str]] = None) -> int:
    M.utf8_stdio()
    ap = argparse.ArgumentParser(description="设计系统与版式配方")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s1 = sub.add_parser("system", help="由种子色推导色彩系统与字号阶梯")
    s1.add_argument("--seed", required=True)
    s1.add_argument("--mood", default="auto", choices=["auto", "warm", "cool"])
    s1.add_argument("--scene", default="meeting", choices=list(SCENES))
    s2 = sub.add_parser("init", help="outline.json → plan.json 初稿")
    s2.add_argument("outline")
    s2.add_argument("out")
    s2.add_argument("--force", action="store_true")
    a = ap.parse_args(argv)
    if a.cmd == "system":
        print(json.dumps(design_system(a.seed, a.mood, a.scene), ensure_ascii=False, indent=2))
        return 0
    out = Path(a.out)
    if out.exists() and not a.force:
        print(f"[HARD] E-EXIST {out} 已存在；确认要替换才加 --force")
        return 2
    outline = json.loads(Path(a.outline).read_text(encoding="utf-8"))
    plan = build_plan(outline)
    out.write_text(json.dumps(plan, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"已写出 {out}（{len(plan['pages'])} 页）。下一步：compile.py → verify.py → 看图，再手改文案与几何。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
