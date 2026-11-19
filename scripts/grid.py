#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
网格 → plan 用的英寸坐标。只做算术，不决定用几列、哪里留白。

用法：
    python scripts/grid.py                          # 16:9 · 12 列 · 默认边距
    python scripts/grid.py --aspect 4:3 --cols 8
    python scripts/grid.py --span 1:5 --span 7:6    # 左 5 列 + 空 1 列 + 右 6 列
    python scripts/grid.py --json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

LIMITS = json.loads((Path(__file__).resolve().parent / "limits.json").read_text(encoding="utf-8-sig"))


def build(width: float, height: float, cols: int, mx: float, my: float, gutter: float) -> Dict[str, Any]:
    usable = width - 2 * mx
    col_w = (usable - gutter * (cols - 1)) / cols
    if col_w <= 0:
        raise SystemExit("边距或沟槽过大：单列宽度 ≤ 0")
    columns = [{"col": i + 1, "x": round(mx + i * (col_w + gutter), 3), "w": round(col_w, 3)} for i in range(cols)]
    return {"canvas_in": {"w": width, "h": height},
            "safe_in": {"x": mx, "y": my, "w": round(usable, 3), "h": round(height - 2 * my, 3)},
            "col_w": round(col_w, 3), "gutter": gutter, "columns": columns}


def span(grid: Dict[str, Any], start: int, count: int) -> Dict[str, float]:
    cols = grid["columns"]
    if start < 1 or count < 1 or start + count - 1 > len(cols):
        raise SystemExit(f"--span {start}:{count} 超出 1..{len(cols)}")
    first, last = cols[start - 1], cols[start + count - 2]
    return {"cols": f"{start}..{start + count - 1}", "x": first["x"],
            "w": round(last["x"] + last["w"] - first["x"], 3)}


def split(grid: Dict[str, Any], ratios: List[float], gutter: float) -> List[Dict[str, Any]]:
    """按比例切分版心。非对称构图的骨架是比例（3:5、2:1、1:1.618），不是等分列。"""
    if any(r <= 0 for r in ratios) or len(ratios) < 2:
        raise SystemExit(f"--split 需要两个以上正数比例，例如 3,5；收到 {ratios}")
    safe = grid["safe_in"]
    free = safe["w"] - gutter * (len(ratios) - 1)
    if free <= 0:
        raise SystemExit("比例切分后版心为负：减少栏数或沟槽")
    total = sum(ratios)
    out, x = [], safe["x"]
    for i, ratio in enumerate(ratios):
        width = free * ratio / total
        out.append({"index": i + 1, "ratio": ratio, "x": round(x, 3), "w": round(width, 3)})
        x += width + gutter
    return out


def guides(width: float, height: float, kind: str) -> Dict[str, List[float]]:
    """比例参考线。三分与黄金位是空间的骨架：先知道线在哪，再决定东西放不放上去。"""
    golden = 0.618
    if kind == "thirds":
        xs = [width / 3, width * 2 / 3]
        ys = [height / 3, height * 2 / 3]
    elif kind == "golden":
        xs = [width * (1 - golden), width * golden]
        ys = [height * (1 - golden), height * golden]
    else:
        raise SystemExit(f"--guides 取 thirds 或 golden，收到 {kind!r}")
    return {"x": [round(v, 3) for v in xs], "y": [round(v, 3) for v in ys]}


def page_number(width: float, height: float, corner: str) -> Dict[str, Any]:
    cfg = LIMITS["page_number"]
    bw, bh = cfg["box_in"]["w"], cfg["box_in"]["h"]
    mx, my = cfg["margin_in"]["x"], cfg["margin_in"]["y"]
    return {"corner": corner, "box_in": {"x": round(width - mx - bw, 3),
                                         "y": round(my if corner == "top_right" else height - my - bh, 3),
                                         "w": bw, "h": bh}}


def parse(text: str):
    try:
        start, count = text.split(":")
        return int(start), int(count)
    except ValueError:
        raise SystemExit(f"--span 应为 起始列:列数，例如 1:5；收到 {text!r}")


def main(argv: Optional[List[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass
    dims = LIMITS["canvas"]
    parser = argparse.ArgumentParser(description="网格 → 英寸坐标")
    parser.add_argument("--aspect", choices=["16:9", "4:3"], default=dims["default"])
    parser.add_argument("--cols", type=int, default=12)
    parser.add_argument("--margin-x", type=float, default=0.7)
    parser.add_argument("--margin-y", type=float, default=0.6)
    parser.add_argument("--gutter", type=float, default=0.25)
    parser.add_argument("--corner", choices=["bottom_right", "top_right"], default=LIMITS["page_number"]["corner"])
    parser.add_argument("--span", action="append", default=[], metavar="起始列:列数")
    parser.add_argument("--split", default=None, metavar="3,5", help="按比例切分版心，非对称构图的骨架")
    parser.add_argument("--guides", choices=["thirds", "golden"], default=None, help="三分线或黄金位坐标")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    width, height = dims[args.aspect]
    grid = build(width, height, args.cols, args.margin_x, args.margin_y, args.gutter)
    grid["aspect"] = args.aspect
    grid["page_number"] = page_number(width, height, args.corner)
    grid["spans"] = [span(grid, *parse(s)) for s in args.span]
    if args.split:
        try:
            ratios = [float(v) for v in args.split.replace(":", ",").split(",") if v.strip()]
        except ValueError:
            raise SystemExit(f"--split 应为 3,5 这样的比例序列，收到 {args.split!r}")
        grid["split"] = split(grid, ratios, args.gutter)
    if args.guides:
        grid["guides"] = guides(width, height, args.guides)

    if args.json:
        print(json.dumps(grid, ensure_ascii=False, indent=2))
        return 0

    safe = grid["safe_in"]
    print(f"画布 {width} × {height} in（{args.aspect}）　安全区 x={safe['x']} y={safe['y']} "
          f"w={safe['w']} h={safe['h']}")
    print(f"{args.cols} 列，列宽 {grid['col_w']}，沟槽 {args.gutter}\n")
    for col in grid["columns"]:
        print(f"  列{col['col']:>2}  x={col['x']:<7} w={col['w']}")
    if grid["spans"]:
        print("\n跨列")
        for item in grid["spans"]:
            print(f"  {item['cols']:>8}  x={item['x']}  w={item['w']}")
    if grid.get("split"):
        ratios = ":".join(f"{c['ratio']:g}" for c in grid["split"])
        print(f"\n按比例切分（{ratios}，沟槽 {args.gutter}）")
        for item in grid["split"]:
            print(f"  第{item['index']}栏  x={item['x']:<7} w={item['w']}")
    if grid.get("guides"):
        g = grid["guides"]
        label = "三分线" if args.guides == "thirds" else "黄金位"
        print(f"\n{label}：" + "｜".join([f"x={v}" for v in g["x"]] + [f"y={v}" for v in g["y"]]))
    box = grid["page_number"]["box_in"]
    print(f"\n页码保留框（{args.corner}）x={box['x']} y={box['y']} w={box['w']} h={box['h']}")
    print("任何元素都不得与页码框相交；右下要放来源时整稿改用 --corner top_right。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
