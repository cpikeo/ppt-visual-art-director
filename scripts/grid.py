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
    dims = LIMITS["canvas"]
    parser = argparse.ArgumentParser(description="网格 → 英寸坐标")
    parser.add_argument("--aspect", choices=["16:9", "4:3"], default=dims["default"])
    parser.add_argument("--cols", type=int, default=12)
    parser.add_argument("--margin-x", type=float, default=0.7)
    parser.add_argument("--margin-y", type=float, default=0.6)
    parser.add_argument("--gutter", type=float, default=0.25)
    parser.add_argument("--corner", choices=["bottom_right", "top_right"], default=LIMITS["page_number"]["corner"])
    parser.add_argument("--span", action="append", default=[], metavar="起始列:列数")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    width, height = dims[args.aspect]
    grid = build(width, height, args.cols, args.margin_x, args.margin_y, args.gutter)
    grid["aspect"] = args.aspect
    grid["page_number"] = page_number(width, height, args.corner)
    grid["spans"] = [span(grid, *parse(s)) for s in args.span]

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
    box = grid["page_number"]["box_in"]
    print(f"\n页码保留框（{args.corner}）x={box['x']} y={box['y']} w={box['w']} h={box['h']}")
    print("任何元素都不得与页码框相交；右下要放来源时整稿改用 --corner top_right。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
