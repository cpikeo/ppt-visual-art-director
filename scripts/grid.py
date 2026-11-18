#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt-visual-art-director · 网格与坐标助手

用途：plan.json 里的 x / y / w / h 单位是【英寸】，原点在页面【左上角】。本工具把
“列数 / 页边距 / 沟槽”换算成可直接粘贴的坐标，并标出页码保留区，避免凭感觉填数。

它只做算术，不做版式决策：用几列、跨几列、哪里留白，仍由艺术指导判断
（见 references/layout-system.md）。

用法示例：
    python scripts/grid.py                          # 16:9，12 列，默认边距
    python scripts/grid.py --aspect 4:3 --cols 8
    python scripts/grid.py --span 1:5 --span 7:6    # 左 5 列 + 右 6 列（中间留 1 列空档）
    python scripts/grid.py --json                   # 机器可读输出
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

POLICY_PATH = Path(__file__).resolve().with_name("policy.json")


def load_policy() -> Dict[str, Any]:
    try:
        return json.loads(POLICY_PATH.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise SystemExit(f"[FATAL] 无法读取 policy.json：{exc}")


def build_grid(width: float, height: float, cols: int, margin_x: float, margin_y: float,
               gutter: float) -> Dict[str, Any]:
    if cols < 1:
        raise SystemExit("列数必须 ≥ 1")
    usable_w = width - 2 * margin_x
    col_w = (usable_w - gutter * (cols - 1)) / cols
    if col_w <= 0:
        raise SystemExit("边距或沟槽过大：单列宽度 ≤ 0")
    columns = [{"col": i + 1, "x": round(margin_x + i * (col_w + gutter), 3), "w": round(col_w, 3)}
               for i in range(cols)]
    return {
        "canvas_in": {"w": width, "h": height},
        "safe_area_in": {"x": margin_x, "y": margin_y, "w": round(usable_w, 3),
                         "h": round(height - 2 * margin_y, 3)},
        "cols": cols, "col_w": round(col_w, 3), "gutter": gutter, "columns": columns,
    }


def span(grid: Dict[str, Any], start: int, count: int) -> Dict[str, float]:
    cols = grid["columns"]
    if start < 1 or count < 1 or start + count - 1 > len(cols):
        raise SystemExit(f"跨列 {start}:{count} 超出 1..{len(cols)} 列")
    first, last = cols[start - 1], cols[start + count - 2]
    return {"start": start, "count": count, "x": first["x"], "w": round(last["x"] + last["w"] - first["x"], 3)}


def page_number_zone(policy: Dict[str, Any], width: float, height: float, corner: str) -> Dict[str, Any]:
    cfg = policy["page_number"]
    w, h = float(cfg["box_in"]["w"]), float(cfg["box_in"]["h"])
    mx, my = float(cfg["margin_in"]["x"]), float(cfg["margin_in"]["y"])
    x = width - mx - w
    y = my if corner == "top_right" else height - my - h
    zone = cfg["corner_zone"]
    zw = width * float(zone["x_from_right_ratio"])
    zh = height * float(zone["y_from_edge_ratio"])
    return {
        "corner": corner,
        "page_number_box_in": {"x": round(x, 3), "y": round(y, 3), "w": w, "h": h},
        "verifier_corner_zone_in": {
            "x": round(width - zw, 3),
            "y": 0.0 if corner == "top_right" else round(height - zh, 3),
            "w": round(zw, 3), "h": round(zh, 3),
        },
        "note": "页码由编译器自动生成：任何元素都不得与页码框相交；角区内与页码框不相交的其他文字不算页码（contract §6），但应避免拥挤。",
    }


def parse_span(text: str) -> Tuple[int, int]:
    try:
        start, count = text.split(":")
        return int(start), int(count)
    except ValueError:
        raise SystemExit(f"--span 格式应为 起始列:列数，例如 1:5；收到 {text!r}")


def main(argv: Optional[List[str]] = None) -> int:
    policy = load_policy()
    dims = policy["canvas"]["dimensions_in"]
    parser = argparse.ArgumentParser(description="plan 坐标（英寸）网格助手")
    parser.add_argument("--aspect", choices=sorted(dims), default=policy["canvas"]["default_aspect_ratio"])
    parser.add_argument("--cols", type=int, default=12)
    parser.add_argument("--margin-x", type=float, default=0.7, help="左右页边距（英寸）")
    parser.add_argument("--margin-y", type=float, default=0.6, help="上下页边距（英寸）")
    parser.add_argument("--gutter", type=float, default=0.25, help="列沟槽（英寸）")
    parser.add_argument("--corner", choices=["bottom_right", "top_right"],
                        default=policy["page_number"]["corner"], help="整稿页码角区")
    parser.add_argument("--span", action="append", default=[], metavar="起始列:列数")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    args = parser.parse_args(argv)

    width, height = dims[args.aspect]
    grid = build_grid(width, height, args.cols, args.margin_x, args.margin_y, args.gutter)
    grid["aspect_ratio"] = args.aspect
    grid["page_number"] = page_number_zone(policy, width, height, args.corner)
    grid["spans"] = [span(grid, *parse_span(item)) for item in args.span]

    if args.json:
        print(json.dumps(grid, ensure_ascii=False, indent=2))
        return 0

    sa = grid["safe_area_in"]
    print(f"画布 {width} × {height} in（{args.aspect}）；原点左上角；单位英寸")
    print(f"安全区 x={sa['x']} y={sa['y']} w={sa['w']} h={sa['h']}；{args.cols} 列，列宽 {grid['col_w']}，沟槽 {args.gutter}")
    print("\n列  x       w")
    for c in grid["columns"]:
        print(f"{c['col']:>2}  {c['x']:<7} {c['w']}")
    if grid["spans"]:
        print("\n跨列")
        for s in grid["spans"]:
            print(f"  列 {s['start']}..{s['start'] + s['count'] - 1}: x={s['x']} w={s['w']}")
    pn = grid["page_number"]
    box, zone = pn["page_number_box_in"], pn["verifier_corner_zone_in"]
    print(f"\n页码保留框（{pn['corner']}）x={box['x']} y={box['y']} w={box['w']} h={box['h']}")
    print(f"验证器角区           x={zone['x']} y={zone['y']} w={zone['w']} h={zone['h']}")
    print(pn["note"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
