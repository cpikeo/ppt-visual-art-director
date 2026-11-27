#!/usr/bin/env python3
"""展示版 PPTX（不可编辑）：把 web_export.mjs 渲染的整页 PNG 按页序原样装进 16:9 幻灯片。

本脚本只做打包：PNG 字节不重采样、不重编码（python-pptx 以原字节嵌入，解包后 SHA-256 与基准一致）。
它不是第二个创作入口；输入只有 `slide-NN.png` 目录，输出是不可编辑的图像 PPTX。
用法: image_pptx.py PNG_DIR OUT.pptx [--title TEXT]
"""
import argparse, re, sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

SLIDE_W, SLIDE_H = 12192000, 6858000   # 13.333 × 7.5 in = 1280 × 720 px @ 96 dpi


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("png_dir"); ap.add_argument("out"); ap.add_argument("--title", default="Deck")
    a = ap.parse_args(argv)
    pngs = sorted((p for p in Path(a.png_dir).iterdir() if re.fullmatch(r"slide-\d+\.png", p.name)), key=lambda p: int(re.findall(r"\d+", p.name)[0]))
    if not pngs:
        print("ERROR 目录里没有 slide-NN.png", file=sys.stderr); return 2
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(SLIDE_W), Emu(SLIDE_H)
    blank = prs.slide_layouts[6]
    for i, png in enumerate(pngs, 1):
        slide = prs.slides.add_slide(blank)
        pic = slide.shapes.add_picture(str(png), 0, 0, width=Emu(SLIDE_W), height=Emu(SLIDE_H))
        pic.name = f"slide-{i:02d}-render"
        pic._element.nvPicPr.cNvPr.set("descr", f"第 {i} 页（HTML 终态渲染图像，不可编辑）")
    cp = prs.core_properties
    cp.title = a.title; cp.author = "ppt-visual-art-director"; cp.comments = "Image-only presentation: each slide is the final HTML render. Not editable."
    from datetime import datetime, timezone
    cp.created = cp.modified = datetime(2000, 1, 1, tzinfo=timezone.utc)   # 固定时间戳：同输入同字节
    prs.save(a.out)
    print(f"wrote {a.out} ({len(pngs)} slides)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
