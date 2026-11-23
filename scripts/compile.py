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
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from lxml import etree

from charts import add_chart, add_legend
from checks import design_signals, validate
from model import EMU, guard_paths, LIMITS, Report, TEXT_TYPES, _crop_geometry, canvas, default_asset_root, explicit_box, load_plan, package_problems, parse_xml, page_bg, page_geometry, page_numbers, paragraph_runs, pn_box, pn_ink, qn, resolve_asset, resolve_points, text_style, utf8_stdio, alt_text
from ooxml import _apply_duotone, _apply_fill_alpha, _apply_gradient, _apply_stroke, _set_corner_radius, _set_run_font, _strip_style


# ---------------------------------------------------------------- 建对象


def _claim_title_placeholder(slide, page, el, top_level):
    """页内第一个顶层 title 元素落到版式的标题占位符上：大纲视图、读屏与幻灯片导航才认得出页标题。
    其余 title（或组内的）仍是普通文本框。位置与样式全部显式写出，不继承母版。"""
    if el.get("type") != "title" or not top_level:
        return None
    holder = slide.shapes.title
    if holder is None or holder.name.startswith(f"{page['id']}."):
        return None
    return holder

def add_text(slide, plan, page, el, top_level=True):
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

    st = text_style(plan, page, el)
    box = explicit_box(el)
    frame = _claim_title_placeholder(slide, page, el, top_level)
    if frame is not None:
        frame.left, frame.top, frame.width, frame.height = Inches(box[0]), Inches(box[1]), Inches(box[2]), Inches(box[3])
    else:
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


def build_element(slide, plan, page, el, boxes, asset_root: Path, top_level: bool = True):
    kind = el["type"]
    obj = None
    if kind in TEXT_TYPES:
        obj = add_text(slide, plan, page, el, top_level)
    elif kind in ("shape", "divider"):
        obj = add_shape(slide, page, el)
    elif kind == "path":
        obj = add_path(slide, page, el, boxes)
    elif kind == "image":
        obj = add_image(slide, page, el, asset_root)
    elif kind == "chart":
        obj = add_chart(slide, plan, page, el)
    elif kind == "legend":
        obj = add_legend(slide, plan, page, el)
    elif kind == "group":
        children = [build_element(slide, plan, page, child, boxes, asset_root, False) for child in el["elements"]]
        obj = slide.shapes.add_group_shape([c for c in children if c is not None])
        obj.name = f"{page['id']}.{el['id']}"
    else:
        raise ValueError(f"未知元素类型：{kind}")
    # 旋转在出口处统一应用：只有一处写 rot，也就不可能出现"某种类型忘了转"。
    # 角度先归一到 [0,360)：270 与 -90 是同一件事，不该因为写法不同被 schema 拦下。
    if obj is not None and el.get("rotation"):
        obj.rotation = float(el["rotation"]) % 360
    # 替代文字写在 cNvPr@descr：读屏与 PowerPoint「查看替代文字」都读这里。
    alt = alt_text(el)
    if obj is not None and alt:
        for node in obj._element.xpath(".//p:cNvPr")[:1]:
            node.set("descr", alt)
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

def _fit_master_to_canvas(prs, old_w: int) -> None:
    """python-pptx 默认模板是 4:3（10 in 宽）。只改页面尺寸不改母版时，母版与版式里的占位符仍是 4:3 几何：
    用户在 PowerPoint 里新建页、重置版式，标题与页码会落在页面左侧 3/4。这里把横向几何按比例铺到新宽度。"""
    ratio = prs.slide_width / float(old_w)
    if abs(ratio - 1.0) < 1e-6:
        return
    containers = [prs.slide_master] + list(prs.slide_layouts)
    for container in containers:
        for shape in container.shapes:
            # 只缩放自带 xfrm 的形状；占位符的继承几何读出来是母版的值，再缩放一次会放大两遍。
            if shape._element.find(".//" + qn("a:xfrm")) is None or shape.width is None:
                continue
            shape.left, shape.width = int(round(shape.left * ratio)), int(round(shape.width * ratio))


def _set_theme_fonts(prs, plan: Dict[str, Any]) -> None:
    """主题字体改成 plan 的字族（拉丁与东亚），用户后续新增的文字继承它，而不是 Calibri 加空的东亚字体。"""
    fonts = plan.get("fonts") or {}
    heading = fonts.get("heading") or LIMITS["defaults"]["fonts"]["heading"]
    body = fonts.get("body") or LIMITS["defaults"]["fonts"]["body"]
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    theme_part = prs.slide_master.part.part_related_by(RT.THEME)
    root = parse_xml(theme_part.blob)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    for tag, family in (("majorFont", heading), ("minorFont", body)):
        node = root.find(f".//a:fontScheme/a:{tag}", ns)
        if node is None:
            continue
        for kind in ("latin", "ea"):
            child = node.find(f"a:{kind}", ns)
            if child is not None:
                child.set("typeface", str(family))
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build(plan: Dict[str, Any], asset_root: Path, out: Path) -> None:
    from pptx import Presentation
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    w, h = canvas(plan)
    prs = Presentation()
    old_w = prs.slide_width
    prs.slide_width = Inches(w)
    prs.slide_height = Inches(h)
    _fit_master_to_canvas(prs, old_w)
    _set_theme_fonts(prs, plan)
    # “Title Only” 版式提供标题占位符；没有 title 元素的页会删掉空占位符，不留“单击添加标题”。
    layout = next((l for l in prs.slide_layouts if l.name == "Title Only"), prs.slide_layouts[5])

    for idx, page in enumerate(plan["pages"]):
        slide = prs.slides.add_slide(layout)
        background = slide.background.fill
        background.solid()
        background.fore_color.rgb = RGBColor.from_string(page_bg(page).lstrip("#"))
        boxes = page_geometry(page)
        built = [build_element(slide, plan, page, el, boxes, asset_root) for el in page["elements"]]
        holder = slide.shapes.title
        if holder is not None and not holder.name.startswith(f"{page['id']}."):
            holder._element.getparent().remove(holder._element)
        _reorder(slide, [b for b in built if b is not None])
        add_page_number(slide, plan, page, idx)
        if page.get("speaker_notes"):
            slide.notes_slide.notes_text_frame.text = page["speaker_notes"]

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
    asset_root = default_asset_root(plan_path, args.assets)

    rep = Report()
    report_path = Path(args.report).resolve() if args.report else None
    # 源文件与交付物分家：路径判定统一走 model.guard_paths（输出/报告不得覆盖 plan、素材，报告只能是 .json）。
    guard_paths(rep, {"plan": plan_path}, [("输出文件", Path(args.out), ".pptx"), ("--report", Path(args.report) if args.report else None, ".json")], args.force)
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
        note = "诊断文件：" + str(error_path) + "（不是交付物）"
        if error_path.exists() and not args.force:
            note = f"诊断文件 {error_path} 已存在，未覆盖（加 --force 才写）"
        else:
            write_error_deck(error_path, f"{type(exc).__name__}: {exc}")
        print(f"[HARD] E-RUNTIME 运行期失败：{type(exc).__name__}: {exc}\n{note}")
        return 3

    problems = package_problems(out_path)
    if problems:
        print(f"[HARD] E-FORMAT 成品包不合格：{'；'.join(problems)}")
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
