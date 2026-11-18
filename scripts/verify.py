#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
核验交付物：plan ↔ PPTX ↔ 真实渲染。

用法：
    python scripts/verify.py plan.json out.pptx [--render-dir DIR] [--report report.json]

核验分三层，各自给出结论，不合并成一个分数：
  1. 对象核对   成品里的对象名、几何、文字、字号、墨色、图表载荷与着色、表格、图片字节是否与 plan 一致。
  2. 渲染核对   用 LibreOffice 把本次 PPTX 转 PDF，再由 PyMuPDF 取页尺寸、文字词框、字体与 PNG。
                缺工具时如实记 SKIP，并说明哪些判断因此无法做出。
  3. 需人判断   构图、主次、概念是否成立、目标 PowerPoint 是否正常——机器不给结论。

exit 0 = 没有阻断项；1 = 有；2 = 无法读取输入。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import compile as C  # noqa: E402  plan 模型与阈值共用一个实现；本文件只从成品读取“实际值”

EMU = C.EMU
TEXT_TYPES = C.TEXT_TYPES


class Findings:
    def __init__(self) -> None:
        self.items: List[C.Issue] = []

    def add(self, level: str, code: str, where: str, message: str) -> None:
        self.items.append(C.Issue(level, code, where, message))

    @property
    def blocked(self) -> bool:
        return any(i.level == "HARD" for i in self.items)

    def count(self, level: str) -> int:
        return sum(1 for i in self.items if i.level == level)

    def dump(self) -> List[Dict[str, str]]:
        return [{"level": i.level, "code": i.code, "where": i.where, "message": i.message} for i in self.items]


# ---------------------------------------------------------------- 对象核对


def _in(value: Optional[int]) -> Optional[float]:
    return None if value is None else round(value / EMU, 3)


def _norm(text: str) -> str:
    return "".join(text.split())


def verify_objects(plan: Dict[str, Any], pptx: Path, f: Findings) -> Dict[str, Any]:
    from pptx import Presentation
    from pptx.enum.dml import MSO_COLOR_TYPE

    prs = Presentation(str(pptx))
    pages = plan["pages"]
    w, h = C.canvas(plan)
    facts: Dict[str, Any] = {"slides": len(prs.slides), "expected_slides": len(pages)}
    if len(prs.slides) != len(pages):
        f.add("HARD", "V-PAGES", "整稿", f"成品 {len(prs.slides)} 页 ≠ plan {len(pages)} 页")

    for index, page in enumerate(pages):
        if index >= len(prs.slides):
            break
        pid = page["id"]
        slide = prs.slides[index]
        shapes = list(slide.shapes)
        by_name = {s.name: s for s in shapes}
        expected_order = [f"{pid}.{el['id']}" for el in C.ordered(page)]
        actual_order = [s.name for s in shapes if s.name != "_page_number"]
        if actual_order != expected_order:
            f.add("SOFT", "V-ORDER", pid, f"对象叠放顺序与声明不同：{actual_order} ≠ {expected_order}")

        for el in C.ordered(page):
            name = f"{pid}.{el['id']}"
            shape = by_name.get(name)
            if shape is None:
                f.add("HARD", "V-OBJ", name, "成品里找不到该对象")
                continue
            box = (_in(shape.left), _in(shape.top), _in(shape.width), _in(shape.height))
            want = (round(float(el["x"]), 3), round(float(el["y"]), 3), round(float(el["w"]), 3), round(float(el["h"]), 3))
            if any(abs(a - b) > 0.02 for a, b in zip(box, want)):
                f.add("HARD", "V-GEOM", name, f"几何 {box} ≠ plan {want}")
            kind = el["type"]
            if kind in TEXT_TYPES:
                actual = shape.text_frame.text
                if _norm(actual) != _norm(el["content"]):
                    f.add("HARD", "V-TEXT", name, f"文字与 plan 不一致：「{actual[:40]}」")
                style = C.text_style(plan, page, el)
                runs = [r for p in shape.text_frame.paragraphs for r in p.runs]
                if runs:
                    size = runs[0].font.size.pt if runs[0].font.size else None
                    if size is None or abs(size - style["size"]) > 0.5:
                        f.add("HARD", "V-STYLE", name, f"字号 {size} ≠ {style['size']}")
                    color = runs[0].font.color
                    if color is not None and color.type == MSO_COLOR_TYPE.RGB:
                        if f"#{color.rgb}".upper() != "#" + style["color"].lstrip("#").upper():
                            f.add("SOFT", "V-STYLE", name, f"墨色 #{color.rgb} ≠ 预期 {style['color']}")
            elif kind == "image":
                blob = getattr(shape, "image", None)
                if blob is None:
                    f.add("HARD", "V-OBJ", name, "该对象不是图片")
                else:
                    actual_hash = hashlib.sha256(shape.image.blob).hexdigest()
                    facts.setdefault("images", {})[name] = actual_hash
                    asset = (plan.get("assets") or {}).get(el["image_id"]) or {}
                    declared = (asset.get("generation") or {}).get("sha256")
                    if declared and declared.lower() != actual_hash:
                        f.add("HARD", "V-ASSET", name, "内嵌图片哈希与生成登记不符")
            elif kind == "chart":
                verify_chart(plan, page, el, shape, name, f)
            elif kind in ("shape", "divider"):
                color = shape.fill.fore_color
                if color.type == MSO_COLOR_TYPE.RGB and f"#{color.rgb}".upper() != "#" + str(el["color"]).lstrip("#").upper():
                    f.add("HARD", "V-STYLE", name, f"填充 #{color.rgb} ≠ {el['color']}")

        # 页码
        cfg = C.page_numbers(plan)
        if pid in cfg["hide"]:
            if "_page_number" in by_name:
                f.add("HARD", "V-PAGENO", pid, "声明隐藏页码，但成品里有页码对象")
        else:
            pn = by_name.get("_page_number")
            label = str(cfg["start"] + index)
            if pn is None:
                f.add("HARD", "V-PAGENO", pid, "缺少页码对象")
            elif _norm(pn.text_frame.text) != label:
                f.add("HARD", "V-PAGENO", pid, f"页码文字「{pn.text_frame.text}」≠「{label}」")
            else:
                box = C.pn_box(plan, page)
                pos = (_in(pn.left), _in(pn.top))
                if abs(pos[0] - box["x"]) > 0.02 or abs(pos[1] - box["y"]) > 0.02:
                    f.add("SOFT", "V-PAGENO", pid, f"页码位置 {pos} ≠ 保留框 ({box['x']},{box['y']})")
                runs = [r for p in pn.text_frame.paragraphs for r in p.runs]
                if runs and runs[0].font.size and abs(runs[0].font.size.pt - C.LIMITS["page_number"]["font_pt"]) > 0.5:
                    f.add("SOFT", "V-PAGENO", pid, f"页码字号 {runs[0].font.size.pt}pt 与整稿约定不符")
    return facts


def _ser_colors(shape) -> List[str]:
    """从成品 XML 读实际系列着色，不用 plan 的 palette 重算。"""
    from pptx.oxml.ns import qn
    out: List[str] = []
    for ser in shape.chart._chartSpace.iter(qn("c:ser")):
        spPr = ser.find(qn("c:spPr"))
        value = None
        if spPr is not None:
            fill = spPr.find(qn("a:solidFill"))
            if fill is not None:
                clr = fill.find(qn("a:srgbClr"))
                if clr is not None:
                    value = clr.get("val")
        out.append(value or "")
    return out


def verify_chart(plan, page, el, shape, name, f: Findings) -> None:
    from pptx.oxml.ns import qn
    chart_spec = el["chart"]
    if chart_spec["kind"] == "table":
        if not shape.has_table:
            f.add("HARD", "V-CHART", name, "该对象不是原生表格")
            return
        table = shape.table
        want_headers = [str(v) for v in chart_spec["data"]["headers"]]
        got_headers = [table.cell(0, c).text.strip() for c in range(len(table.columns))]
        if [_norm(x) for x in got_headers] != [_norm(x) for x in want_headers]:
            f.add("HARD", "V-CHART", name, f"表头 {got_headers} ≠ plan {want_headers}")
        if len(table.rows) != len(chart_spec["data"]["rows"]) + 1:
            f.add("HARD", "V-CHART", name, f"行数 {len(table.rows)} ≠ plan {len(chart_spec['data']['rows']) + 1}")
        return

    if not shape.has_chart:
        f.add("HARD", "V-CHART", name, "该对象不是原生图表")
        return
    obj = shape.chart
    plot = obj.plots[0]
    got_cats = [str(c) for c in getattr(plot, "categories", [])]
    want_cats = [str(c) for c in chart_spec["data"]["categories"]]
    if [_norm(x) for x in got_cats] != [_norm(x) for x in want_cats]:
        f.add("HARD", "V-CHART", name, f"类别 {got_cats} ≠ plan {want_cats}")
    for i, spec in enumerate(chart_spec["data"]["series"]):
        if i >= len(plot.series):
            f.add("HARD", "V-CHART", name, f"缺少系列「{spec['name']}」")
            continue
        got = [round(float(v), 6) for v in plot.series[i].values]
        want = [round(float(v), 6) for v in spec["values"]]
        if got != want:
            f.add("HARD", "V-CHART", name, f"系列「{spec['name']}」数值 {got} ≠ plan {want}")

    colors = [c for c in _ser_colors(shape) if c]
    bg = C.page_bg(page)
    floor = C.LIMITS["chart"]["min_contrast"].get(chart_spec["kind"], 2.0)
    for hex_color in colors:
        ratio = C.contrast_ratio("#" + hex_color.lstrip("#"), bg)
        if ratio < floor:
            f.add("SOFT", "V-CONTRAST", name, f"系列色 #{hex_color} 对底色 {bg} 仅 {ratio:.2f}:1")

    if chart_spec.get("highlight_category") is not None:
        pts = [d for d in obj._chartSpace.iter(qn("c:dPt"))]
        if not pts:
            f.add("SOFT", "V-CHART", name, "声明了 highlight_category，但成品没有逐点着色")
        else:
            fills = []
            for dpt in pts:
                spPr = dpt.find(qn("c:spPr"))
                val = None
                if spPr is not None:
                    solid = spPr.find(qn("a:solidFill"))
                    if solid is not None:
                        srgb = solid.find(qn("a:srgbClr"))
                        if srgb is not None:
                            val = srgb.get("val")
                fills.append(val or "")
            for idx, hex_value in enumerate(fills):
                if hex_value and hex_value in [c for c in colors]:
                    continue
                if hex_value:
                    ratio = C.contrast_ratio("#" + hex_value, bg)
                    if ratio < floor:
                        f.add("SOFT", "V-CONTRAST", name, f"退隐柱 #{hex_value} 对底色 {bg} 仅 {ratio:.2f}:1，读数会消失")


# ---------------------------------------------------------------- 渲染核对


def find_soffice() -> Optional[str]:
    from shutil import which
    for name in ("soffice", "libreoffice"):
        for ext in ("", ".exe"):
            found = which(name + ext)
            if found:
                return found
    for path in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                 "/usr/bin/soffice", "/usr/local/bin/soffice"):
        if Path(path).is_file():
            return path
    return None


def render(pptx: Path, render_dir: Path, f: Findings) -> Optional[Path]:
    soffice = find_soffice()
    if soffice is None:
        f.add("INFO", "V-RENDER", "环境", "找不到 LibreOffice；本页尺寸、断行、字体替换与溢出均未核验")
        return None
    try:
        import pymupdf  # noqa: F401
    except ImportError:
        f.add("INFO", "V-RENDER", "环境", "找不到 PyMuPDF（pip install pymupdf）；渲染核对跳过")
        return None
    render_dir.mkdir(parents=True, exist_ok=True)
    for stale in render_dir.glob("*.pdf"):
        if stale.stem == pptx.stem:
            stale.unlink()
    env = dict(os.environ)
    env.pop("PYTHONPATH", None)
    cmd = [soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", str(render_dir), str(pptx)]
    try:
        subprocess.run(cmd, capture_output=True, timeout=240, env=env, check=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        f.add("INFO", "V-RENDER", "环境", f"LibreOffice 调用失败：{exc}")
        return None
    pdf = render_dir / (pptx.stem + ".pdf")
    if not pdf.is_file():
        f.add("INFO", "V-RENDER", "环境", "LibreOffice 未产出 PDF（可能是文件被占用）")
        return None
    return pdf


def verify_render(plan: Dict[str, Any], pptx: Path, pdf: Path, render_dir: Path, f: Findings) -> Dict[str, Any]:
    import pymupdf
    from PIL import Image

    doc = pymupdf.open(str(pdf))
    w_in, h_in = C.canvas(plan)
    dpi = C.LIMITS["preview"]["dpi"]
    tol_pt = C.LIMITS["preview"]["overflow_tolerance_pt"]
    facts: Dict[str, Any] = {"pages": doc.page_count, "dpi": dpi, "pngs": []}
    if doc.page_count != len(plan["pages"]):
        f.add("HARD", "V-PAGES", "渲染", f"PDF {doc.page_count} 页 ≠ plan {len(plan['pages'])} 页")

    pngs: List[Path] = []
    fonts_seen: set = set()
    for index, pdf_page in enumerate(doc):
        spec_page = plan["pages"][index] if index < len(plan["pages"]) else {"id": f"p{index + 1}", "elements": []}
        pid = spec_page["id"]
        # 允许区域：文本元素框本身，以及图表/表格框（它们的轴标签、单元格在图内，不是独立文本框）
        boxes = [(float(el["x"]), float(el["y"]), float(el["w"]), float(el["h"]))
                 for el in spec_page["elements"] if el["type"] in TEXT_TYPES + ("chart",)]
        page_w = round(pdf_page.rect.width / 72, 3)
        page_h = round(pdf_page.rect.height / 72, 3)
        if abs(page_w - w_in) > 0.05 or abs(page_h - h_in) > 0.05:
            f.add("HARD", "V-PAGES", pid, f"页面尺寸 {page_w}×{page_h}in ≠ 画布 {w_in}×{h_in}in")
        for font in pdf_page.get_fonts():
            base = font[3].split("+")[-1]
            fonts_seen.add(base)

        pix = pdf_page.get_pixmap(dpi=dpi)
        png_path = render_dir / f"p{index + 1:02d}.png"
        pix.save(str(png_path))
        pngs.append(png_path)

        # 空白页：有效像素过少
        with Image.open(png_path) as img:
            gray = img.convert("L").resize((320, 180))
            pixels = gray.tobytes()
            ink = sum(1 for v in pixels if v < 245) / max(len(pixels), 1)
            if ink < C.LIMITS["preview"]["blank_ink_ratio"]:
                f.add("SOFT", "V-BLANK", pid, f"页面近乎空白（墨迹占比 {ink:.4%}）")

        # 文字越出画布 / 越出所有声明框
        page_box = (0.0, 0.0, page_w, page_h)
        for block in pdf_page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span["text"].strip()
                    if not text:
                        continue
                    x0, y0, x1, y1 = span["bbox"]
                    sx0, sy0, sx1, sy1 = x0 / 72, y0 / 72, x1 / 72, y1 / 72
                    if sx0 < -0.02 or sy0 < -0.02 or sx1 > page_box[2] + 0.02 or sy1 > page_box[3] + 0.02:
                        f.add("HARD", "V-CLIP", pid, f"文字超出页面：{text[:24]}")
                        continue
                    # 容差随字号放大：一行字的下伸与行距留白本来就落在框外一点，不是溢出
                    tol = max(0.06, 0.15 * (sy1 - sy0))
                    if any(sx0 >= bx - tol and sy0 >= by - tol and sx1 <= bx + bw + tol and sy1 <= by + bh + tol
                           for bx, by, bw, bh in boxes):
                        continue
                    if text.isdigit() and len(text) <= 3:
                        continue  # 页码由 V-PAGENO 单独查
                    f.add("SOFT", "V-OVERFLOW", pid, f"「{text[:20]}」不在任何声明框内（可能溢出或被替换字体撑开）")

        # 页码是否真的落在角区
        cfg = C.page_numbers(plan)
        if pid not in cfg["hide"]:
            want = str(cfg["start"] + index)
            box = C.pn_box(plan, spec_page)
            hit = False
            for word in pdf_page.get_text("words"):
                if word[4].strip() == want:
                    cx, cy = (word[0] + word[2]) / 2 / 72, (word[1] + word[3]) / 2 / 72
                    if box["x"] - 0.35 <= cx <= box["x"] + box["w"] + 0.35 and \
                       box["y"] - 0.2 <= cy <= box["y"] + box["h"] + 0.2:
                        hit = True
                        break
            if not hit:
                f.add("HARD", "V-PAGENO", pid, f"渲染里角区找不到页码「{want}」")

    facts["fonts"] = sorted(fonts_seen)
    # 字族一致性：中文族名与 PDF 里的拉丁名无法逐字对应，改为比较“出现过的字族数量”。
    # 声明一套字族却渲染出两套，说明某处被替代或回退——这正是混排断掉的地方。
    bases = {font_family_base(name) for name in fonts_seen}
    requested = {(plan.get("fonts") or {}).get(key) for key in ("heading", "body")}
    requested |= {el.get("font") for page in plan["pages"] for el in page["elements"] if el.get("font")}
    requested.discard(None)
    if len(bases) > max(len(requested), 1):
        f.add("SOFT", "V-FONT", "渲染",
              f"渲染里出现 {len(bases)} 套字族 {sorted(bases)}，计划只有 {sorted(requested)}；"
              f"某处被替换或回退，断行与字宽不可信")

    facts["pngs"] = make_contact_sheet(pngs, render_dir)
    return facts


def font_family_base(name: str) -> str:
    base = name.split("+")[-1]
    for suffix in ("-BoldItalic", "-BoldOblique", "-Bold", "-Italic", "-Oblique", "-Regular",
                   "-Light", "-Medium", "-Semibold", "-Black", "-Heavy", "-Thin"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
    return base.lower().replace(" ", "")


def make_contact_sheet(pngs: List[Path], render_dir: Path) -> List[str]:
    from PIL import Image
    if not pngs:
        return []
    try:
        with Image.open(pngs[0]) as probe:
            ratio = probe.height / probe.width
    except OSError:
        return [str(p) for p in pngs]
    cols = 2 if len(pngs) <= 4 else 3
    thumb_w = C.LIMITS["preview"]["contact_sheet_width_px"] // cols
    thumb_h = int(thumb_w * ratio)
    rows = (len(pngs) + cols - 1) // cols
    pad = 24
    sheet = Image.new("RGB", (cols * thumb_w + pad * (cols + 1), rows * thumb_h + pad * (rows + 1)), "#E9E6DF")
    for i, path in enumerate(pngs):
        with Image.open(path) as img:
            thumb = img.convert("RGB").resize((thumb_w, thumb_h), Image.LANCZOS)
        x = pad + (i % cols) * (thumb_w + pad)
        y = pad + (i // cols) * (thumb_h + pad)
        sheet.paste(thumb, (x, y))
    out = render_dir / "contact_sheet.png"
    sheet.save(out)
    return [str(out)] + [str(p) for p in pngs]


# ---------------------------------------------------------------- 主流程


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="核验 plan ↔ PPTX ↔ 渲染")
    parser.add_argument("plan")
    parser.add_argument("pptx")
    parser.add_argument("--render-dir", default=None, help="写 PDF/PNG/缩略总览的目录")
    parser.add_argument("--report", default=None)
    args = parser.parse_args(argv)

    plan_path = Path(args.plan).resolve()
    pptx = Path(args.pptx).resolve()
    if not pptx.is_file():
        print(f"找不到 {pptx}")
        return 2
    rep = C.Report()
    plan = C.load_plan(plan_path, rep)
    if plan is None:
        for issue in rep.items:
            print(issue)
        return 2

    f = Findings()
    key = pptx.read_bytes()
    sha = hashlib.sha256(key).hexdigest()
    facts: Dict[str, Any] = {"pptx": str(pptx), "sha256": sha,
                             "bytes": len(key), "render": {"done": False}}
    facts.update(verify_objects(plan, pptx, f))

    render_dir = Path(args.render_dir).resolve() if args.render_dir else pptx.parent / "render"
    pdf = render(pptx, render_dir, f)
    if pdf is not None:
        facts["render"] = {"done": True, "pdf": str(pdf)}
        facts["render"].update(verify_render(plan, pptx, pdf, render_dir, f))

    for issue in f.items:
        print(issue)

    status = "BLOCK" if f.blocked else "READY"
    print(f"\n对象与渲染核验：{status}   硬问题 {f.count('HARD')}｜设计提示 {f.count('SOFT')}｜说明 {f.count('INFO')}")
    print(f"PPTX SHA-256：{sha}")
    if facts["render"].get("done"):
        print(f"渲染：{facts['render']['pdf']}")
        print(f"缩略总览：{render_dir / 'contact_sheet.png'}")
        print("机器能证明的到此为止：构图、主次、概念是否成立，必须自己看 PNG 并在目标 PowerPoint 复核。")
    else:
        print("未渲染：视觉相关判断一律没有依据。不要声称「已预览」「版面无溢出」「断行正常」。")

    if args.report:
        Path(args.report).write_text(json.dumps({"status": status, **facts, "findings": f.dump()},
                                                ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if not f.blocked else 1


if __name__ == "__main__":
    sys.exit(main())
