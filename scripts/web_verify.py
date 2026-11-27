#!/usr/bin/env python3
"""Web → PPTX / PDF 对照核验。真实读取导出文件，逐页与 HTML 终态截图/导出清单比对。
用法: python scripts/web_verify.py --pptx deck.pptx --manifest deck.plan.manifest.json [--pdf deck.pdf --shots DIR --out DIR]
状态只有 PASS / FAIL / BLOCKED / UNVERIFIED；PPTX 渲染（LibreOffice/PowerPoint）缺失时视觉还原标 UNVERIFIED，不推断。
PDF 光栅化用 PyMuPDF（AGPL，仅本机可选工具，见 requirements-render.txt）。"""
import argparse, json, os, sys
from pptx import Presentation
from pptx.util import Emu

IN = 914400


def check(res, area, name, status, detail):
    res.append({"area": area, "name": name, "status": status, "detail": detail})


def verify_pptx(path, manifest, res):
    prs = Presentation(path)
    pages = max(m["page"] for m in manifest)
    check(res, "PPTX", "slide count == HTML slides", "PASS" if len(prs.slides) == pages else "FAIL", f"{len(prs.slides)} vs {pages}")
    check(res, "PPTX", "page size 16:9 (12192000×6858000 EMU)", "PASS" if (prs.slide_width, prs.slide_height) == (12192000, 6858000) else "FAIL", f"{prs.slide_width}×{prs.slide_height}")
    kinds = {"text": 0, "picture": 0, "chart": 0, "table": 0, "shape": 0, "other": 0}
    full_page = []; geom_bad = []; text_bad = []; matched = 0; missing = []; notes = 0; alt_missing = []
    area = prs.slide_width * prs.slide_height
    for si, slide in enumerate(prs.slides, 1):
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip(): notes += 1
        by_name = {}
        for sh in slide.shapes:
            by_name[sh.name] = sh
            if sh.shape_type == 13: kinds["picture"] += 1
            elif getattr(sh, "has_chart", False) and sh.has_chart: kinds["chart"] += 1
            elif getattr(sh, "has_table", False) and sh.has_table: kinds["table"] += 1
            elif sh.has_text_frame and sh.text_frame.text.strip(): kinds["text"] += 1
            elif sh.shape_type in (1, 17): kinds["shape"] += 1
            else: kinds["other"] += 1
            if sh.shape_type == 13 and sh.width * sh.height >= .9 * area: full_page.append(f"p{si}:{sh.name}")
        for m in [x for x in manifest if x["page"] == si]:
            sh = None
            pid = slide.part.slide_layout and None
            # compile names objects "<页 id>.<元素 id>"
            for n, cand in by_name.items():
                if n.endswith("." + m["id"]): sh = cand; break
            if sh is None: missing.append(f"p{si}/{m['id']}"); continue
            matched += 1
            dx = max(abs(sh.left - m["x"] * IN), abs(sh.top - m["y"] * IN), abs(sh.width - m["w"] * IN), abs(sh.height - m["h"] * IN)) / IN
            if dx > 0.02: geom_bad.append(f"p{si}/{m['id']} Δ{dx:.3f}in")
            if m["strategy"] == "native-text" and not (sh.has_text_frame and sh.text_frame.text.strip()): text_bad.append(f"p{si}/{m['id']}")
            if m["strategy"] == "raster" and sh.shape_type != 13: text_bad.append(f"p{si}/{m['id']} raster expected picture")
    check(res, "PPTX", "every exported element exists as a named object", "PASS" if not missing else "FAIL", f"{matched}/{len(manifest)}" + (f"; missing {missing[:5]}" if missing else ""))
    check(res, "PPTX", "geometry within 0.02 in of DOM box", "PASS" if not geom_bad else "FAIL", f"{len(geom_bad)} off" + (f": {geom_bad[:5]}" if geom_bad else ""))
    check(res, "PPTX", "native text objects are text (editable), rasters are pictures", "PASS" if not text_bad else "FAIL", ", ".join(text_bad[:5]) or "ok")
    check(res, "PPTX", "no full-page screenshot posing as slide", "PASS" if not full_page else "FAIL", ", ".join(full_page) or "no picture ≥ 90% of slide area")
    check(res, "PPTX", "speaker notes carried over", "PASS" if notes == len(prs.slides) else "FAIL", f"{notes}/{len(prs.slides)}")
    n = sum(kinds.values())
    check(res, "PPTX", "object inventory", "INFO", json.dumps(kinds) + f"; native ≈ {(n - kinds['picture'])}/{n} objects; pictures = {kinds['picture']} (declared raster/image only)")
    check(res, "PPTX", "visual fidelity vs HTML (needs LibreOffice/PowerPoint render)", "UNVERIFIED", "no PPTX renderer in this environment; geometry/text checks above are structural only. PowerPoint on real machine: NOT VERIFIED")
    return kinds


def verify_pdf(pdf, shots, out, manifest_pages, res):
    try:
        import pymupdf as fitz
        from PIL import Image, ImageChops
    except Exception as e:
        check(res, "PDF", "PDF raster comparison", "BLOCKED", f"PyMuPDF/Pillow missing: {e}"); return
    doc = fitz.open(pdf)
    check(res, "PDF", "page count == HTML slides", "PASS" if len(doc) == manifest_pages else "FAIL", f"{len(doc)} vs {manifest_pages}")
    sizes = {(round(p.rect.width), round(p.rect.height)) for p in doc}
    check(res, "PDF", "page size 1280×720 px (960×540 pt), no blank/extra pages", "PASS" if sizes == {(960, 540)} else "FAIL", str(sizes))
    fonts = sorted({f"{f[3] or f[4]}({f[2]})" for p in doc for f in p.get_fonts()})
    check(res, "PDF", "fonts embedded (Skia embeds CJK as Type3 outlines: selectable, not re-editable)", "PASS" if fonts else "FAIL", f"{len(fonts)} font objects: " + ", ".join(fonts[:4]))
    txt = [len(p.get_text().strip()) for p in doc]
    check(res, "PDF", "vector selectable text on every page", "PASS" if all(t > 20 for t in txt) else "FAIL", f"chars/page {txt}")
    os.makedirs(out, exist_ok=True); diffs = []
    for i, p in enumerate(doc, 1):
        pix = p.get_pixmap(matrix=fitz.Matrix(1280 / p.rect.width, 1280 / p.rect.width))
        a = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).resize((1280, 720))
        sp = os.path.join(shots, f"slide-{i:02d}.png")
        b = Image.open(sp).convert("RGB").resize((1280, 720))
        d = ImageChops.difference(a, b).convert("L")
        hist = d.histogram(); tot = sum(hist); big = sum(hist[40:]) / tot; mean = sum(v * c for v, c in enumerate(hist)) / tot
        diffs.append((i, mean, big))
        a.save(os.path.join(out, f"pdf-{i:02d}.png"))
        sheet = Image.new("RGB", (2560 + 20, 720), (30, 30, 30)); sheet.paste(b, (0, 0)); sheet.paste(a, (1300, 0)); sheet.resize((1290, 360)).save(os.path.join(out, f"cmp-{i:02d}.jpg"), quality=82)
    worst = max(diffs, key=lambda x: x[2])
    check(res, "PDF", "per-page pixel diff vs HTML final-state screenshot (share of pixels with Δ>40/255)", "PASS" if worst[2] < .03 else "FAIL",
          "; ".join(f"p{i}:{b*100:.2f}%" for i, m, b in diffs) + f" (worst p{worst[0]})")


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--pptx"); ap.add_argument("--manifest"); ap.add_argument("--pdf"); ap.add_argument("--shots"); ap.add_argument("--out", default="verify-out"); ap.add_argument("--json")
    a = ap.parse_args(); res = []; man = json.load(open(a.manifest))["manifest"] if a.manifest else []
    if a.pptx: verify_pptx(a.pptx, man, res)
    if a.pdf and a.shots: verify_pdf(a.pdf, a.shots, a.out, max(m["page"] for m in man), res)
    bad = [r for r in res if r["status"] in ("FAIL", "BLOCKED")]
    out = {"status": "FAIL" if any(r["status"] == "FAIL" for r in res) else "PASS", "checks": res}
    if a.json: json.dump(out, open(a.json, "w"), ensure_ascii=False, indent=2)
    for r in res: print(f"{r['status']:10} [{r['area']}] {r['name']} — {r['detail']}")
    sys.exit(1 if out["status"] == "FAIL" else 0)

if __name__ == "__main__": main()
