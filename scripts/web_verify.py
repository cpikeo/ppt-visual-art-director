#!/usr/bin/env python3
"""Web → PPTX / PDF 对照核验：真实读取导出文件，逐页与 HTML 终态截图、导出清单（manifest）比对。
用法: python scripts/web_verify.py --pptx deck.pptx --manifest deck.plan.manifest.json
        [--pdf deck.pdf [--pdf-report r.json]] --shots DIR [--render-pptx | --pptx-render 渲染.pdf] --out DIR [--json report.json] [--force]

状态口径（唯一）：检查项 PASS｜BLOCK（核验失败，必须修）｜UNVERIFIED（没执行、环境缺失或无法判定）｜INFO。
汇总只有一条规则：BLOCK > UNVERIFIED > PASS，且 PASS 要求没有任何 UNVERIFIED 项；没跑的不是通过。
退出码与状态一一对应（EXIT_BY_STATUS 唯一出处，报告里的 exit_code 同值）：PASS 0、BLOCK 1、输入/契约错误 2、UNVERIFIED 4。
PDF 光栅化用 PyMuPDF（AGPL，仅本机可选工具，见 requirements-render.txt）；SSIM 只需 numpy。"""
import argparse, io, json, os, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from model import Report, guard_paths, write_atomic  # noqa: E402  共用路径守卫与原子写入（与 compile 同一实现）

IN = 914400
EXIT_BY_STATUS = {"PASS": 0, "BLOCK": 1, "ERROR": 2, "UNVERIFIED": 4}
# 渲染差异阈值不在这里写常数：见 limits.json 的 web_verify（见 limits.json 注明的来源与暂定状态）。
RESULTS = {"native", "raster", "downgraded", "skipped", "dropped"}   # manifest 每个源元素必须落在其中之一


def check(res, area, name, status, detail):
    res.append({"area": area, "name": name, "status": status, "detail": detail})


def rollup(res):
    """检查项 → 总状态。UNVERIFIED 不得被汇总成 PASS。"""
    st = {r["status"] for r in res}
    for s in ("BLOCK", "UNVERIFIED"):
        if s in st:
            return s
    return "PASS" if "PASS" in st else "UNVERIFIED"   # 一项都没跑，不是通过


def load_manifest(path):
    data = json.load(open(path, encoding="utf-8"))
    man = data["manifest"] if isinstance(data, dict) else data
    pages = data.get("pages", []) if isinstance(data, dict) else []
    for m in man:   # 旧版 manifest 没有 result：按 strategy 推断，推不出的留空，由检查项报出
        if "result" not in m:
            s = m.get("strategy", "")
            m["result"] = "native" if s.startswith("native") else "raster" if s.startswith("raster") else ""
    return man, (data.get("flags", []) if isinstance(data, dict) else []), pages


def verify_manifest(manifest, flags, res):
    """每个源元素必须有明确转换结果；降级与错误标记必须进入验收结论，而不是留在日志里。"""
    bad = [f"p{m.get('page')}/{m.get('id')}" for m in manifest if m.get("result") not in RESULTS]
    dropped = [f"p{m['page']}/{m['id']} ({m.get('reason') or m.get('src')})" for m in manifest if m.get("result") == "dropped"]
    check(res, "清单", "every source element has an explicit conversion result", "PASS" if not bad else "BLOCK",
          f"{len(manifest)} elements" + (f"; no result: {bad[:5]}" if bad else ""))
    check(res, "清单", "no element silently dropped (unsupported format / invalid chart / broken source)", "PASS" if not dropped else "BLOCK",
          "; ".join(dropped[:5]) or "none")
    errs = [f for f in flags if str(f.get("code", "")).startswith("E-")]
    check(res, "清单", "no unresolved conversion errors (E-*)", "PASS" if not errs else "BLOCK",
          "; ".join(f"p{f.get('page')} {f.get('code')} {f.get('msg')}" for f in errs[:5]) or "none")
    rasters = [m for m in manifest if m.get("result") == "raster"]
    w_raster = [f for f in flags if f.get("code") == "W-RASTER"]
    check(res, "清单", "every raster is declared by a W-RASTER flag", "PASS" if len(w_raster) >= len(rasters) else "BLOCK",
          f"{len(rasters)} rasters / {len(w_raster)} flags")
    down = [{"page": m["page"], "id": m["id"], "result": m["result"], "why": m.get("reason") or m.get("src")}
            for m in manifest if m.get("result") in ("raster", "downgraded", "skipped")]
    codes = {}
    for f in flags:
        codes[f.get("code", "?")] = codes.get(f.get("code", "?"), 0) + 1
    check(res, "清单", "downgrade list (carried into this report; deliver it with the deck)", "INFO",
          f"{len(down)} elements downgraded/skipped; flags {json.dumps(codes, ensure_ascii=False)}")
    return down


def _norm(t):
    return "".join(str(t or "").split())


def _rgb(fmt_fill):
    try:
        return "#" + str(fmt_fill.fore_color.rgb).upper()
    except Exception:
        return None


def _chart_check(sh, m, bad):
    """逐图核验：类型、类别、系列名与数值、颜色（系列或扇区）、图例、数据标签，全部对照 manifest 里的契约解析结果。"""
    from model import chart_type_name
    c, ch = m["chart"], sh.chart
    loc = f"p{m['page']}/{m['id']}"
    spec = {"kind": c["kind"], "stacked": c.get("stacked"), "horizontal": c.get("horizontal")}
    want_type = chart_type_name(spec)
    if ch.chart_type.name != want_type:
        bad.append(f"{loc} type {ch.chart_type.name}≠{want_type}")
    plot = ch.plots[0]
    cats = [str(x) for x in plot.categories]
    if c.get("categories") is not None and cats != [str(x) for x in c["categories"]]:
        bad.append(f"{loc} categories {cats}≠{c['categories']}")
    got = [(str(se.name), [float(v) for v in se.values]) for se in plot.series]
    want = [(str(se["name"]), [float(v) for v in se["values"]]) for se in c.get("series") or []]
    if got != want:
        bad.append(f"{loc} series data mismatch")
    if bool(ch.has_legend) != bool(c.get("legend")):
        bad.append(f"{loc} legend {ch.has_legend}≠{c.get('legend')}")
    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    shows = ch._chartSpace.xpath('.//c:dLbls/c:showVal[@val="1"] | .//c:dLbls/c:showCatName[@val="1"] | .//c:dLbls/c:showPercent[@val="1"]')
    if bool(shows) != bool(c.get("data_labels")):
        bad.append(f"{loc} data_labels {bool(shows)}≠{c.get('data_labels')}")
    cm = c.get("color_map") or {}
    pie_like = c["kind"] in ("pie", "donut")
    if pie_like:
        for k, cat in enumerate(c.get("categories") or []):
            col = _rgb(plot.series[0].points[k].format.fill)
            if col and cm.get(str(cat)) and col != cm[str(cat)].upper():
                bad.append(f"{loc} slice '{cat}' {col}≠{cm[str(cat)]}")
    else:
        for se in plot.series:
            if c.get("highlight") is not None:
                break
            col = _rgb(se.format.fill) if c["kind"] in ("bar", "area") else (_rgb(se.format.line) if c["kind"] == "line" else _rgb(se.format.fill))
            if col and cm.get(str(se.name)) and col != cm[str(se.name)].upper():
                bad.append(f"{loc} series '{se.name}' {col}≠{cm[str(se.name)]}")
        if c.get("highlight") is not None:
            pts = [_rgb(plot.series[0].points[k].format.fill) for k in range(len(cats))]
            hi, base = c["highlight"], cm.get(str(plot.series[0].name), "").upper()
            rest = {p for k, p in enumerate(pts) if k != hi}
            if pts[hi] != base or len(rest) != 1 or pts[hi] in rest:
                bad.append(f"{loc} highlight colours {pts}")


def verify_pptx(path, manifest, pages, res):
    """DOM id ↔ manifest ↔ PPTX 对象：存在、类型、几何、文字、叠放顺序、图片裁切、图表、备注、页数。"""
    from pptx import Presentation
    prs = Presentation(path)
    n_pages = max((m["page"] for m in manifest), default=0)
    check(res, "PPTX", "slide count == HTML slides", "PASS" if len(prs.slides) == n_pages else "BLOCK", f"{len(prs.slides)} vs {n_pages}")
    check(res, "PPTX", "page size 16:9 (12192000×6858000 EMU)", "PASS" if (prs.slide_width, prs.slide_height) == (12192000, 6858000) else "BLOCK", f"{prs.slide_width}×{prs.slide_height}")
    kinds = {"text": 0, "picture": 0, "chart": 0, "table": 0, "shape": 0, "other": 0}
    full_page, geom_bad, text_bad, missing, type_bad, crop_bad, order_bad, chart_bad, table_bad, notes_bad = ([] for _ in range(10))
    matched = 0
    area = prs.slide_width * prs.slide_height
    declared = {p["page"]: p for p in pages}
    for si, slide in enumerate(prs.slides, 1):
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
        info = declared.get(si)
        if info is not None:   # 备注按 HTML 实际声明核验：声明了必须在；没声明则 PPTX 不得凭空出现
            if info["notes_declared"] and not _norm(notes):
                notes_bad.append(f"p{si} declared but missing")
            elif not info["notes_declared"] and _norm(notes):
                notes_bad.append(f"p{si} not declared in HTML but present")
        by_name = {sh.name: sh for sh in slide.shapes}
        order = {sh.name: k for k, sh in enumerate(slide.shapes)}
        for sh in slide.shapes:
            if sh.shape_type == 13: kinds["picture"] += 1
            elif getattr(sh, "has_chart", False) and sh.has_chart: kinds["chart"] += 1
            elif getattr(sh, "has_table", False) and sh.has_table: kinds["table"] += 1
            elif sh.has_text_frame and sh.text_frame.text.strip(): kinds["text"] += 1
            elif sh.shape_type in (1, 17): kinds["shape"] += 1
            else: kinds["other"] += 1
            # 整页大图只有一种合法来源：DOM 里真实存在的 <img>/CSS 背景图（manifest 记为 native）。栅格结果或来历不明的整页图一律拒绝。
            if sh.shape_type == 13 and sh.width * sh.height >= .9 * area and sh.name not in {f"{m['page_id']}.{m['id']}" for m in manifest if m.get("type") == "image" and m.get("result") == "native"}:
                full_page.append(f"p{si}:{sh.name}")
        last = -1
        for m in manifest:
            if m["page"] != si or m.get("result") in ("skipped", "dropped"):
                continue
            sh = by_name.get(f"{m['page_id']}.{m['id']}") if m.get("page_id") else next((c for n, c in by_name.items() if n.endswith("." + m["id"])), None)
            loc = f"p{si}/{m['id']}"
            if sh is None:
                missing.append(loc); continue
            matched += 1
            if order[sh.name] < last: order_bad.append(loc)
            last = max(last, order[sh.name])
            dx = max(abs(sh.left - m["x"] * IN), abs(sh.top - m["y"] * IN), abs(sh.width - m["w"] * IN), abs(sh.height - m["h"] * IN)) / IN
            if dx > 0.02: geom_bad.append(f"{loc} Δ{dx:.3f}in")
            st = m.get("strategy", "")
            if st == "native-text":
                if not (sh.has_text_frame and sh.text_frame.text.strip()): type_bad.append(loc)
                elif m.get("text") is not None and _norm(sh.text_frame.text) != _norm(m["text"]): text_bad.append(f"{loc} '{sh.text_frame.text[:20]}'≠'{str(m['text'])[:20]}'")
            if m.get("result") == "raster" and sh.shape_type != 13: type_bad.append(f"{loc} raster expected picture")
            if st == "native-image" and m.get("result") == "native":
                if sh.shape_type != 13: type_bad.append(f"{loc} image expected picture")
                else:
                    cl, cr, ct, cb = sh.crop_left, sh.crop_right, sh.crop_top, sh.crop_bottom
                    iw, ih = sh.image.size
                    vis = ((1 - cl - cr) * iw) / max(1e-9, (1 - ct - cb) * ih)
                    if min(cl, cr, ct, cb) < -1e-6 or abs(vis / (sh.width / sh.height) - 1) > 0.02: crop_bad.append(f"{loc} visible AR {vis:.3f} vs box {sh.width / sh.height:.3f}")
            if st == "native-chart":
                c = m.get("chart") or {}
                if c.get("kind") == "table":
                    if not getattr(sh, "has_table", False): type_bad.append(f"{loc} expected native table")
                    else:
                        tb = sh.table
                        if len(tb.rows) != (c.get("rows") or 0) + 1 or len(tb.columns) != len(c.get("headers") or []) or [tb.cell(0, k).text for k in range(len(tb.columns))] != list(c.get("headers") or []): table_bad.append(loc)
                elif not getattr(sh, "has_chart", False): type_bad.append(f"{loc} expected native chart")
                elif c:
                    try: _chart_check(sh, m, chart_bad)
                    except Exception as e: chart_bad.append(f"{loc} unreadable: {type(e).__name__}: {e}")
    expect = sum(1 for m in manifest if m.get("result") not in ("skipped", "dropped"))
    def line(ok_name, bad, detail_ok):
        check(res, "PPTX", ok_name, "PASS" if not bad else "BLOCK", (", ".join(bad[:5]) + (f" (+{len(bad) - 5})" if len(bad) > 5 else "")) if bad else detail_ok)
    line("every exported element exists as a named object (DOM id ↔ manifest ↔ PPTX)", missing, f"{matched}/{expect}")
    line("object types match the manifest strategy (text/picture/chart/table)", type_bad, "ok")
    line("geometry within 0.02 in of DOM box", geom_bad, f"{matched} objects")
    line("text content equals DOM text (whitespace-insensitive)", text_bad, "ok")
    line("z-order follows DOM order", order_bad, "ok")
    line("native image crop: visible aspect == box aspect (±2%)", crop_bad, "ok")
    line("charts: type, categories, series data, colours, legend, data labels == contract", chart_bad, f"{kinds['chart']} charts")
    line("tables: header and row count == DOM", table_bad, f"{kinds['table']} tables")
    check(res, "PPTX", "no full-page screenshot posing as slide", "PASS" if not full_page else "BLOCK", ", ".join(full_page) or "no raster/unknown picture ≥ 90% of slide area (declared native full-bleed images are allowed)")
    nd = sum(1 for p in pages if p["notes_declared"])
    line("speaker notes follow what the HTML declares (declared → present; undeclared → absent)", notes_bad, f"{nd}/{len(pages)} slides declare notes" if pages else "no page info in manifest (legacy)")
    n = sum(kinds.values())
    check(res, "PPTX", "object inventory", "INFO", json.dumps(kinds) + f"; pictures = {kinds['picture']} (declared raster/image only); non-picture objects = {n - kinds['picture']}/{n}")


def stats_of(manifest):
    """原生 / 局部栅格 / 降级 / 跳过 / 丢弃 统计（披露不可编辑项）。"""
    by = {}
    for m in manifest:
        by[m["result"]] = by.get(m["result"], 0) + 1
    tot = sum(v for k, v in by.items() if k != "skipped")
    return {"elements": len(manifest), "by_result": by, "editable_share": round((by.get("native", 0) + by.get("downgraded", 0)) / tot, 4) if tot else None,
            "not_editable": [{"page": m["page"], "id": m["id"], "sel": m.get("sel"), "why": m.get("reason"), "raster": m.get("raster")} for m in manifest if m["result"] == "raster"],
            "downgraded": [{"page": m["page"], "id": m["id"], "why": m.get("reason")} for m in manifest if m["result"] == "downgraded"]}


def _arr(img):
    import numpy as np
    return np.asarray(img.convert("RGB"), dtype=np.int16)


def _box(x, k=7):
    import numpy as np
    c = np.cumsum(np.cumsum(np.pad(x, ((1, 0), (1, 0))), 0), 1)
    return (c[k:, k:] - c[:-k, k:] - c[k:, :-k] + c[:-k, :-k]) / (k * k)


def ssim(a, b):
    """灰度 SSIM（7×7 均匀窗，C1/C2 取标准值）；无 skimage 依赖，只需 numpy。"""
    import numpy as np
    g = lambda im: np.asarray(im.convert("L"), dtype=np.float64)
    x, y = g(a), g(b)
    c1, c2 = (0.01 * 255) ** 2, (0.03 * 255) ** 2
    mx, my = _box(x), _box(y)
    vx, vy, cxy = _box(x * x) - mx * mx, _box(y * y) - my * my, _box(x * y) - mx * my
    return float(np.mean(((2 * mx * my + c1) * (2 * cxy + c2)) / ((mx * mx + my * my + c1) * (vx + vy + c2))))


def pixel_stats(a, b):
    import numpy as np
    d = np.abs(_arr(a) - _arr(b))
    return {"differing_pixels": int((d.max(axis=2) > 0).sum()), "max_abs": int(d.max()), "mean_abs": round(float(d.mean()), 4),
            "share_gt40": round(float((d.max(axis=2) > 40).mean()), 6), "ssim": round(ssim(a, b), 5)}


def _doc_pages(pdf, scale_to_px):
    import pymupdf as fitz
    from PIL import Image
    doc = fitz.open(pdf)
    for p in doc:
        k = scale_to_px / p.rect.width
        pix = p.get_pixmap(matrix=fitz.Matrix(k, k), alpha=False)
        yield p, Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def _baseline(shots, i):
    from PIL import Image
    f = os.path.join(shots, f"slide-{i:02d}.png")
    return Image.open(f).convert("RGB") if os.path.isfile(f) else None


def verify_pdf_pixel(pdf, rep, shots, out, force, pages, res):
    """像素模式：逐页栅格化 PDF，与同环境 HTML 基准图逐像素比对，目标 0 差异。再抽取页内嵌入图直接比对，区分“封装无损”与“查看器重采样”。"""
    import pymupdf as fitz
    from PIL import Image
    dpr = rep["dpr"]
    px = 1280 * dpr
    doc = fitz.open(pdf)
    check(res, "PDF", "page count == HTML slides", "PASS" if len(doc) == pages else "BLOCK", f"{len(doc)} vs {pages}")
    sizes = {(round(p.rect.width), round(p.rect.height)) for p in doc}
    check(res, "PDF", "page size 16:9 (960×540 pt = 1280×720 px), no blank/extra pages", "PASS" if sizes == {(960, 540)} else "BLOCK", str(sizes))
    missing = [i for i in range(1, len(doc) + 1) if _baseline(shots, i) is None]
    if missing:
        check(res, "PDF", "pixel-exact vs HTML baseline", "UNVERIFIED", f"missing baseline pages {missing[:8]} in {shots} (run web_export.mjs shots)"); return
    base0 = _baseline(shots, 1)
    if base0.size != (px, 720 * dpr):
        check(res, "PDF", "pixel-exact vs HTML baseline", "UNVERIFIED", f"baseline is {base0.size}, PDF report says {px}×{720 * dpr} (re-run shots with --dpr {dpr})"); return
    raster, embedded = [], []
    for i, (p, im) in enumerate(_doc_pages(pdf, px), 1):
        b = _baseline(shots, i)
        raster.append(pixel_stats(im, b) if im.size == b.size else {"differing_pixels": -1, "max_abs": -1, "note": f"size {im.size}≠{b.size}"})
        imgs = p.get_images(full=True)
        if len(imgs) == 1:
            ex = Image.open(io.BytesIO(doc.extract_image(imgs[0][0])["image"])).convert("RGB")
            embedded.append(pixel_stats(ex, b) if ex.size == b.size else {"differing_pixels": -1, "max_abs": -1, "note": f"size {ex.size}≠{b.size}"})
        else:
            embedded.append({"differing_pixels": -1, "max_abs": -1, "note": f"{len(imgs)} images on page"})
        if raster[-1]["differing_pixels"] != 0:
            buf = io.BytesIO(); im.save(buf, "PNG"); write_atomic(Path(out) / f"pdf-{i:02d}.png", buf.getvalue(), force)
    worst = max(raster, key=lambda r: r["differing_pixels"])
    ok = all(r["differing_pixels"] == 0 for r in raster)
    check(res, "PDF", f"pixel mode: rasterised PDF page == HTML baseline at {dpr}× (differing pixels must be 0)", "PASS" if ok else "BLOCK",
          "; ".join(f"p{i}:{r['differing_pixels']}px(max Δ{r['max_abs']})" for i, r in enumerate(raster, 1)))
    eok = all(r["differing_pixels"] == 0 for r in embedded)
    check(res, "PDF", "pixel mode: embedded page image == HTML baseline bytes-for-pixels (lossless packaging)", "PASS" if eok else "BLOCK",
          "; ".join(f"p{i}:{r['differing_pixels']}px" for i, r in enumerate(embedded, 1)))
    text = sum(len("".join(p.get_text().split())) for p in doc)
    check(res, "PDF", "text layer (pixel mode is image pages by design: NOT selectable, NOT vector-editable)", "INFO", f"{text} selectable chars (expected 0)")
    rep_out = {"mode": "pixel", "dpr": dpr, "raster_vs_baseline": raster, "embedded_vs_baseline": embedded, "worst_raster_differing_px": worst["differing_pixels"]}
    return rep_out


def verify_image_pptx(pptx, rep, shots, res):
    """展示版 PPTX（不可编辑）：解包后逐页核对 ppt/media 原始字节 SHA-256、像素尺寸、16:9、页数、顺序，并与独立渲染的基准逐像素比对。"""
    import hashlib, zipfile
    from PIL import Image
    from pptx import Presentation
    prs = Presentation(pptx)
    want = rep["pages"]; dpr = rep["dpr"]
    check(res, "展示版 PPTX", "slide count == HTML slides", "PASS" if len(prs.slides) == len(want) else "BLOCK", f"{len(prs.slides)} vs {len(want)}")
    check(res, "展示版 PPTX", "page size 16:9 (12192000×6858000 EMU)", "PASS" if (prs.slide_width, prs.slide_height) == (12192000, 6858000) else "BLOCK", f"{prs.slide_width}×{prs.slide_height}")
    zf = zipfile.ZipFile(pptx)
    bad_obj, bad_sha, bad_dim, bad_cover, bad_base, rows = [], [], [], [], [], []
    for i, (slide, w) in enumerate(zip(prs.slides, want), 1):
        shapes = list(slide.shapes)
        if len(shapes) != 1 or shapes[0].shape_type != 13 or any(sh.has_text_frame and sh.text_frame.text.strip() for sh in shapes):
            bad_obj.append(f"p{i}:{len(shapes)} shapes"); continue
        pic = shapes[0]
        if (pic.left, pic.top, pic.width, pic.height) != (0, 0, 12192000, 6858000): bad_cover.append(f"p{i}")
        raw = zf.read(str(slide.part.related_part(pic._element.blip_rId).partname).lstrip("/"))
        sha = hashlib.sha256(raw).hexdigest(); rows.append(sha[:12])
        if sha != w["png_sha256"]: bad_sha.append(f"p{i}")
        im = Image.open(io.BytesIO(raw))
        if im.size != (1280 * dpr, 720 * dpr) or im.size[0] * 9 != im.size[1] * 16: bad_dim.append(f"p{i}:{im.size}")
        if shots:
            b = _baseline(shots, i)
            if b is None or b.size != im.size: bad_base.append(f"p{i}:no comparable baseline")
            else:
                st = pixel_stats(im.convert("RGB"), b)
                if st["differing_pixels"] != 0: bad_base.append(f"p{i}:{st['differing_pixels']}px")
    check(res, "展示版 PPTX", "each slide = exactly one picture, no text/editable objects (image-only by design)", "PASS" if not bad_obj else "BLOCK", ", ".join(bad_obj) or "image-only")
    check(res, "展示版 PPTX", "picture fills the whole slide (no crop/stretch offset)", "PASS" if not bad_cover else "BLOCK", ", ".join(bad_cover) or "0,0,full")
    check(res, "展示版 PPTX", "unpacked media SHA-256 == PNG batch used for PDF, in slide order", "PASS" if not bad_sha and len(rows) == len(want) else "BLOCK", ", ".join(bad_sha) or " ".join(rows))
    check(res, "展示版 PPTX", f"original pixel size {1280 * dpr}×{720 * dpr} (16:9, not resampled)", "PASS" if not bad_dim else "BLOCK", ", ".join(bad_dim) or "ok")
    if shots:
        check(res, "展示版 PPTX", "image pixels == independently rendered HTML baseline (differing pixels must be 0)", "PASS" if not bad_base else "BLOCK", ", ".join(bad_base) or "0 differing pixels on every page")
    else:
        check(res, "展示版 PPTX", "image pixels vs independent HTML baseline", "UNVERIFIED", "no --shots baseline given")


def render_pptx(pptx, outdir):
    """LibreOffice 真实渲染 PPTX → PDF；独立用户配置目录，超时即 UNVERIFIED。返回 (pdf 路径|None, 说明)。"""
    import shutil, subprocess, tempfile
    so = shutil.which("soffice") or shutil.which("libreoffice") or os.environ.get("SOFFICE")
    if not so:
        return None, "soffice not found (install LibreOffice or set SOFFICE)"
    with tempfile.TemporaryDirectory() as prof:
        try:
            r = subprocess.run([so, f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(pptx)], capture_output=True, text=True, timeout=240)
        except subprocess.TimeoutExpired:
            return None, "soffice timed out (240 s)"
    pdf = Path(outdir) / (Path(pptx).stem + ".pdf")
    ver = subprocess.run([so, "--version"], capture_output=True, text=True).stdout.strip()
    return (str(pdf), ver) if pdf.is_file() else (None, f"soffice produced no PDF: {r.stderr.strip()[:200]}")


def verify_pptx_render(render_pdf, shots, pages, res, ver):
    from model import LIMITS
    from PIL import Image
    lim = LIMITS["web_verify"]["pptx_render"]
    stats = []
    for i, (p, im) in enumerate(_doc_pages(render_pdf, 1280), 1):
        b = _baseline(shots, i)
        if b is None:
            check(res, "PPTX", "visual fidelity vs HTML", "UNVERIFIED", f"missing baseline page {i}"); return None
        stats.append(pixel_stats(im.resize((1280, 720)), b.resize((1280, 720), Image.LANCZOS)))
    check(res, "PPTX", "rendered page count == HTML slides", "PASS" if len(stats) == pages else "BLOCK", f"{len(stats)} vs {pages}")
    worst, wssim = max(r["share_gt40"] for r in stats), min(r["ssim"] for r in stats)
    ok = worst <= lim["share_gt40_max"] and wssim >= lim["ssim_min"]
    check(res, "PPTX", f"LibreOffice render ({ver}) vs HTML: share Δ>40 ≤ {lim['share_gt40_max']:.0%}, SSIM ≥ {lim['ssim_min']} (calibrated: {lim['calibration']})", "PASS" if ok else "BLOCK",
          "; ".join(f"p{i}:{r['share_gt40'] * 100:.1f}%/ssim {r['ssim']:.3f}" for i, r in enumerate(stats, 1)))
    check(res, "PPTX", "Microsoft PowerPoint real render", "UNVERIFIED", "not run in this environment (no PowerPoint); LibreOffice is a different layout engine")
    return {"renderer": ver, "per_page": stats}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx"); ap.add_argument("--manifest"); ap.add_argument("--pdf"); ap.add_argument("--pdf-report", help="web_export.mjs pdf 写出的 *.pdf-report.json（缺省取 PDF 同名）")
    ap.add_argument("--image-pptx", help="展示版（整页图像）PPTX"); ap.add_argument("--image-pptx-report", help="pptx-image 写出的 *.image-pptx-report.json")
    ap.add_argument("--shots", help="web_export.mjs shots 的 HTML 基准图目录")
    ap.add_argument("--render-pptx", action="store_true", help="用本机 LibreOffice 真实渲染 PPTX 后比对")
    ap.add_argument("--pptx-render", help="已有的 PPTX 渲染 PDF（LibreOffice / PowerPoint 导出）")
    ap.add_argument("--out", default="verify-out"); ap.add_argument("--json"); ap.add_argument("--force", action="store_true", help="允许覆盖已存在的报告/对比图")
    a = ap.parse_args(argv)
    err = []
    if not (a.pptx or a.pdf or a.image_pptx): err.append("至少给 --pptx、--pdf 或 --image-pptx")
    if a.image_pptx and not a.image_pptx_report: err.append("--image-pptx 需要 --image-pptx-report（逐页 SHA-256 来源）")
    if a.pptx and not a.manifest: err.append("--pptx 需要 --manifest（没有清单无法逐元素核对）")
    if (a.pptx_render or a.render_pptx) and not a.pptx: err.append("--pptx-render / --render-pptx 需要 --pptx")
    for k in ("pptx", "manifest", "pdf", "pptx_render", "pdf_report", "image_pptx", "image_pptx_report"):
        v = getattr(a, k)
        if v and not os.path.isfile(v): err.append(f"--{k.replace('_', '-')} 文件不存在: {v}")
    if a.shots and not os.path.isdir(a.shots): err.append(f"--shots 不是目录: {a.shots}")
    rep = Report()
    outdir = Path(a.out).resolve()
    ins = {k: Path(getattr(a, k)) for k in ("pptx", "manifest", "pdf", "pptx_render", "pdf_report", "image_pptx", "image_pptx_report") if getattr(a, k)}
    outs = [("--json", Path(a.json), ".json")] if a.json else []
    guard_paths(rep, ins, outs, a.force, dirs={"--out": outdir, **({"--shots": Path(a.shots)} if a.shots else {})})
    if a.shots and outdir == Path(a.shots).resolve(): rep.hard("E-EXIST", str(outdir), "--out 不能等于 --shots：对比图不得混进 HTML 基准图")
    err += [f"{i.code} {i.where}: {i.message}" for i in rep.items if i.level == "HARD"]
    if err:
        for e in err: print("ERROR", e, file=sys.stderr)
        return EXIT_BY_STATUS["ERROR"]
    manifest, flags, pages = load_manifest(a.manifest) if a.manifest else ([], [], [])
    n_pages = max((m["page"] for m in manifest), default=None)
    res, down, detail, stats = [], [], {}, None
    if a.manifest:
        down = verify_manifest(manifest, flags, res); stats = stats_of(manifest)
    try:
        if a.pptx: verify_pptx(a.pptx, manifest, pages, res)
    except Exception as e:   # 打不开/损坏的 PPTX 是 BLOCK，不是崩溃后的“无结论”
        check(res, "PPTX", "PPTX opens as a valid package", "BLOCK", f"{type(e).__name__}: {e}")
    outdir.mkdir(parents=True, exist_ok=True)
    try:
        if a.pdf:
            rp = a.pdf_report or str(Path(a.pdf).with_suffix("")) + ".pdf-report.json"
            prep = json.load(open(rp, encoding="utf-8")) if os.path.isfile(rp) else None
            if not a.shots: check(res, "PDF", "PDF raster comparison", "UNVERIFIED", "需要 --shots（web_export.mjs shots 的基准图）才能逐页比对")
            elif prep is None: check(res, "PDF", "PDF report missing", "UNVERIFIED", f"找不到 {rp}：无法确认 PDF 的来源与模式")
            else:
                check(res, "PDF", "pdf-report sha256 == file sha256", "PASS" if prep.get("pdf_sha256") == __import__("hashlib").sha256(open(a.pdf, "rb").read()).hexdigest() else "BLOCK", f"mode={prep.get('mode')}")
                if prep.get("mode") != "pixel": check(res, "PDF", "PDF is pixel mode", "BLOCK", f"mode={prep.get('mode')}: only pixel-mode PDFs are supported")
                else: detail["pdf"] = verify_pdf_pixel(a.pdf, prep, a.shots, outdir, a.force, n_pages, res)
        if a.image_pptx:
            verify_image_pptx(a.image_pptx, json.load(open(a.image_pptx_report, encoding="utf-8")), a.shots, res)
        if a.pptx:
            rpdf, ver = a.pptx_render, "external render"
            if a.render_pptx and not rpdf:
                rpdf, ver = render_pptx(a.pptx, outdir)
                if rpdf is None: check(res, "PPTX", "LibreOffice render", "UNVERIFIED", ver)
            if rpdf and a.shots: detail["pptx_render"] = verify_pptx_render(rpdf, a.shots, n_pages, res, ver)
            elif rpdf: check(res, "PPTX", "visual fidelity vs HTML", "UNVERIFIED", "需要 --shots 基准图")
            elif not a.render_pptx:
                check(res, "PPTX", "visual fidelity vs HTML (needs LibreOffice/PowerPoint render: --render-pptx or --pptx-render)", "UNVERIFIED", "未提供 PPTX 的真实渲染；上面是结构层核验。PowerPoint 实机: UNVERIFIED")
    except ImportError as e:
        check(res, "PDF/PPTX", "raster comparison", "UNVERIFIED", f"PyMuPDF/Pillow/numpy missing: {e} (pip install -r requirements-render.txt)")
    status = rollup(res)
    counts = {s: sum(1 for r in res if r["status"] == s) for s in ("PASS", "BLOCK", "UNVERIFIED", "INFO")}
    out = {"status": status, "exit_code": EXIT_BY_STATUS[status], "counts": counts, "stats": stats, "downgrades": down, "detail": detail, "checks": res}
    if a.json: write_atomic(Path(a.json), json.dumps(out, ensure_ascii=False, indent=2).encode("utf-8"), a.force)
    for r in res: print(f"{r['status']:10} [{r['area']}] {r['name']} — {r['detail']}")
    print(f"STATUS {status}  exit={EXIT_BY_STATUS[status]}  " + " ".join(f"{k}={v}" for k, v in counts.items()) + f"  downgrades={len(down)}")
    return EXIT_BY_STATUS[status]


if __name__ == "__main__":
    sys.exit(main())
