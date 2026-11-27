#!/usr/bin/env python3
"""Web → PPTX / PDF 对照核验：真实读取导出文件，逐页与 HTML 终态截图、导出清单（manifest）比对。
用法: python scripts/web_verify.py --pptx deck.pptx --manifest deck.plan.manifest.json
        [--pdf deck.pdf --shots DIR --out DIR] [--pptx-render PPTX渲染得到的.pdf] [--json report.json] [--force]

状态口径（唯一）：检查项 PASS｜FAIL｜BLOCKED（环境/输入缺失，无法执行）｜UNVERIFIED（没执行或无法判定）｜INFO。
汇总只有一条规则：FAIL > BLOCKED > UNVERIFIED > PASS，且 PASS 要求没有任何 BLOCKED / UNVERIFIED 项。
退出码与状态一一对应（EXIT_BY_STATUS 唯一出处，报告里的 exit_code 同值）：PASS 0、FAIL 1、输入/契约错误 2、BLOCKED 3、UNVERIFIED 4。
PDF 光栅化用 PyMuPDF（AGPL，仅本机可选工具，见 requirements-render.txt）。"""
import argparse, io, json, os, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from model import Report, guard_paths, write_atomic  # noqa: E402  共用路径守卫与原子写入（与 compile/revise 同一实现）

IN = 914400
EXIT_BY_STATUS = {"PASS": 0, "FAIL": 1, "ERROR": 2, "BLOCKED": 3, "UNVERIFIED": 4}
RESULTS = {"native", "raster", "downgraded", "skipped", "dropped"}   # manifest 每个源元素必须落在其中之一
PPTX_DIFF_LIMIT = 0.12   # UNCALIBRATED：PPTX 渲染 PDF 与 HTML 终态的像素差上限，尚无实机样本校准，见 dev/evidence.json
PDF_DIFF_LIMIT = 0.03    # Chromium 同引擎导出，差异应近零；0.03 为经验值（见 dev/evidence.json）


def check(res, area, name, status, detail):
    res.append({"area": area, "name": name, "status": status, "detail": detail})


def rollup(res):
    """检查项 → 总状态。BLOCKED / UNVERIFIED 不得被汇总成 PASS。"""
    st = {r["status"] for r in res}
    for s in ("FAIL", "BLOCKED", "UNVERIFIED"):
        if s in st:
            return s
    return "PASS" if "PASS" in st else "UNVERIFIED"   # 一项都没跑，不是通过


def load_manifest(path):
    data = json.load(open(path, encoding="utf-8"))
    man = data["manifest"] if isinstance(data, dict) else data
    for m in man:   # 旧版 manifest 没有 result：按 strategy 推断，推不出的留空，由检查项报出
        if "result" not in m:
            s = m.get("strategy", "")
            m["result"] = "native" if s.startswith("native") else "raster" if s.startswith("raster") else ""
    return man, (data.get("flags", []) if isinstance(data, dict) else [])


def verify_manifest(manifest, flags, res):
    """每个源元素必须有明确转换结果；降级与错误标记必须进入验收结论，而不是留在日志里。"""
    bad = [f"p{m.get('page')}/{m.get('id')}" for m in manifest if m.get("result") not in RESULTS]
    dropped = [f"p{m['page']}/{m['id']} ({m.get('reason') or m.get('src')})" for m in manifest if m.get("result") == "dropped"]
    check(res, "清单", "every source element has an explicit conversion result", "PASS" if not bad else "FAIL",
          f"{len(manifest)} elements" + (f"; no result: {bad[:5]}" if bad else ""))
    check(res, "清单", "no element silently dropped (unsupported format / invalid chart / broken source)", "PASS" if not dropped else "FAIL",
          "; ".join(dropped[:5]) or "none")
    errs = [f for f in flags if str(f.get("code", "")).startswith("E-")]
    check(res, "清单", "no unresolved conversion errors (E-*)", "PASS" if not errs else "FAIL",
          "; ".join(f"p{f.get('page')} {f.get('code')} {f.get('msg')}" for f in errs[:5]) or "none")
    rasters = [m for m in manifest if m.get("result") == "raster"]
    w_raster = [f for f in flags if f.get("code") == "W-RASTER"]
    check(res, "清单", "every raster is declared by a W-RASTER flag", "PASS" if len(w_raster) >= len(rasters) else "FAIL",
          f"{len(rasters)} rasters / {len(w_raster)} flags")
    down = [{"page": m["page"], "id": m["id"], "result": m["result"], "why": m.get("reason") or m.get("src")}
            for m in manifest if m.get("result") in ("raster", "downgraded", "skipped")]
    codes = {}
    for f in flags:
        codes[f.get("code", "?")] = codes.get(f.get("code", "?"), 0) + 1
    check(res, "清单", "downgrade list (carried into this report; deliver it with the deck)", "INFO",
          f"{len(down)} elements downgraded/skipped; flags {json.dumps(codes, ensure_ascii=False)}")
    return down


def verify_pptx(path, manifest, res):
    from pptx import Presentation
    prs = Presentation(path)
    pages = max((m["page"] for m in manifest), default=0)
    check(res, "PPTX", "slide count == HTML slides", "PASS" if len(prs.slides) == pages else "FAIL", f"{len(prs.slides)} vs {pages}")
    check(res, "PPTX", "page size 16:9 (12192000×6858000 EMU)", "PASS" if (prs.slide_width, prs.slide_height) == (12192000, 6858000) else "FAIL", f"{prs.slide_width}×{prs.slide_height}")
    kinds = {"text": 0, "picture": 0, "chart": 0, "table": 0, "shape": 0, "other": 0}
    full_page, geom_bad, text_bad, missing = [], [], [], []
    matched = notes = 0
    area = prs.slide_width * prs.slide_height
    for si, slide in enumerate(prs.slides, 1):
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            notes += 1
        by_name = {sh.name: sh for sh in slide.shapes}
        for sh in slide.shapes:
            if sh.shape_type == 13: kinds["picture"] += 1
            elif getattr(sh, "has_chart", False) and sh.has_chart: kinds["chart"] += 1
            elif getattr(sh, "has_table", False) and sh.has_table: kinds["table"] += 1
            elif sh.has_text_frame and sh.text_frame.text.strip(): kinds["text"] += 1
            elif sh.shape_type in (1, 17): kinds["shape"] += 1
            else: kinds["other"] += 1
            if sh.shape_type == 13 and sh.width * sh.height >= .9 * area:
                full_page.append(f"p{si}:{sh.name}")
        for m in manifest:
            if m["page"] != si or m.get("result") in ("skipped", "dropped"):
                continue   # 跳过/丢弃的元素不应出现在 PPTX；其结论由清单检查项负责
            # compile 命名对象 "<页 id>.<元素 id>"；旧 manifest 没有 page_id 时退回后缀匹配
            sh = by_name.get(f"{m['page_id']}.{m['id']}") if m.get("page_id") else next((c for n, c in by_name.items() if n.endswith("." + m["id"])), None)
            if sh is None:
                missing.append(f"p{si}/{m['id']}"); continue
            matched += 1
            dx = max(abs(sh.left - m["x"] * IN), abs(sh.top - m["y"] * IN), abs(sh.width - m["w"] * IN), abs(sh.height - m["h"] * IN)) / IN
            if dx > 0.02: geom_bad.append(f"p{si}/{m['id']} Δ{dx:.3f}in")
            st = m.get("strategy", "")
            if st == "native-text" and not (sh.has_text_frame and sh.text_frame.text.strip()): text_bad.append(f"p{si}/{m['id']}")
            if m.get("result") == "raster" and sh.shape_type != 13: text_bad.append(f"p{si}/{m['id']} raster expected picture")
            if st == "native-chart" and not getattr(sh, "has_chart", False) and not getattr(sh, "has_table", False): text_bad.append(f"p{si}/{m['id']} chart expected native chart/table")
    expect = sum(1 for m in manifest if m.get("result") not in ("skipped", "dropped"))
    check(res, "PPTX", "every exported element exists as a named object", "PASS" if not missing else "FAIL", f"{matched}/{expect}" + (f"; missing {missing[:5]}" if missing else ""))
    check(res, "PPTX", "geometry within 0.02 in of DOM box", "PASS" if not geom_bad else "FAIL", f"{len(geom_bad)} off" + (f": {geom_bad[:5]}" if geom_bad else ""))
    check(res, "PPTX", "native objects stay native (text/chart editable), rasters are pictures", "PASS" if not text_bad else "FAIL", ", ".join(text_bad[:5]) or "ok")
    check(res, "PPTX", "no full-page screenshot posing as slide", "PASS" if not full_page else "FAIL", ", ".join(full_page) or "no picture ≥ 90% of slide area")
    want_notes = len(prs.slides)
    check(res, "PPTX", "speaker notes carried over", "PASS" if notes == want_notes else "FAIL", f"{notes}/{want_notes}")
    n = sum(kinds.values())
    check(res, "PPTX", "object inventory", "INFO", json.dumps(kinds) + f"; native ≈ {n - kinds['picture']}/{n} objects; pictures = {kinds['picture']} (declared raster/image only)")


def _raster_pages(pdf, shots, out, force, area, limit, res, pages, tag):
    """PDF（Chromium 导出，或 PPTX 经 LibreOffice/PowerPoint 渲染）→ 逐页光栅，与 HTML 终态截图比对。成功返回已打开的 PDF 文档。"""
    try:
        import pymupdf as fitz
        from PIL import Image, ImageChops
    except Exception as e:
        check(res, area, f"{area} raster comparison", "BLOCKED", f"PyMuPDF/Pillow missing: {e}"); return None
    doc = fitz.open(pdf)
    if pages is not None:
        check(res, area, "page count == HTML slides", "PASS" if len(doc) == pages else "FAIL", f"{len(doc)} vs {pages}")
    sizes = {(round(p.rect.width), round(p.rect.height)) for p in doc}
    check(res, area, "page size 1280×720 px (960×540 pt), no blank/extra pages", "PASS" if sizes == {(960, 540)} else "FAIL", str(sizes))
    missing = [i for i in range(1, len(doc) + 1) if not os.path.isfile(os.path.join(shots, f"slide-{i:02d}.png"))]
    if missing:
        check(res, area, "per-page pixel diff vs HTML final-state screenshot", "BLOCKED", f"missing HTML screenshots for pages {missing[:8]} in {shots} (run web_export.mjs audit --shots)")
        return doc
    diffs = []
    for i, p in enumerate(doc, 1):
        pix = p.get_pixmap(matrix=fitz.Matrix(1280 / p.rect.width, 1280 / p.rect.width))
        a = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).resize((1280, 720))
        b = Image.open(os.path.join(shots, f"slide-{i:02d}.png")).convert("RGB").resize((1280, 720))
        d = ImageChops.difference(a, b).convert("L")
        hist = d.histogram(); tot = sum(hist)
        diffs.append((i, sum(hist[40:]) / tot))
        buf = io.BytesIO(); a.save(buf, "PNG"); write_atomic(Path(out) / f"{tag}-{i:02d}.png", buf.getvalue(), force)
        sheet = Image.new("RGB", (2580, 720), (30, 30, 30)); sheet.paste(b, (0, 0)); sheet.paste(a, (1300, 0))
        buf = io.BytesIO(); sheet.resize((1290, 360)).save(buf, "JPEG", quality=82); write_atomic(Path(out) / f"cmp-{tag}-{i:02d}.jpg", buf.getvalue(), force)
    worst = max(diffs, key=lambda x: x[1])
    check(res, area, f"per-page pixel diff vs HTML final-state screenshot (share of pixels with Δ>40/255, limit {limit:.0%})", "PASS" if worst[1] < limit else "FAIL",
          "; ".join(f"p{i}:{b * 100:.2f}%" for i, b in diffs) + f" (worst p{worst[0]})")
    return doc


def verify_pdf(pdf, shots, out, force, manifest, res):
    pages = max((m["page"] for m in manifest), default=None)
    doc = _raster_pages(pdf, shots, out, force, "PDF", PDF_DIFF_LIMIT, res, pages, "pdf")
    if doc is None:
        return
    fonts = sorted({f"{f[3] or f[4]}({f[2]})" for p in doc for f in p.get_fonts()})
    check(res, "PDF", "fonts embedded (Skia embeds CJK as Type3 outlines: selectable, not re-editable)", "PASS" if fonts else "FAIL", f"{len(fonts)} font objects: " + ", ".join(fonts[:4]))
    # 文字量按清单里的期望字数比较（大数字页只有几个字也是合法的），而不是写死“每页 > 20 字”
    got = ["".join(p.get_text().split()) for p in doc]
    want = {}
    for m in manifest:
        want[m["page"]] = want.get(m["page"], 0) + int(m.get("text_len") or 0)
    bad = [i for i, t in enumerate(got, 1) if (want.get(i) and len(t) < .8 * want[i]) or (not want.get(i) and not t)]
    if not any(want.values()):
        check(res, "PDF", "vector selectable text on every page", "PASS" if not bad else "FAIL", f"chars/page {[len(t) for t in got]} (manifest has no text_len: only non-empty checked)")
    else:
        check(res, "PDF", "vector selectable text matches DOM text volume (≥ 80%)", "PASS" if not bad else "FAIL", f"pdf chars {[len(t) for t in got]} vs dom {[want.get(i, 0) for i in range(1, len(got) + 1)]}")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx"); ap.add_argument("--manifest"); ap.add_argument("--pdf"); ap.add_argument("--shots")
    ap.add_argument("--pptx-render", help="PPTX 经 LibreOffice/PowerPoint 渲染得到的 PDF；缺它则 PPTX 视觉还原只能是 UNVERIFIED")
    ap.add_argument("--out", default="verify-out"); ap.add_argument("--json"); ap.add_argument("--force", action="store_true", help="允许覆盖已存在的报告/对比图")
    a = ap.parse_args(argv)
    # 输入/契约错误 → 退出码 2，不写任何文件
    err = []
    if not (a.pptx or a.pdf): err.append("至少给 --pptx 或 --pdf")
    if a.pptx and not a.manifest: err.append("--pptx 需要 --manifest（没有清单无法逐元素核对）")
    if a.pptx_render and not a.pptx: err.append("--pptx-render 需要 --pptx")
    for k in ("pptx", "manifest", "pdf", "pptx_render"):
        v = getattr(a, k)
        if v and not os.path.isfile(v): err.append(f"--{k.replace('_', '-')} 文件不存在: {v}")
    if a.shots and not os.path.isdir(a.shots): err.append(f"--shots 不是目录: {a.shots}")
    rep = Report()
    outdir = Path(a.out).resolve()
    ins = {k: Path(getattr(a, k)) for k in ("pptx", "manifest", "pdf", "pptx_render") if getattr(a, k)}
    outs = [("--json", Path(a.json), ".json")] if a.json else []
    guard_paths(rep, ins, outs, a.force, dirs={"--out": outdir, **({"--shots": Path(a.shots)} if a.shots else {})})
    if a.shots and outdir == Path(a.shots).resolve(): rep.hard("E-EXIST", str(outdir), "--out 不能等于 --shots：对比图不得混进 HTML 终态截图")
    err += [f"{i.code} {i.where}: {i.message}" for i in rep.items if i.level == "HARD"]
    if err:
        for e in err: print("ERROR", e, file=sys.stderr)
        return EXIT_BY_STATUS["ERROR"]
    manifest, flags = load_manifest(a.manifest) if a.manifest else ([], [])
    res, down = [], []
    if a.manifest: down = verify_manifest(manifest, flags, res)
    try:
        if a.pptx: verify_pptx(a.pptx, manifest, res)
    except Exception as e:   # 打不开/损坏的 PPTX 是 FAIL，不是崩溃后的“无结论”
        check(res, "PPTX", "PPTX opens as a valid package", "FAIL", f"{type(e).__name__}: {e}")
    if a.pdf and not a.shots:
        check(res, "PDF", "PDF raster comparison", "BLOCKED", "需要 --shots（web_export.mjs audit 产出的 slide-NN.png）才能逐页比对")
    if (a.pdf or a.pptx_render) and a.shots:
        outdir.mkdir(parents=True, exist_ok=True)
    if a.pdf and a.shots: verify_pdf(a.pdf, a.shots, outdir, a.force, manifest, res)
    if a.pptx:
        if a.pptx_render and a.shots:
            _raster_pages(a.pptx_render, a.shots, outdir, a.force, "PPTX-render", PPTX_DIFF_LIMIT, res, max((m["page"] for m in manifest), default=None), "pptx")
        else:
            check(res, "PPTX", "visual fidelity vs HTML (needs LibreOffice/PowerPoint render: --pptx-render + --shots)", "UNVERIFIED",
                  "未提供 PPTX 的真实渲染；上面的几何/文本/类型检查只是结构层。PowerPoint 实机: NOT VERIFIED")
    status = rollup(res)
    counts = {s: sum(1 for r in res if r["status"] == s) for s in ("PASS", "FAIL", "BLOCKED", "UNVERIFIED", "INFO")}
    out = {"status": status, "exit_code": EXIT_BY_STATUS[status], "counts": counts, "downgrades": down, "checks": res}
    if a.json: write_atomic(Path(a.json), json.dumps(out, ensure_ascii=False, indent=2).encode("utf-8"), a.force)
    for r in res: print(f"{r['status']:10} [{r['area']}] {r['name']} — {r['detail']}")
    print(f"STATUS {status}  exit={EXIT_BY_STATUS[status]}  " + " ".join(f"{k}={v}" for k, v in counts.items()) + f"  downgrades={len(down)}")
    return EXIT_BY_STATUS[status]


if __name__ == "__main__":
    sys.exit(main())
