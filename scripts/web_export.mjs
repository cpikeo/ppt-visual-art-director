#!/usr/bin/env node
/* Web Presentation → audit / PDF / PPTX.   node scripts/web_export.mjs <audit|pdf|plan|pptx> deck.html [out] [--assets DIR] [--shots DIR]
   Deterministic: DOM (final state, motion off) → per-element strategy (native text/shape/image/chart | raster | skip) → plan.json → compile.py.
   Never screenshots a whole slide into a PPTX. Everything downgraded is written to the report, not hidden. */
import { chromium } from 'playwright';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs'; import path from 'node:path'; import { fileURLToPath, pathToFileURL } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const argv = process.argv.slice(2), [cmd, deckArg] = argv, rest = argv.slice(2);
const outArg = rest[0] && !rest[0].startsWith('--') ? rest[0] : undefined;
const opt = (k, d) => { const i = rest.indexOf(k); return i >= 0 ? rest[i + 1] : d; };
if (!cmd || !deckArg) { console.error('usage: web_export.mjs <audit|pdf|plan|pptx> deck.html [out] [--assets DIR] [--shots DIR]'); process.exit(2); }
const deckPath = path.resolve(deckArg), deckUrl = pathToFileURL(deckPath).href;
const J = (o) => JSON.stringify(o, null, 2);

async function open(browser, { w = 1280, h = 720, dpr = 2, q = 'export', mobile = false } = {}) {
  const ctx = await browser.newContext({ viewport: { width: w, height: h }, deviceScaleFactor: dpr, hasTouch: mobile, reducedMotion: q ? 'reduce' : 'no-preference' });
  const page = await ctx.newPage(); const logs = [];
  page.on('console', m => { if (['error', 'warning'].includes(m.type())) logs.push(`${m.type()}: ${m.text()}`); });
  page.on('pageerror', e => logs.push('pageerror: ' + e.message));
  page.on('requestfailed', r => logs.push('requestfailed: ' + r.url()));
  await page.goto(deckUrl + (q ? `?${q}` : ''), { waitUntil: 'load' });
  await page.waitForSelector('html[data-dk-ready="1"]', { state: 'attached', timeout: 8000 });
  return { ctx, page, logs };
}

/* ---------------- in-page extraction (runs in the browser) ---------------- */
function extractSlide(idx, cfg) {
  const S = document.querySelectorAll('.dk-stage > .slide')[idx];
  const SR = S.getBoundingClientRect(), items = [], flags = [];
  let n = 0;
  const PX = 1 / 96;
  const rgba = (c) => { const m = /rgba?\(([^)]+)\)/.exec(c || ''); if (!m) return null; const p = m[1].split(/[ ,\/]+/).filter(Boolean).map(Number); return { r: p[0], g: p[1], b: p[2], a: p[3] ?? 1 }; };
  const hex = (c) => '#' + [c.r, c.g, c.b].map(v => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
  const rel = (r) => ({ x: r.left - SR.left, y: r.top - SR.top, w: r.width, h: r.height });
  const inch = (b) => ({ x: +(b.x * PX).toFixed(4), y: +(b.y * PX).toFixed(4), w: +(b.w * PX).toFixed(4), h: +(b.h * PX).toFixed(4) });
  const flag = (id, code, msg) => flags.push({ id, code, msg });
  const effOpacity = (el) => { let o = 1; for (let e = el; e && e !== S.parentElement; e = e.parentElement) o *= +getComputedStyle(e).opacity; return o; };
  const natural = (f) => /noto|source ?han|思源/i.test(f) ? 1.448 : /yahei|微软/i.test(f) ? 1.32 : /simsun|simhei|宋体|黑体/i.test(f) ? 1.15 : 1.2;
  const fam = (cs) => cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  const INLINE = new Set(['SPAN', 'B', 'STRONG', 'EM', 'I', 'BR', 'A', 'SMALL', 'SUB', 'SUP', 'MARK', 'U', 'S', 'CODE']);
  const hasBox = (cs) => (rgba(cs.backgroundColor)?.a > 0) || cs.backgroundImage !== 'none' || ['Top', 'Right', 'Bottom', 'Left'].some(s => parseFloat(cs['border' + s + 'Width']) > 0 && cs['border' + s + 'Style'] !== 'none');
  const needsRaster = (el, cs) => {
    if (el.dataset.x === 'raster') return 'data-x=raster';
    if (['SVG', 'CANVAS', 'VIDEO'].includes(el.tagName.toUpperCase())) return el.tagName.toLowerCase() + ' → raster';
    if (cs.filter !== 'none') return 'filter'; if ((cs.backdropFilter || 'none') !== 'none') return 'backdrop-filter';
    if (cs.mixBlendMode !== 'normal') return 'mix-blend-mode'; if (cs.clipPath !== 'none') return 'clip-path';
    if ((cs.maskImage || 'none') !== 'none') return 'mask'; if ((cs.webkitBackgroundClip || cs.backgroundClip) === 'text') return 'background-clip:text';
    if (cs.transform !== 'none') { const m = new DOMMatrix(cs.transform); if (Math.abs(Math.hypot(m.a, m.b) - 1) > .01) return 'transform-scale/skew'; }
    return '';
  };
  const rotation = (cs) => { if (cs.transform === 'none') return 0; const m = new DOMMatrix(cs.transform); return Math.round(Math.atan2(m.b, m.a) * 1800 / Math.PI) / 10; };
  const layoutBox = (el, cs) => { // un-rotated layout box, centred on the visual centre
    const r = el.getBoundingClientRect(), rot = rotation(cs);
    if (!rot) return rel(r);
    const w = el.offsetWidth, h = el.offsetHeight, cx = r.left + r.width / 2 - SR.left, cy = r.top + r.height / 2 - SR.top;
    return { x: cx - w / 2, y: cy - h / 2, w, h };
  };
  const mark = (el, mode = 'tree') => { const id = 'r' + idx + '_' + (++n); el.setAttribute('data-dk-eid', id); return id; };
  const gradient = (bg, id) => {
    const m = /linear-gradient\((.*)\)$/.exec(bg); if (!m) return null;
    const parts = m[1].match(/(?:[^,(]|\([^)]*\))+/g).map(s => s.trim());
    let ang = 180; const first = parts[0];
    if (/deg$/.test(first)) { ang = parseFloat(first); parts.shift(); }
    else if (/^to /.test(first)) { ang = { 'to right': 90, 'to bottom': 180, 'to left': 270, 'to top': 0 }[first] ?? 180; parts.shift(); }
    const stops = parts.map(p => rgba(/rgba?\([^)]*\)/.exec(p)?.[0])).filter(Boolean);
    if (stops.length < 2) return null;
    if (stops.length > 2) flag(id, 'W-GRADIENT', `${stops.length} stops simplified to first/last`);
    const a = stops[0], b = stops[stops.length - 1], plan = ((ang - 90 + 540) % 360) - 180;
    const g = { from: hex(a), to: hex(b), angle: plan }; if (a.a < 1) g.from_alpha = Math.max(.01, +a.a.toFixed(3)); if (b.a < 1) g.to_alpha = Math.max(.01, +b.a.toFixed(3));
    return g;
  };
  const textFrom = (root) => { // → runs [{t,style}] with CSS whitespace collapsed
    const out = [];
    const walk = (node, cs) => {
      if (node.nodeType === 3) { let t = node.textContent; const ws = cs.whiteSpace; if (!/^pre/.test(ws)) t = t.replace(/\s+/g, ' ');
        const tt = cs.textTransform; if (tt === 'uppercase') t = t.toUpperCase(); else if (tt === 'lowercase') t = t.toLowerCase(); if (t) out.push({ t, cs }); }
      else if (node.nodeType === 1) { if (node.tagName === 'BR') { out.push({ t: '\n', cs }); return; } const c = getComputedStyle(node); if (c.display === 'none') return; node.childNodes.forEach(k => walk(k, c)); }
    };
    root.childNodes.forEach(k => walk(k, getComputedStyle(root)));
    return out;
  };
  const emitText = (el, cs, rectOverride) => {
    const runs = textFrom(el); let all = runs.map(r => r.t).join('');
    if (!all.replace(/[\s\n]/g, '')) return;
    // trim edges of first/last runs
    runs[0].t = runs[0].t.replace(/^\s+/, ''); runs[runs.length - 1].t = runs[runs.length - 1].t.replace(/\s+$/, '');
    all = runs.map(r => r.t).join('').replace(/ *\n */g, '\n');
    const op = effOpacity(el), bgc = (() => { for (let e = el; e; e = e.parentElement) { const c = rgba(getComputedStyle(e).backgroundColor); if (c && c.a > .5) return c; } return { r: 255, g: 255, b: 255 }; })();
    const colorOf = (c) => { const k = rgba(c.color); const a = k.a * op; return hex(a < 1 ? { r: k.r * a + bgc.r * (1 - a), g: k.g * a + bgc.g * (1 - a), b: k.b * a + bgc.b * (1 - a) } : k); };
    const px2pt = (v) => +(parseFloat(v) * .75).toFixed(2);
    const base = { fs: px2pt(cs.fontSize), col: colorOf(cs), bold: +cs.fontWeight >= 600, font: fam(cs), ls: cs.letterSpacing === 'normal' ? 0 : +(parseFloat(cs.letterSpacing) * .75).toFixed(2) };
    const rr = runs.filter(r => r.t).map(r => { const c = r.cs; const o = { text: r.t.replace(/ *\n */g, '\n') };
      const f = px2pt(c.fontSize), col = colorOf(c), b = +c.fontWeight >= 600, fm = fam(c);
      if (f !== base.fs) o.font_size = f; if (col !== base.col) o.color = col; if (b !== base.bold) o.bold = b; if (fm !== base.font) o.font = fm; return o; });
    const merged = []; rr.forEach(r => { const p = merged[merged.length - 1]; if (p && JSON.stringify({ ...p, text: 0 }) === JSON.stringify({ ...r, text: 0 })) p.text += r.text; else merged.push({ ...r }); });
    const rng = document.createRange(); rng.selectNodeContents(el);
    const rects = [...rng.getClientRects()].filter(r => r.width > 0 && r.height > 0);
    if (!rects.length) return;
    const ur = rects.reduce((a, r) => ({ l: Math.min(a.l, r.left), t: Math.min(a.t, r.top), r: Math.max(a.r, r.right), b: Math.max(a.b, r.bottom) }), { l: 1e9, t: 1e9, r: -1e9, b: -1e9 });
    const er = el.getBoundingClientRect(), pl = parseFloat(cs.paddingLeft) || 0, pr = parseFloat(cs.paddingRight) || 0;
    let bx = rectOverride ? rel(rectOverride) : { x: er.left - SR.left + pl, y: ur.t - SR.top, w: er.width - pl - pr, h: ur.b - ur.t };
    if (!rectOverride && cs.display.startsWith('inline')) bx = { x: ur.l - SR.left, y: ur.t - SR.top, w: ur.r - ur.l, h: ur.b - ur.t };
    const lhpx = cs.lineHeight === 'normal' ? parseFloat(cs.fontSize) * natural(base.font) : parseFloat(cs.lineHeight);
    const lines = Math.max(1, Math.round(bx.h / lhpx)), grow = lines === 1 ? .08 : .02;
    const align = ({ start: 'left', left: 'left', center: 'center', right: 'right', end: 'right', justify: 'left' })[cs.textAlign] || 'left';
    const dw = bx.w * grow; if (align === 'left') bx.w += dw; else if (align === 'right') { bx.x -= dw; bx.w += dw; } else { bx.x -= dw / 2; bx.w += dw; }
    const tag = el.tagName, type = el.dataset.xType || (/^H[12]$/.test(tag) ? 'title' : 'text');
    const id = el.id || el.dataset.id || 'e' + (++n);
    const ls = lhpx / (parseFloat(cs.fontSize) * natural(base.font));
    const it = { id: 't' + idx + '_' + id.replace(/[^\w-]/g, ''), type, ...inch(bx), content: all, font_size: base.fs, font: base.font, color: base.col, bold: base.bold, align, pad_in: 0,
      line_spacing: +Math.min(3, Math.max(.8, ls)).toFixed(3), valign: 'top' };
    if (base.ls) it.letter_spacing_pt = Math.max(-40, Math.min(40, base.ls));
    if (merged.length > 1 || merged.some(m => Object.keys(m).length > 1)) it.runs = merged;
    const rot = rotation(cs); if (rot) it.rotation = rot;
    if (cs.textShadow !== 'none') flag(it.id, 'W-SHADOW', 'text-shadow dropped (no native equivalent)');
    if (cs.textDecorationLine !== 'none') flag(it.id, 'W-DECOR', 'text-decoration dropped');
    items.push({ item: it, strategy: 'native-text', src: el.tagName.toLowerCase() + (el.className ? '.' + String(el.className).split(' ')[0] : '') });
  };
  const emitBox = (el, cs, op, isSlide) => {
    const b = layoutBox(el, cs); if (b.w < .5 || b.h < .5) return;
    const id = el.id || el.dataset.id || 'b' + (++n), pid = 's' + idx + '_' + id.replace(/[^\w-]/g, '');
    const bg = isSlide ? null : rgba(cs.backgroundColor), bgImg = cs.backgroundImage;
    const rad = isSlide ? 0 : parseFloat(cs.borderTopLeftRadius) || 0, mn = Math.min(b.w, b.h);
    const shape = (rad >= mn * .48 && Math.abs(b.w - b.h) < mn * .15) || cs.borderTopLeftRadius.includes('%') && parseFloat(cs.borderTopLeftRadius) >= 50 ? 'oval' : rad > 1 ? 'round_rect' : 'rect';
    const base = { type: 'shape', shape_kind: shape, ...inch(b), overlap_allowed: true };
    if (shape === 'round_rect') base.corner_in = +Math.min(rad, mn / 2) * PX;
    const rot = rotation(cs); if (rot) base.rotation = rot;
    let emitted = false;
    if (/^url\(/.test(bgImg)) {
      const url = /url\(["']?([^"')]+)["']?\)/.exec(bgImg)[1];
      items.push({ item: { id: pid + '_bg', type: 'image', ...inch(b), overlap_allowed: true, image_role: el.dataset.xRole || 'background', info_duty: el.dataset.duty || 'atmosphere', alpha: op < 1 ? +op.toFixed(3) : undefined, _src: url, _pos: cs.backgroundPosition, alt: el.dataset.alt, asset_reason: el.dataset.reason || 'CSS background image' }, strategy: 'native-image', src: 'css-background' });
      emitted = true;
    } else if (bgImg !== 'none') {
      const g = gradient(bgImg, pid);
      if (g) { items.push({ item: { ...base, id: pid, gradient: g }, strategy: 'native-shape(gradient)', src: 'css-gradient' }); emitted = true; }
      else { const rid = mark(el); flag(pid, 'W-RASTER', 'complex background-image rasterized (self only)'); items.push({ item: { id: pid, type: 'image', ...inch(b), overlap_allowed: true, image_role: 'texture', info_duty: 'atmosphere', _raster: rid, _mode: 'self', alt: '' }, strategy: 'raster', src: 'complex-bg' }); emitted = true; }
    }
    if (!emitted && bg && bg.a > 0) {
      const a = +(bg.a * op).toFixed(3);
      const it = { ...base, id: pid, color: hex(bg) }; if (a < 1) it.fill_alpha = a;
      items.push({ item: it, strategy: 'native-shape', src: 'css-box' }); emitted = true;
    }
    const sides = isSlide ? [] : ['Top', 'Right', 'Bottom', 'Left'].map(s => ({ s, w: parseFloat(cs['border' + s + 'Width']) || 0, c: rgba(cs['border' + s + 'Color']), st: cs['border' + s + 'Style'] })).filter(x => x.w > 0 && x.st !== 'none' && x.c && x.c.a > 0) ;
    if (sides.length === 4 && sides.every(x => x.w === sides[0].w && hex(x.c) === hex(sides[0].c))) {
      const last = items[items.length - 1]; const st = { color: hex(sides[0].c), width_pt: +(sides[0].w * .75).toFixed(2) };
      if (emitted && last.item.id === pid) last.item.stroke = st; else items.push({ item: { ...base, id: pid + '_bd', stroke: st }, strategy: 'native-shape(stroke)', src: 'css-border' });
    } else sides.forEach(x => { // single-side rules → thin rects
      const t = { Top: [0, 0, b.w, x.w], Bottom: [0, b.h - x.w, b.w, x.w], Left: [0, 0, x.w, b.h], Right: [b.w - x.w, 0, x.w, b.h] }[x.s];
      items.push({ item: { type: 'shape', shape_kind: 'rect', overlap_allowed: true, id: pid + '_' + x.s[0], ...inch({ x: b.x + t[0], y: b.y + t[1], w: t[2], h: t[3] }), color: hex(x.c) }, strategy: 'native-shape(rule)', src: 'css-border-' + x.s.toLowerCase() });
    });
    if (cs.boxShadow !== 'none') flag(pid, 'W-SHADOW', 'box-shadow dropped (HTML/PDF only)');
  };
  const walk = (el, top) => {
    if (el.nodeType !== 1 || el.closest('aside.notes') || el.classList.contains('notes')) return;
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden' || el.dataset.x === 'skip') { if (el.dataset.x === 'skip') flag(el.id || el.className || el.tagName, 'I-SKIP', 'data-x=skip → HTML/PDF only'); return; }
    const op = effOpacity(el); if (op < .01) return;
    if (el.dataset.chart) {
      let chart; try { chart = JSON.parse(el.dataset.chart); } catch (e) { flag(el.id, 'E-CHART', 'invalid data-chart JSON'); return; }
      const id = 'c' + idx + '_' + (el.id || ++n);
      items.push({ item: { id, type: 'chart', ...inch(rel(el.getBoundingClientRect())), chart, alt: el.dataset.alt || chart.title || 'chart', overlap_allowed: true }, strategy: 'native-chart', src: 'data-chart' }); return;
    }
    const why = top ? '' : needsRaster(el, cs);
    if (why) {
      const rid = mark(el), b = rel(el.getBoundingClientRect());
      flag(rid, 'W-RASTER', `${el.tagName.toLowerCase()}#${el.id || ''} rasterized (${why}); not editable`);
      items.push({ item: { id: 'x' + rid, type: 'image', ...inch(b), overlap_allowed: true, image_role: 'illustration', info_duty: el.dataset.duty || 'structure', alt: el.dataset.alt || el.getAttribute('aria-label') || '', asset_reason: `rasterized ${why}`, _raster: rid, _mode: 'tree' }, strategy: 'raster', src: why }); return;
    }
    if (el.tagName === 'IMG') {
      const b = layoutBox(el, cs), fit = cs.objectFit, ar = el.naturalWidth / el.naturalHeight, br = b.w / b.h, rad = parseFloat(cs.borderTopLeftRadius) || 0;
      if (rad > 0 || (fit !== 'cover' && Math.abs(ar - br) > .02)) { const rid = mark(el); flag(rid, 'W-RASTER', `img rasterized (${rad > 0 ? 'border-radius' : 'object-fit:' + fit + ' ≠ cover'})`); items.push({ item: { id: 'x' + rid, type: 'image', ...inch(b), overlap_allowed: true, image_role: el.dataset.xRole || 'illustration', info_duty: el.dataset.duty || 'atmosphere', alt: el.alt, asset_reason: 'rasterized img', _raster: rid, _mode: 'tree' }, strategy: 'raster', src: 'img' }); return; }
      items.push({ item: { id: 'i' + idx + '_' + (el.id || ++n), type: 'image', ...inch(b), overlap_allowed: true, image_role: el.dataset.xRole || 'illustration', info_duty: el.dataset.duty || 'atmosphere', alpha: op < 1 ? +op.toFixed(3) : undefined, _src: el.currentSrc || el.src, _pos: cs.objectPosition, alt: el.alt || undefined, asset_reason: el.dataset.reason || el.alt || 'image' }, strategy: 'native-image', src: 'img' }); return;
    }
    if (hasBox(cs) && (!top || cs.backgroundImage !== 'none')) emitBox(el, cs, op, top);
    const kids = [...el.childNodes], textual = kids.some(k => k.nodeType === 3 && k.textContent.trim());
    const plain = kids.every(k => k.nodeType === 3 || k.nodeType === 8 || (k.nodeType === 1 && INLINE.has(k.tagName) && !hasBox(getComputedStyle(k)) && !needsRaster(k, getComputedStyle(k)) && !k.dataset.chart && k.dataset.x !== 'skip'));
    if (textual && plain) { emitText(el, cs); return; }
    kids.forEach(k => {
      if (k.nodeType === 1) walk(k, false);
      else if (k.nodeType === 3 && k.textContent.trim()) { const w = document.createElement('span'); k.replaceWith(w); w.append(k); walk(w, false); } /* stray text among blocks */
    });
  };
  const scs = getComputedStyle(S), bgc = rgba(scs.backgroundColor);
  walk(S, true);
  return { items, flags, bg: bgc && bgc.a > 0 ? hex(bgc) : '#FFFFFF', role: S.dataset.role || 'build', focus: S.dataset.focus || '', rationale: S.dataset.rationale || '', concept: S.dataset.concept || '', notes: (S.querySelector('aside.notes')?.textContent || S.dataset.notes || '').trim(), title: (S.querySelector('h1,h2')?.textContent || '').trim().replace(/\s+/g, ' '), id: S.id || 'p' + (idx + 1) };
}

/* ---------------- commands ---------------- */
async function buildPlan(browser, assetsDir) {
  const { page, logs } = await open(browser, { dpr: 2, q: 'export' });
  const meta = await page.evaluate(() => JSON.parse(document.getElementById('deck-meta')?.textContent || '{}'));
  const count = await page.evaluate(() => window.Deck.count);
  fs.mkdirSync(assetsDir, { recursive: true });
  await page.addStyleTag({ content: `html.dk-iso,html.dk-iso body,html.dk-iso .dk-stage,html.dk-iso .dk-viewport,html.dk-iso .dk-fit{background:transparent!important}html.dk-iso .slide{background:transparent!important}html.dk-iso .slide *{visibility:hidden!important}html.dk-iso .slide [data-dk-iso],html.dk-iso .slide [data-dk-iso] *{visibility:visible!important}html.dk-iso .slide [data-dk-iso=self] *{visibility:hidden!important}` });
  const assets = {}, pages = [], manifest = [], report = { flags: [], logs };
  let aid = 0;
  for (let i = 0; i < count; i++) {
    await page.evaluate((k) => window.Deck.show(k), i);
    await page.waitForTimeout(60);
    const ex = await page.evaluate(extractSlide, i);
    const els = [];
    for (const e of ex.items) {
      const it = e.item;
      if (it._raster) { // isolated element screenshot with transparent background
        await page.evaluate(([id, mode]) => { document.documentElement.classList.add('dk-iso'); document.querySelector(`[data-dk-eid="${id}"]`).setAttribute('data-dk-iso', mode); }, [it._raster, it._mode]);
        const file = `${it._raster}.png`, clip = { x: Math.max(0, it.x * 96), y: Math.max(0, it.y * 96), width: Math.min(1280, (it.x + it.w) * 96) - Math.max(0, it.x * 96), height: Math.min(720, (it.y + it.h) * 96) - Math.max(0, it.y * 96) };
        await page.screenshot({ path: path.join(assetsDir, file), clip, omitBackground: true });
        await page.evaluate(() => { document.documentElement.classList.remove('dk-iso'); document.querySelectorAll('[data-dk-iso]').forEach(e => e.removeAttribute('data-dk-iso')); });
        const id = it.id; assets[id] = { source: 'rasterized from web deck (element-level)', rights: 'cleared', generated: false, reference: it._raster }; it.image_id = id;
        it.x = +(clip.x / 96).toFixed(4); it.y = +(clip.y / 96).toFixed(4); it.w = +(clip.width / 96).toFixed(4); it.h = +(clip.height / 96).toFixed(4);
        fs.renameSync(path.join(assetsDir, file), path.join(assetsDir, id + '.png'));
      } else if (it._src) {
        const id = 'img' + (++aid); let buf, ext;
        if (it._src.startsWith('data:')) { const m = /^data:image\/(\w+)[^,]*;base64,(.*)$/.exec(it._src); ext = m[1].replace('jpeg', 'jpg'); buf = Buffer.from(m[2], 'base64'); }
        else { const p = it._src.startsWith('file:') ? fileURLToPath(it._src) : path.resolve(path.dirname(deckPath), it._src); ext = path.extname(p).slice(1).toLowerCase(); buf = fs.readFileSync(p); }
        if (ext === 'svg' || !['png', 'jpg', 'jpeg', 'webp'].includes(ext)) { report.flags.push({ id: it.id, code: 'E-IMGFMT', msg: `unsupported native image format .${ext}` }); continue; }
        fs.writeFileSync(path.join(assetsDir, `${id}.${ext}`), buf);
        assets[id] = { source: `web deck asset ${path.basename(String(it._src).slice(0, 80))}`, rights: 'pending', reference: String(it._src).slice(0, 60) }; it.image_id = id;
        const pos = (it._pos || '50% 50%').split(' ').map(v => v.endsWith('%') ? parseFloat(v) / 100 : v === 'left' || v === 'top' ? 0 : v === 'right' || v === 'bottom' ? 1 : .5);
        it.crop_focus = { x: pos[0] ?? .5, y: pos[1] ?? .5 };
      }
      for (const k of Object.keys(it)) if (k.startsWith('_') || it[k] === undefined) delete it[k];
      els.push(it); manifest.push({ page: i + 1, id: it.id, type: it.type, strategy: e.strategy, src: e.src, x: it.x, y: it.y, w: it.w, h: it.h });
    }
    ex.flags.forEach(f => report.flags.push({ page: i + 1, ...f }));
    pages.push({ id: ex.id, role: ex.role, focus: ex.focus || `【待判断】${ex.title}`, composition_rationale: ex.rationale || undefined, concept: ex.concept || undefined, background: ex.bg, speaker_notes: ex.notes || undefined, design_state: ex.focus ? 'judged' : 'draft', composition: { strategy: 'free', source: 'author' }, elements: els });
  }
  const plan = { schema_version: 2, deck_id: meta.deck_id || path.basename(deckPath, '.html'), title: meta.title || 'Untitled', audience: meta.audience || '未声明', goal: meta.goal || '未声明', core_claim: meta.core_claim || '未声明',
    brief: meta.brief || { insight: '未声明', concept: '未声明', veto: ['未声明'] }, page_numbers: { hide: pages.map(p => p.id) }, fonts: meta.fonts, palette: meta.palette, assets, pages };
  Object.keys(plan).forEach(k => plan[k] === undefined && delete plan[k]);
  await page.context().close();
  return { plan, manifest, report };
}

const browser = await chromium.launch();
try {
  if (cmd === 'pdf') {
    const { page, logs } = await open(browser, { q: 'export' });
    await page.emulateMedia({ media: 'print' });
    await page.pdf({ path: path.resolve(outArg), width: '1280px', height: '720px', printBackground: true, preferCSSPageSize: true });
    console.log(J({ status: 'DONE', pdf: outArg, console: logs }));
  } else if (cmd === 'plan' || cmd === 'pptx') {
    const out = path.resolve(outArg), assetsDir = path.resolve(opt('--assets', path.join(path.dirname(out), 'assets')));
    const planPath = cmd === 'plan' ? out : out.replace(/\.pptx$/i, '.plan.json');
    const { plan, manifest, report } = await buildPlan(browser, assetsDir);
    fs.writeFileSync(planPath, J(plan)); fs.writeFileSync(planPath.replace(/\.json$/, '.manifest.json'), J({ manifest, flags: report.flags }));
    let compile = null;
    if (cmd === 'pptx') {
      const r = spawnSync(process.env.PYTHON || 'python3', [path.join(HERE, 'compile.py'), planPath, out, '--assets', assetsDir, '--web', '--force'], { encoding: 'utf8' });
      compile = { exit: r.status, log: (r.stdout + r.stderr).split('\n').filter(l => /^\[(HARD|ERROR|SOFT|PENDING|INFO)|E-|已写出/.test(l)).slice(0, 40) };
    }
    const by = {}; manifest.forEach(m => by[m.strategy] = (by[m.strategy] || 0) + 1);
    console.log(J({ status: compile && compile.exit !== 0 ? 'FAIL' : 'DONE', plan: planPath, pptx: cmd === 'pptx' ? out : undefined, elements: manifest.length, by_strategy: by, flags: report.flags.length, compile }));
    if (compile && compile.exit !== 0) process.exitCode = 1;
  } else if (cmd === 'audit') {
    const shotsDir = path.resolve(opt('--shots', 'audit-shots')); fs.mkdirSync(shotsDir, { recursive: true });
    const R = { deck: deckArg, checks: [], slides: [] };
    const add = (area, name, status, detail) => R.checks.push({ area, name, status, detail });
    // 1 static audit at final state
    const { page, logs } = await open(browser, { q: 'export' });
    const count = await page.evaluate(() => window.Deck.count);
    for (let i = 0; i < count; i++) {
      await page.evaluate((k) => window.Deck.show(k), i);
      const s = await page.evaluate((idx) => {
        const S = document.querySelectorAll('.dk-stage > .slide')[idx], SR = S.getBoundingClientRect(), out = { overflow: [], small: [], broken: [], contrast: [], fonts: [], motion: { n: 0, steps: 0, ends: 0, bad: [] }, textLen: 0 };
        const rgba = (c) => { const m = /rgba?\(([^)]+)\)/.exec(c); if (!m) return null; const p = m[1].split(/[ ,\/]+/).filter(Boolean).map(Number); return { r: p[0], g: p[1], b: p[2], a: p[3] ?? 1 }; };
        const lum = (c) => { const f = v => { v /= 255; return v <= .03928 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4; }; return .2126 * f(c.r) + .7152 * f(c.g) + .0722 * f(c.b); };
        const avail = (name) => { const t = 'mmmmwwwiii测试あ한', c = document.createElement('canvas').getContext('2d'); const w = (f) => { c.font = `40px ${f}`; return c.measureText(t).width; }; return w(`"${name}", monospace`) !== w('monospace') || w(`"${name}", serif`) !== w('serif'); };
        const seen = new Set();
        S.querySelectorAll('*').forEach(el => {
          const cs = getComputedStyle(el); if (cs.display === 'none' || el.closest('aside.notes,[aria-hidden=true]') || el.dataset.x === 'skip') return;
          const r = el.getBoundingClientRect(), id = (el.tagName.toLowerCase() + (el.id ? '#' + el.id : el.className && typeof el.className === 'string' ? '.' + el.className.split(' ')[0] : ''));
          const own = [...el.childNodes].some(k => k.nodeType === 3 && k.textContent.trim());
          if (own) {
            out.textLen += el.textContent.length;
            const fs = parseFloat(cs.fontSize); if (fs < 16 && !el.closest('[data-x=skip]')) out.small.push(`${id} ${fs}px`);
            if (el.scrollWidth > el.clientWidth + 2 && cs.overflow !== 'visible') out.overflow.push(`${id} scrollW ${el.scrollWidth}>${el.clientWidth}`);
            if (el.scrollHeight > el.clientHeight + 2 && cs.overflowY !== 'visible' && cs.display !== 'inline') out.overflow.push(`${id} scrollH ${el.scrollHeight}>${el.clientHeight}`);
            const f = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(); if (!seen.has(f)) { seen.add(f); if (!/^(serif|sans-serif|monospace|system-ui|ui-.*)$/.test(f) && !avail(f)) out.fonts.push(f); }
            // contrast vs nearest opaque ancestor background
            let bg = null, unknown = false; for (let e = el; e && !bg; e = e.parentElement) { const c2 = getComputedStyle(e); if (c2.backgroundImage !== 'none') { unknown = true; break; } const c = rgba(c2.backgroundColor); if (c && c.a > .95) bg = c; }
            const fg = rgba(cs.color);
            if (bg && fg && !unknown) { const a = lum(fg) + .05, b = lum(bg) + .05, ratio = Math.max(a, b) / Math.min(a, b), large = fs >= 24 || (fs >= 18.66 && +cs.fontWeight >= 700); if (ratio < (large ? 3 : 4.5) * (fg.a < 1 ? 1 : 1)) out.contrast.push(`${id} ${ratio.toFixed(2)}:1 (${fs}px)`); }
          }
          const out2 = (r.left < SR.left - 2 || r.top < SR.top - 2 || r.right > SR.right + 2 || r.bottom > SR.bottom + 2);
          if (out2 && own && !el.hasAttribute('data-bleed') && !el.closest('[data-bleed]')) out.overflow.push(`${id} text outside stage`);
          if (el.tagName === 'IMG' && (!el.complete || el.naturalWidth === 0)) out.broken.push(el.src.slice(-60));
          if (el.tagName === 'IMG' && el.naturalWidth && el.naturalWidth < r.width * 1.0) out.small.push(`img ${el.naturalWidth}px < ${Math.round(r.width)} rendered`);
          if (el.dataset.m) { out.motion.n++; if (!['fade', 'rise', 'sink', 'slide', 'scale', 'wipe', 'mask', 'grow', 'growx', 'draw', 'count', 'drift'].includes(el.dataset.m)) out.motion.bad.push(el.dataset.m);
            const end = (+(el.dataset.at ?? .6)) + (+(el.dataset.dur ?? .72)); if (el.dataset.m !== 'drift') out.motion.ends = Math.max(out.motion.ends, end); }
          if (el.dataset.step) out.motion.steps = Math.max(out.motion.steps, +el.dataset.step);
        });
        out.hasTitle = !!S.querySelector('h1,h2'); out.transition = S.dataset.transition || null; out.role = S.dataset.role || null; out.hasNotes = !!(S.querySelector('aside.notes') || S.dataset.notes);
        return out;
      }, i);
      await page.screenshot({ path: path.join(shotsDir, `slide-${String(i + 1).padStart(2, '0')}.png`) });
      R.slides.push({ page: i + 1, ...s });
    }
    const all = (k) => R.slides.flatMap(s => s[k].map(x => `p${s.page} ${x}`));
    add('Web', 'deck loads, slides found', count > 0 ? 'PASS' : 'FAIL', `${count} slides`);
    add('Web', 'no console errors / failed requests', logs.length ? 'FAIL' : 'PASS', logs.slice(0, 5).join(' | ') || 'clean');
    add('设计', 'text overflow / outside stage', all('overflow').length ? 'FAIL' : 'PASS', all('overflow').slice(0, 6).join('; ') || 'none');
    add('设计', 'min text size ≥ 16px / image resolution', all('small').length ? 'WARN' : 'PASS', all('small').slice(0, 6).join('; ') || 'none');
    add('设计', 'broken images', all('broken').length ? 'FAIL' : 'PASS', all('broken').join('; ') || 'none');
    add('设计', 'text contrast on flat backgrounds (WCAG AA; image/gradient bg = UNVERIFIED)', all('contrast').length ? 'FAIL' : 'PASS', all('contrast').slice(0, 6).join('; ') || 'none');
    const fonts = [...new Set(R.slides.flatMap(s => s.fonts))];
    add('PDF/PPTX', 'declared fonts installed on this machine', fonts.length ? 'WARN' : 'PASS', fonts.length ? `missing: ${fonts.join(', ')} → substituted in PDF; PPTX depends on the target machine` : 'all resolved');
    add('设计', 'every slide has a title (h1/h2) for outline + a11y', R.slides.every(s => s.hasTitle) ? 'PASS' : 'FAIL', R.slides.filter(s => !s.hasTitle).map(s => 'p' + s.page).join(',') || 'ok');
    const badm = R.slides.flatMap(s => s.motion.bad.map(x => `p${s.page} ${x}`));
    add('动态', 'motion tokens valid', badm.length ? 'FAIL' : 'PASS', badm.join('; ') || 'ok');
    const slow = R.slides.filter(s => s.motion.ends > 3.2).map(s => `p${s.page} ${s.motion.ends.toFixed(1)}s`);
    add('动态', 'entrance choreography ≤ 3.2 s per slide (audience must not wait)', slow.length ? 'WARN' : 'PASS', slow.join('; ') || `max ${Math.max(...R.slides.map(s => s.motion.ends)).toFixed(2)}s`);
    add('动态', 'every slide declares role + notes (narrative intent / presenter)', R.slides.every(s => s.role && s.hasNotes) ? 'PASS' : 'WARN', R.slides.filter(s => !s.role || !s.hasNotes).map(s => 'p' + s.page).join(',') || 'ok');
    await page.context().close();
    // 2 real playthrough (motion ON): keyboard through every slide and build step, then back, overview, presenter, hash recovery
    {
      const { page: p2, logs: l2 } = await open(browser, { q: '', dpr: 1 });
      const t0 = Date.now(); const seq = [];
      await p2.evaluate(() => window.Deck.go(0, { instant: true, force: true }));
      const total = await p2.evaluate(() => window.Deck.count);
      for (let k = 0; k < total * 4 + 4; k++) { const st = await p2.evaluate(() => [window.Deck.index]); seq.push(st[0]); await p2.keyboard.press('ArrowRight'); await p2.waitForTimeout(90); if (st[0] === total - 1 && (await p2.evaluate(() => window.Deck.index)) === total - 1 && k > total) break; }
      await p2.waitForTimeout(900);
      const reachedEnd = await p2.evaluate(() => window.Deck.index === window.Deck.count - 1);
      await p2.keyboard.press('Home'); await p2.waitForTimeout(100); const home = await p2.evaluate(() => window.Deck.index);
      await p2.keyboard.press('End'); await p2.waitForTimeout(100); const end = await p2.evaluate(() => window.Deck.index);
      await p2.keyboard.press('ArrowLeft'); await p2.waitForTimeout(900); const back = await p2.evaluate(() => window.Deck.index);
      await p2.keyboard.press('o'); const ov = await p2.evaluate(() => document.querySelectorAll('.dk-thumb').length);
      await p2.keyboard.press('Escape'); await p2.keyboard.press('p'); const pn = await p2.evaluate(() => document.querySelector('.dk-panel').classList.contains('is-open')); await p2.keyboard.press('p');
      const jt = Math.min(3, total); await p2.keyboard.type(String(jt)); await p2.keyboard.press('Enter'); await p2.waitForTimeout(900); const jump = await p2.evaluate(() => window.Deck.index);
      await p2.reload({ waitUntil: 'load' }); await p2.waitForSelector('html[data-dk-ready="1"]', { state: 'attached' }); const rec = await p2.evaluate(() => window.Deck.index);
      const errs = await p2.evaluate(() => window.Deck.errors);
      add('Web', 'keyboard playthrough start→end with motion on', reachedEnd ? 'PASS' : 'FAIL', `${seq.length} presses, ${Date.now() - t0} ms`);
      add('Web', 'Home/End/←/jump(n⏎)', home === 0 && end === total - 1 && back === total - 2 && jump === jt - 1 ? 'PASS' : 'FAIL', `home=${home} end=${end} back=${back} jump=${jump}`);
      add('Web', 'overview (O) + presenter panel (P)', ov === total && pn ? 'PASS' : 'FAIL', `thumbs=${ov}/${total}, panel=${pn}`);
      add('Web', 'refresh recovery (hash/localStorage)', rec === jt - 1 ? 'PASS' : 'FAIL', `restored index ${rec}`);
      add('动态', 'runtime motion errors / console during playthrough', errs || l2.length ? 'FAIL' : 'PASS', l2.slice(0, 4).join(' | ') || `Deck.errors=${errs}`);
      // frame pacing probe during a push transition
      const fp = await p2.evaluate(async () => { window.Deck.go(0, { instant: true, force: true }); await new Promise(r => setTimeout(r, 200)); const ts = []; let on = true; const loop = (t) => { ts.push(t); if (on) requestAnimationFrame(loop); }; requestAnimationFrame(loop); window.Deck.go(1); await new Promise(r => setTimeout(r, 1800)); on = false; const d = ts.slice(1).map((t, i) => t - ts[i]); return { frames: ts.length, avg: d.reduce((a, b) => a + b, 0) / d.length, p95: d.sort((a, b) => a - b)[Math.floor(d.length * .95)] }; });
      add('性能', 'frame pacing during transition (headless software GL; indicative only)', fp.p95 < 50 ? 'PASS' : 'WARN', `avg ${fp.avg.toFixed(1)} ms, p95 ${fp.p95.toFixed(1)} ms over ${fp.frames} frames`);
      R.frame = fp;
      await p2.context().close();
    }
    // 3 responsive
    R.responsive = [];
    for (const [w, h, name] of [[375, 812, 'mobile'], [768, 1024, 'tablet'], [1440, 900, 'desktop']]) {
      const { page: p3, ctx, logs: l3 } = await open(browser, { w, h, dpr: 1, q: 'still', mobile: w < 800 });
      await p3.evaluate(() => window.Deck.go(1, { instant: true, force: true })); await p3.waitForTimeout(200);
      const m = await p3.evaluate(() => { const st = document.querySelector('.dk-stage').getBoundingClientRect(), bar = document.querySelector('.dk-bar').getBoundingClientRect(), bs = [...document.querySelectorAll('.dk-bar button')].filter(b => b.offsetParent).map(b => b.getBoundingClientRect()); return { sw: document.scrollingElement.scrollWidth, iw: innerWidth, stage: [Math.round(st.width), Math.round(st.height), Math.round(st.left), Math.round(st.top)], barTop: bar.top, stageBottom: st.bottom, minBtn: Math.min(...bs.map(b => Math.min(b.width, b.height))), scale: +getComputedStyle(document.documentElement).getPropertyValue('--dk-s') }; });
      await p3.screenshot({ path: path.join(shotsDir, `resp-${w}.png`) });
      const fits = m.sw <= m.iw && m.stage[2] >= -1 && m.stage[2] + m.stage[0] <= w + 1 && m.stageBottom <= m.barTop + 1;
      R.responsive.push({ w, h, name, ...m, fits });
      add('响应式', `${w}px: stage fits, no horizontal scroll, controls clear of stage, buttons ≥ 32px`, fits && m.minBtn >= 32 ? 'PASS' : 'FAIL', `stage ${m.stage[0]}×${m.stage[1]} scale ${m.scale.toFixed(3)}; minBtn ${m.minBtn}px; scrollW ${m.sw}/${m.iw}${m.scale < .4 ? ' · text <40% scale: preview-grade legibility' : ''}`);
      await ctx.close();
    }
    fs.writeFileSync(path.join(shotsDir, 'audit.json'), J(R));
    const bad = R.checks.filter(c => c.status === 'FAIL').length;
    console.log(J({ status: bad ? 'FAIL' : 'PASS', checks: R.checks, shots: shotsDir }));
    if (bad) process.exitCode = 1;
  } else { console.error('unknown command ' + cmd); process.exitCode = 2; }
} finally { await browser.close(); }
