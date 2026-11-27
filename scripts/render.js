#!/usr/bin/env node
/* PptxGenJS 生成引擎（L3）。输入 spec.py 解析好的 render spec，输出未后处理的 PPTX。
 * 只负责“画”：文字、形状、路径、图片、图表、表格、备注、母版/标题占位符。不含任何设计计算。
 * PptxGenJS 画不出的部分（渐变、蒙版、双色调、软边、分组、图表直标、东亚字体…）由 post.py 按 spec 补齐。
 * 用法：node render.js spec.json out.pptx
 */
'use strict';
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

const [specPath, outPath] = process.argv.slice(2);
if (!specPath || !outPath) { console.error('usage: render.js spec.json out.pptx'); process.exit(2); }
const spec = JSON.parse(fs.readFileSync(specPath, 'utf8'));

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'VAD_16x9', width: spec.canvas.w_in, height: spec.canvas.h_in });
pptx.layout = 'VAD_16x9';
pptx.title = spec.core.title;
pptx.author = spec.core.author;
pptx.company = '';
pptx.theme = { headFontFace: spec.theme.heading, bodyFontFace: spec.theme.body };

const T = spec.theme;
pptx.defineSlideMaster({
  title: 'VAD_MASTER',
  background: { color: T.master_bg },
  objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 0.6, w: 11.7, h: 1.2, fontFace: T.heading,
    fontSize: T.title_pt, bold: true, color: T.master_ink, align: 'left', valign: 'top' }, text: '' } }],
});

const pt2in = (p) => p / 72;
const fillOf = (f) => (f ? { color: f.color, transparency: f.alpha < 1 ? Math.round((1 - f.alpha) * 100) : 0 } : { type: 'none' });
const lineOf = (l) => {
  if (!l) return { type: 'none' };
  const o = { color: l.color, width: l.width_pt, dashType: l.dash };
  if (l.arrow === 'end') o.endArrowheadType = 'triangle';
  if (l.arrow === 'both') { o.endArrowheadType = 'triangle'; o.beginArrowheadType = 'triangle'; }
  return o;
};
const ALIGN = { left: 'left', center: 'center', right: 'right', justify: 'justify' };
const VALIGN = { top: 'top', middle: 'middle', bottom: 'bottom' };

function addText(slide, e) {
  const [x, y, w, h] = e.box;
  const runs = [];
  e.paras.forEach((p, pi_) => {
    const base = {
      lineSpacingMultiple: p.line_spacing, paraSpaceAfter: p.space_after_pt || 0,
    };
    if (p.list) {
      base.bullet = p.list.kind === 'number' ? { type: 'number', indent: p.list.indent_in * 72 }
        : { characterCode: p.list.kind === 'bullet' ? '2022' : '2013', indent: p.list.indent_in * 72 };
    }
    const rs = p.runs.length ? p.runs : [{ text: '', font: p.fallback.font, size: p.fallback.size, color: p.fallback.color, bold: false, spc: 0 }];
    rs.forEach((r, ri) => {
      // 段落级属性只放在段首 run：PptxGenJS 会为每个带段落属性的 run 重复写 a:pPr（非法结构）。
      runs.push({ text: r.text, options: { ...(ri === 0 ? base : {}), fontFace: r.font, fontSize: r.size, color: r.color, bold: !!r.bold,
        charSpacing: r.spc || undefined, breakLine: ri === rs.length - 1 && pi_ < e.paras.length - 1 } });
    });
  });
  const o = { x, y, w, h, margin: [e.pad[0] * 72, e.pad[0] * 72, e.pad[1] * 72, e.pad[1] * 72] /* PptxGenJS 次序：左、右、下、上（pt） */, valign: VALIGN[e.valign], align: ALIGN[e.align], // 对齐写在整框：run 级 align 会让 PptxGenJS 在占位符里拆段
    fit: 'none', wrap: e.wrap, isTextBox: true, objectName: e.name };
  if (e.rotation) o.rotate = e.rotation;
  if (e.placeholder) { o.placeholder = e.placeholder; }
  slide.addText(runs, o);
}

function addShape(slide, e) {
  const [x, y, w, h] = e.box;
  const o = { x, y, w, h, fill: fillOf(e.fill), line: lineOf(e.line), objectName: e.name };
  if (e.rotation) o.rotate = e.rotation;
  slide.addShape(e.prst, o);
}

function addPath(slide, e) {
  const xs = e.points.map((p) => p[0]), ys = e.points.map((p) => p[1]);
  const x0 = Math.min(...xs), y0 = Math.min(...ys);
  const w = Math.max(Math.max(...xs) - x0, 0.01), h = Math.max(Math.max(...ys) - y0, 0.01);
  const line = lineOf(e.line);
  if (e.points.length === 2) {
    const [a, b] = e.points;
    const o = { x: x0, y: y0, w: Math.max(w, 0), h: Math.max(h, 0), line, objectName: e.name, flipH: b[0] < a[0], flipV: b[1] < a[1] };
    if (e.rotation) o.rotate = e.rotation;
    slide.addShape('line', o);
    return;
  }
  const pts = e.points.map((p, i) => ({ x: p[0] - x0, y: p[1] - y0, moveTo: i === 0 }));
  if (e.closed) pts.push({ close: true });
  const o = { x: x0, y: y0, w, h, points: pts, fill: e.closed ? fillOf(e.fill) : { type: 'none' }, line, objectName: e.name };
  if (e.rotation) o.rotate = e.rotation;
  slide.addShape('custGeom', o);
}

function addImage(slide, e) {
  const [x, y, w, h] = e.box;
  const o = { path: e.path, x, y, w, h, objectName: e.name, altText: e.alt || '' };
  if (e.alpha) o.transparency = Math.round((1 - e.alpha) * 100);
  if (e.rotation) o.rotate = e.rotation;
  slide.addImage(o);
}

function addTable(slide, e) {
  const none = { type: 'none' };
  const cell = (text, head, align, ri) => ({
    text, options: { fontFace: e.font, fontSize: e.size, color: head ? e.head_ink : e.body_ink, bold: head, align,
      valign: 'middle', margin: [2.16, 5.76, 2.16, 5.76],
      border: [{ type: 'solid', pt: 0.75, color: ri <= 1 ? e.rule : e.grid }, none, { type: 'solid', pt: 0.75, color: ri <= 1 ? e.rule : e.grid }, none] },
  });
  const rows = [e.headers.map((t, j) => cell(t, true, e.aligns[j], 0))];
  e.rows.forEach((r, i) => rows.push(r.map((t, j) => cell(t, false, e.aligns[j], i + 1))));
  const o = { x: e.box[0], y: e.box[1], w: e.box[2], h: e.box[3], colW: e.col_w, rowH: e.row_h, objectName: e.name };
  if (e.rotation) o.rotate = e.rotation;
  slide.addTable(rows, o);
}

function addChart(slide, e) {
  const st = e.style, fp = e.font_pt;
  const [x, y, w, h] = e.box;
  const k = e.chart_kind;
  const o = {
    x, y, w, h, objectName: e.name, altText: e.alt || '',
    showLegend: e.legend, legendPos: 'b', legendFontFace: e.family, legendFontSize: fp.legend, legendColor: st.body,
    showTitle: !!e.title, title: e.title || '', titleFontFace: e.family, titleFontSize: fp.title, titleColor: st.title, titleBold: false,
    showValue: false, // 直标由 post.py 按 label_plan 逐系列写入（墨色、衬底、位置各自决定）
    catAxisLabelFontFace: e.family, catAxisLabelFontSize: fp.axis, catAxisLabelColor: st.body,
    valAxisLabelFontFace: e.family, valAxisLabelFontSize: fp.axis, valAxisLabelColor: st.body,
    catGridLine: { style: 'none' }, valGridLine: e.axis_labels ? { color: st.grid, size: 0.75, style: 'solid' } : { style: 'none' },
    catAxisLineShow: true, valAxisLineShow: e.axis_labels, valAxisHidden: !e.axis_labels,
    catAxisLabelPos: 'low', lang: 'zh-CN',
  };
  if (e.val_axis.min !== null) o.valAxisMinVal = e.val_axis.min;
  if (e.val_axis.max !== null && e.val_axis.max !== undefined) o.valAxisMaxVal = e.val_axis.max;
  if (e.val_axis.major) o.valAxisMajorUnit = e.val_axis.major;
  if (e.number_format) o.valAxisLabelFormatCode = e.number_format;
  if (e.axis_titles.y) { o.showValAxisTitle = true; o.valAxisTitle = e.axis_titles.y; o.valAxisTitleFontFace = e.family; o.valAxisTitleFontSize = fp.legend; o.valAxisTitleColor = st.body; }
  if (e.axis_titles.x) { o.showCatAxisTitle = true; o.catAxisTitle = e.axis_titles.x; o.catAxisTitleFontFace = e.family; o.catAxisTitleFontSize = fp.legend; o.catAxisTitleColor = st.body; }
  if (e.horizontal) o.catAxisOrientation = 'maxMin';
  const grouping = e.stacked ? (e.stacked === 'percent' ? 'percentStacked' : 'stacked') : 'clustered';
  const series = (names) => e.series.filter((s) => names(s.name));
  const dataOf = (list) => list.map((s) => ({ name: s.name, labels: e.categories, values: s.values }));
  const colorsFor = (list) => list.map((s) => e.colors[e.series.indexOf(s)]);

  if (k === 'pie' || k === 'donut') {
    const s = e.series[0];
    Object.assign(o, { chartColors: e.colors, dataBorder: { pt: 1, color: e.surface }, holeSize: 50, showLeaderLines: true, showPercent: false });
    slide.addChart(k === 'pie' ? pptx.charts.PIE : pptx.charts.DOUGHNUT, dataOf([s]), o);
  } else if (k === 'scatter') {
    const xs = e.series[0].points.map((p) => p[0]);
    for (const s of e.series) {
      if (s.points.length !== xs.length || s.points.some((p, i) => p[0] !== xs[i])) {
        throw new Error('BLOCK E-ENGINE-SCATTER: PptxGenJS 散点各系列须共用同一组 X 值；本图系列 X 不一致，无法可靠生成（不静默降级，请统一 X 或改用 --engine legacy）');
      }
    }
    const data = [{ name: 'X', values: xs }, ...e.series.map((s) => ({ name: s.name, values: s.points.map((p) => p[1]) }))];
    Object.assign(o, { chartColors: e.colors, lineSize: 0, lineDataSymbol: 'circle', lineDataSymbolSize: 7 });
    slide.addChart(pptx.charts.SCATTER, data, o);
  } else if (k === 'line') {
    Object.assign(o, { chartColors: e.colors, lineSize: 2, lineSmooth: false, lineDataSymbol: 'none' });
    slide.addChart(pptx.charts.LINE, dataOf(e.series), o);
  } else if (k === 'area') {
    Object.assign(o, { chartColors: e.colors, barGrouping: e.stacked ? 'stacked' : 'standard' });
    slide.addChart(pptx.charts.AREA, dataOf(e.series), o);
  } else if (e.combo.length) {
    const bars = series((n) => !e.combo.includes(n)), lines = series((n) => e.combo.includes(n));
    const barOpts = { barDir: 'col', barGrouping: grouping, barGapWidthPct: e.gap_width, chartColors: colorsFor(bars) };
    const lineOpts = { chartColors: colorsFor(lines), lineSize: 2.5, lineSmooth: false, lineDataSymbol: 'circle', lineDataSymbolSize: 7 };
    slide.addChart([{ type: pptx.charts.BAR, data: dataOf(bars), options: barOpts },
      { type: pptx.charts.LINE, data: dataOf(lines), options: lineOpts }], o);
  } else {
    Object.assign(o, { barDir: e.horizontal ? 'bar' : 'col', barGrouping: grouping, barGapWidthPct: e.gap_width,
      chartColors: e.point_colors || e.colors });
    if (e.point_colors) o.chartColors = e.point_colors;
    slide.addChart(pptx.charts.BAR, dataOf(e.series), o);
  }
}

function emit(slide, e) {
  switch (e.kind) {
    case 'text': return addText(slide, e);
    case 'shape': return addShape(slide, e);
    case 'path': return addPath(slide, e);
    case 'image': return addImage(slide, e);
    case 'chart': return addChart(slide, e);
    case 'table': return addTable(slide, e);
    case 'group': return e.children.forEach((c) => emit(slide, c)); // 分组在 post.py 里包成 p:grpSp（PptxGenJS 无原生分组）
    default: throw new Error('unknown kind ' + e.kind);
  }
}

try {
for (const page of spec.pages) {
  const slide = pptx.addSlide({ masterName: 'VAD_MASTER' });
  slide.background = { color: page.bg };
  page.elements.forEach((e) => emit(slide, e));
  if (page.page_number) emit(slide, page.page_number);
  if (page.notes) slide.addNotes(page.notes);
}
} catch (err) { console.error(String(err && err.message || err)); process.exit(1); }

pptx.write({ outputType: 'nodebuffer', compression: true }).then((buf) => {
  fs.writeFileSync(outPath, buf);
}).catch((err) => { console.error(String(err && err.message || err)); process.exit(1); });
