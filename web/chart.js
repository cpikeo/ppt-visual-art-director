/* data-chart → HTML 视图。与 PPTX 导出读取同一份 JSON，且缺省值只来自 web/contract.js（由 scripts/limits.json 生成）。
   window.DeckChart.resolve(chart, {meta, mode}) 是 HTML 的唯一“解析”入口：web_export.mjs 把它的结果显式写进 plan，
   PPTX 因此不再自行推导 legend / data_labels / palette / color_map / unit / 轴起点。
   支持 bar（竖/横、多系列）、line、pie、donut、table；area / scatter 由 deck 自己画（宿主保留 data-chart，导出仍是原生图表）。
   数值轴：柱/面积/线数据全非负且未写 value_axis.min 时从 0 起（契约 axis_zero_based_kinds），不截轴。须在 deck.js 之前、contract.js 之后加载。 */
(() => {
  const CT = window.DK_CONTRACT;
  let META = {}; try { META = JSON.parse(document.getElementById('deck-meta')?.textContent || '{}'); } catch (e) { /* 无 meta 用契约默认色 */ }
  const NS = 'http://www.w3.org/2000/svg';
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
  const fail = (host, why) => { host.dataset.chartError = why; };
  const lum = (hex) => { const f = (v) => { v /= 255; return v <= .03928 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4; }; const n = parseInt(hex.slice(1), 16); return .2126 * f(n >> 16 & 255) + .7152 * f(n >> 8 & 255) + .0722 * f(n & 255); };
  const hexOf = (css) => { const m = /rgba?\(([^)]+)\)/.exec(css || ''); if (!m) return null; const p = m[1].split(/[ ,\/]+/).filter(Boolean).map(Number); if ((p[3] ?? 1) < .5) return null; return '#' + p.slice(0, 3).map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase(); };
  const surfaceMode = (host) => { for (let e = host; e; e = e.parentElement) { const h = hexOf(getComputedStyle(e).backgroundColor); if (h) return lum(h) < CT.dark_luminance ? 'dark' : 'light'; } return 'light'; };
  const pick = (table, kind) => (kind in table ? table[kind] : table['*']);
  const nums = (c) => ((c.data || {}).series || []).flatMap((s) => s.values || []).filter((v) => typeof v === 'number' && Number.isFinite(v));
  const niceFloor = (v) => { const m = Math.abs(v), e = 10 ** Math.floor(Math.log10(m)); for (const k of [1, 2, 2.5, 5, 10]) if (k * e >= m) return -k * e; return -10 * e; };

  /* 解析：契约 + 图自带声明 → 显式值。字段与 plan.schema.json 的 chart 完全同名。 */
  function resolve(c, o = {}) {
    const meta = o.meta || META, kind = c.kind, d = c.data || {};
    if (kind === 'table') return { kind, legend: false, data_labels: false, items: [], palette: [], color_map: {}, unit: '' };
    const order = pick(CT.contract.item_order, kind);
    const items = order === 'categories' ? (d.categories || []).map(String) : (d.series || []).map((s) => String(s.name));
    const palette = ((c.palette && c.palette.length ? c.palette : meta.palette && meta.palette.length ? meta.palette : CT.palette)).map((x) => String(x).toUpperCase());
    const explicit = Object.fromEntries(Object.entries(c.color_map || {}).map(([k, v]) => [String(k), String(v).toUpperCase()]));
    const color_map = Object.fromEntries(items.map((n, i) => [n, explicit[n] || palette[i % palette.length]]));
    const nSeries = (d.series || []).length, legRule = pick(CT.contract.legend_default, kind);
    const legend = c.legend != null ? !!c.legend : legRule === 'multi_series' ? nSeries > 1 : !!legRule;
    const data_labels = c.data_labels != null ? !!c.data_labels : !!pick(CT.contract.data_labels_default, kind);
    const units = [...new Set((d.series || []).map((s) => s.unit && String(s.unit).trim()).filter(Boolean))];
    const unit = String(c.unit || '').trim() || (units.length === 1 ? units[0] : '');
    const vals = nums(c), va = c.value_axis || {};
    let axis_min = va.min != null ? va.min : null;
    if (axis_min == null && vals.length && CT.contract.axis_zero_based_kinds.includes(kind) && Math.min(...vals) >= 0) axis_min = 0;
    const hl = CT.contract.highlight;
    const highlight = c.highlight_category != null && kind === hl.kind && nSeries === 1 ? c.highlight_category : null;
    const mode = o.mode || 'light';
    return { kind, legend, data_labels, items, palette, color_map, unit, axis_min, axis_max: va.max != null ? va.max : null, highlight, dim_color: CT.deemphasis[mode], mode };
  }

  const num = (v, u) => Number(v).toLocaleString('en-US', { maximumFractionDigits: 2 }) + (u || '');
  function table(host, d, ink) {
    host.innerHTML = `<table style="width:100%;border-collapse:collapse;font:inherit;color:inherit"><thead><tr>${d.headers.map((h) => `<th style="text-align:left;padding:6px 10px;border-bottom:2px solid ${ink}">${esc(h)}</th>`).join('')}</tr></thead><tbody>${d.rows.map((r) => `<tr>${r.map((v) => `<td style="padding:6px 10px;border-bottom:1px solid rgba(128,128,128,.4)">${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  }
  function pie(host, c, R, ink) {
    const cats = c.data.categories || [], vals = (c.data.series || [])[0]?.values || [];
    if (!cats.length || vals.length !== cats.length || vals.some((v) => !Number.isFinite(v) || v < 0)) return fail(host, 'bad series');
    const W = 640, H = 360, cx = 180, cy = 190, r = 140, ri = c.kind === 'donut' ? r * .58 : 0, tot = vals.reduce((a, b) => a + b, 0) || 1;
    let a0 = -Math.PI / 2, g = '';
    vals.forEach((v, i) => {
      const a1 = a0 + (v / tot) * Math.PI * 2, big = a1 - a0 > Math.PI ? 1 : 0, p = (a, rr) => `${cx + rr * Math.cos(a)},${cy + rr * Math.sin(a)}`;
      const dpath = vals.length === 1 ? `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0` : ri ? `M${p(a0, r)}A${r},${r} 0 ${big} 1 ${p(a1, r)}L${p(a1, ri)}A${ri},${ri} 0 ${big} 0 ${p(a0, ri)}Z` : `M${cx},${cy}L${p(a0, r)}A${r},${r} 0 ${big} 1 ${p(a1, r)}Z`;
      g += `<path d="${dpath}" fill="${R.color_map[R.items[i]]}"/>`; a0 = a1;
    });
    // 与 PPTX 同口径：饼 = 扇区外标“类别 读数”，环 = 环带内只标读数；图例（色块 + 类别名）只在 legend 为真时出现
    const read = (i) => (c.stacked === 'percent' ? vals[i] + '%' : num(vals[i], R.unit));
    let lab = '', acc = -Math.PI / 2;
    vals.forEach((v, i) => { const am = acc + (v / tot) * Math.PI, rr = ri ? (r + ri) / 2 : r + 26;
      if (R.data_labels && v > 0) lab += `<text x="${cx + rr * Math.cos(am)}" y="${cy + rr * Math.sin(am) + 5}" text-anchor="${ri || Math.abs(Math.cos(am)) < .3 ? 'middle' : Math.cos(am) > 0 ? 'start' : 'end'}" font-size="16" fill="${ri ? '#fff' : ink}">${ri ? '' : esc(cats[i]) + ' '}${read(i)}</text>`;
      acc += (v / tot) * Math.PI * 2; });
    const leg = R.legend ? cats.map((t, i) => `<rect x="440" y="${70 + i * 34}" width="14" height="14" fill="${R.color_map[R.items[i]]}"/><text x="462" y="${82 + i * 34}" font-size="18" fill="${ink}">${esc(t)}</text>`).join('') : '';
    host.innerHTML = `<svg xmlns="${NS}" viewBox="0 0 ${W} ${H}" width="100%" height="100%" role="img" aria-label="${esc(c.title || '图表')}" style="font-family:inherit">${g}${lab}${leg}</svg>`;
  }
  function draw(host) {
    let c; try { c = JSON.parse(host.dataset.chart); } catch (e) { return fail(host, 'invalid JSON'); }
    if (!CT) return fail(host, 'contract.js missing');
    const ink = host.dataset.ink || 'currentColor', R = resolve(c, { mode: surfaceMode(host) }), d = c.data || {};
    if (c.kind === 'table') return table(host, d, ink);
    if (c.kind === 'pie' || c.kind === 'donut') return pie(host, c, R, ink);
    if (!['bar', 'line'].includes(c.kind)) return;
    const cats = d.categories || [], ser = d.series || [];
    if (!cats.length || !ser.length || ser.some((s) => s.values.length !== cats.length || s.values.some((v) => !Number.isFinite(v)))) return fail(host, 'bad series');
    const W = 640, H = 360, L = c.horizontal ? 110 : 48, B = 40, T = 28, Rr = 16, pw = W - L - Rr, ph = H - T - B, n = cats.length, k = ser.length, u = R.unit, labels = R.data_labels;
    const all = nums(c), mn = all.length ? Math.min(...all) : 0, lo = R.axis_min != null ? R.axis_min : (mn < 0 ? niceFloor(mn * 1.25) : 0);
    const hi = R.axis_max != null ? R.axis_max : (Math.max(...all, 0) * 1.1 || 1), span = hi - lo || 1;
    const base = Math.min(Math.max(0, lo), hi);          // 柱的基线：0（若在轴内）
    const vy = (v) => T + ph - ((v - lo) / span) * ph, vx = (v) => L + ((v - lo) / span) * pw;
    const col = (j, i) => R.highlight != null && R.highlight !== i ? R.dim_color : R.color_map[R.items[j]];
    let g = '';
    if (c.kind === 'line') {
      const x = (i) => L + (n === 1 ? pw / 2 : (pw * i) / (n - 1));
      g += `<line x1="${L}" y1="${vy(base)}" x2="${L + pw}" y2="${vy(base)}" stroke="${ink}" stroke-opacity=".5"/>`;
      ser.forEach((s, j) => { const cl = R.color_map[R.items[j]]; g += `<polyline fill="none" stroke="${cl}" stroke-width="3" points="${s.values.map((v, i) => `${x(i)},${vy(v)}`).join(' ')}"/>` + s.values.map((v, i) => `<circle cx="${x(i)}" cy="${vy(v)}" r="4" fill="${cl}"/>${labels ? `<text x="${x(i)}" y="${vy(v) - 10}" text-anchor="middle" font-size="16" fill="${ink}">${num(v, u)}</text>` : ''}`).join(''); });
      cats.forEach((t, i) => { g += `<text x="${x(i)}" y="${H - 14}" text-anchor="middle" font-size="16" fill="${ink}">${esc(t)}</text>`; });
    } else if (c.horizontal) {
      const bh = ph / n, w = (bh * 0.7) / k;
      g += `<line x1="${vx(base)}" y1="${T}" x2="${vx(base)}" y2="${T + ph}" stroke="${ink}" stroke-opacity=".5"/>`;
      cats.forEach((t, i) => { g += `<text x="${L - 8}" y="${T + bh * i + bh / 2 + 4}" text-anchor="end" font-size="16" fill="${ink}">${esc(t)}</text>`; ser.forEach((s, j) => { const v = s.values[i], x0 = Math.min(vx(v), vx(base)), len = Math.abs(vx(v) - vx(base)), yy = T + bh * i + bh * 0.15 + w * j; g += `<rect x="${x0}" y="${yy}" width="${len}" height="${w}" fill="${col(j, i)}"/>${labels ? `<text x="${x0 + len + 6}" y="${yy + w / 2 + 4}" font-size="16" fill="${ink}">${num(v, u)}</text>` : ''}`; }); });
    } else {
      const bw = pw / n, w = (bw * 0.7) / k;
      g += `<line x1="${L}" y1="${vy(base)}" x2="${L + pw}" y2="${vy(base)}" stroke="${ink}" stroke-opacity=".5"/>`;
      cats.forEach((t, i) => { g += `<text x="${L + bw * i + bw / 2}" y="${H - 14}" text-anchor="middle" font-size="16" fill="${ink}">${esc(t)}</text>`; ser.forEach((s, j) => { const v = s.values[i], y0 = Math.min(vy(v), vy(base)), h = Math.abs(vy(v) - vy(base)), xx = L + bw * i + bw * 0.15 + w * j; g += `<rect x="${xx}" y="${y0}" width="${w}" height="${h}" fill="${col(j, i)}"/>${labels ? `<text x="${xx + w / 2}" y="${y0 - 6}" text-anchor="middle" font-size="16" fill="${ink}">${num(v, u)}</text>` : ''}`; }); });
    }
    const legend = R.legend ? `<g font-size="16" fill="${ink}">${ser.map((s, j) => `<rect x="${L + j * 130}" y="4" width="12" height="12" fill="${R.color_map[R.items[j]]}"/><text x="${L + j * 130 + 18}" y="15">${esc(s.name)}</text>`).join('')}</g>` : '';
    host.innerHTML = `<svg xmlns="${NS}" viewBox="0 0 ${W} ${H}" width="100%" height="100%" role="img" aria-label="${esc(c.title || '图表')}" style="font-family:inherit">${legend}${g}</svg>`;
  }
  window.DeckChart = { resolve, draw, mode: surfaceMode, contract: CT };
  const run = () => document.querySelectorAll('[data-chart]').forEach(draw);
  document.readyState === 'loading' ? document.addEventListener('DOMContentLoaded', run) : run();
})();
