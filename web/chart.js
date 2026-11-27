/* data-chart → HTML 视图。与 PPTX 导出读取的是同一份 JSON（单一数据源）：HTML 与 PPTX 的数据、系列色、图例必然一致。
   支持 bar（竖/横、多系列）、line、table；其余 kind 由 deck 自己画（宿主元素保留 data-chart，导出仍是原生图表）。
   数值轴恒从 0 起，不截轴；非有限数值直接标 data-chart-error，不画。须在 deck.js 之前加载。 */
(() => {
  let META = {}; try { META = JSON.parse(document.getElementById('deck-meta')?.textContent || '{}'); } catch (e) { /* 无 meta 用默认色 */ }
  const NS = 'http://www.w3.org/2000/svg', PAL = META.palette?.length ? META.palette : ['#E8734A', '#4A90C2', '#7BB661', '#C9A227', '#8E6BBF', '#5FB8B0'];
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
  const num = (v, u) => Number(v).toLocaleString('en-US', { maximumFractionDigits: 2 }) + (u || '');
  const fail = (host, why) => { host.dataset.chartError = why; };
  function draw(host) {
    let c; try { c = JSON.parse(host.dataset.chart); } catch (e) { return fail(host, 'invalid JSON'); }
    const d = c.data || {}, pal = c.palette || PAL, ink = host.dataset.ink || 'currentColor';
    if (c.kind === 'table') {
      host.innerHTML = `<table style="width:100%;border-collapse:collapse;font:inherit;color:inherit"><thead><tr>${d.headers.map((h) => `<th style="text-align:left;padding:6px 10px;border-bottom:2px solid ${ink}">${esc(h)}</th>`).join('')}</tr></thead><tbody>${d.rows.map((r) => `<tr>${r.map((v) => `<td style="padding:6px 10px;border-bottom:1px solid rgba(128,128,128,.4)">${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
      return;
    }
    if (!['bar', 'line'].includes(c.kind)) return;
    const cats = d.categories || [], ser = d.series || [];
    if (!cats.length || !ser.length || ser.some((s) => s.values.length !== cats.length || s.values.some((v) => !Number.isFinite(v)))) return fail(host, 'bad series');
    const W = 640, H = 360, L = c.horizontal ? 110 : 48, B = 40, T = 28, R = 16, max = Math.max(...ser.flatMap((s) => s.values), 0) * 1.1 || 1;
    const pw = W - L - R, ph = H - T - B, n = cats.length, k = ser.length, u = c.unit || '', labels = c.data_labels !== false;
    // 系列色 = palette[系列序]（与 PPTX 同序）；单系列设了 highlight_category 时，其余柱淡化（PPTX 同做法）
    const col = (j, i) => c.highlight_category != null && k === 1 && c.highlight_category !== i ? 'rgba(128,128,128,.45)' : pal[j % pal.length];
    let g = `<line x1="${L}" y1="${T + ph}" x2="${L + pw}" y2="${T + ph}" stroke="${ink}" stroke-opacity=".5"/>`;
    if (c.kind === 'line') {
      const x = (i) => L + (n === 1 ? pw / 2 : (pw * i) / (n - 1)), y = (v) => T + ph - (v / max) * ph;
      ser.forEach((s, j) => { g += `<polyline fill="none" stroke="${pal[j % pal.length]}" stroke-width="3" points="${s.values.map((v, i) => `${x(i)},${y(v)}`).join(' ')}"/>` + s.values.map((v, i) => `<circle cx="${x(i)}" cy="${y(v)}" r="4" fill="${pal[j % pal.length]}"/>${labels ? `<text x="${x(i)}" y="${y(v) - 10}" text-anchor="middle" font-size="16" fill="${ink}">${num(v, u)}</text>` : ''}`).join(''); });
      cats.forEach((t, i) => { g += `<text x="${x(i)}" y="${H - 14}" text-anchor="middle" font-size="16" fill="${ink}">${esc(t)}</text>`; });
    } else if (c.horizontal) {
      const bh = ph / n, w = (bh * 0.7) / k;
      cats.forEach((t, i) => { g += `<text x="${L - 8}" y="${T + bh * i + bh / 2 + 4}" text-anchor="end" font-size="16" fill="${ink}">${esc(t)}</text>`; ser.forEach((s, j) => { const v = s.values[i], len = (v / max) * pw, yy = T + bh * i + bh * 0.15 + w * j; g += `<rect x="${L}" y="${yy}" width="${len}" height="${w}" fill="${col(j, i)}"/>${labels ? `<text x="${L + len + 6}" y="${yy + w / 2 + 4}" font-size="16" fill="${ink}">${num(v, u)}</text>` : ''}`; }); });
    } else {
      const bw = pw / n, w = (bw * 0.7) / k;
      cats.forEach((t, i) => { g += `<text x="${L + bw * i + bw / 2}" y="${H - 14}" text-anchor="middle" font-size="16" fill="${ink}">${esc(t)}</text>`; ser.forEach((s, j) => { const v = s.values[i], h = (v / max) * ph, xx = L + bw * i + bw * 0.15 + w * j; g += `<rect x="${xx}" y="${T + ph - h}" width="${w}" height="${h}" fill="${col(j, i)}"/>${labels ? `<text x="${xx + w / 2}" y="${T + ph - h - 6}" text-anchor="middle" font-size="16" fill="${ink}">${num(v, u)}</text>` : ''}`; }); });
    }
    const legend = (c.legend ?? k > 1) ? `<g font-size="16" fill="${ink}">${ser.map((s, j) => `<rect x="${L + j * 130}" y="4" width="12" height="12" fill="${pal[j % pal.length]}"/><text x="${L + j * 130 + 18}" y="15">${esc(s.name)}</text>`).join('')}</g>` : '';
    host.innerHTML = `<svg xmlns="${NS}" viewBox="0 0 ${W} ${H}" width="100%" height="100%" role="img" aria-label="${esc(c.title || '图表')}" style="font-family:inherit">${legend}${g}</svg>`;
  }
  const run = () => document.querySelectorAll('[data-chart]').forEach(draw);
  document.readyState === 'loading' ? document.addEventListener('DOMContentLoaded', run) : run();
})();
