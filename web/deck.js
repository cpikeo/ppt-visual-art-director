/* PPT Visual Art Director · Web Presentation runtime v22.0.0
   Contract: <section class="slide" data-role data-transition data-notes> … elements with data-m (motion) / data-step (build).
   Content is visible by default; motion is layered on by WAAPI, so a JS failure never hides content. */
(() => {
'use strict';
const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
const Q = new URLSearchParams(location.search);
const EXPORT = Q.has('export');
const RM = matchMedia('(prefers-reduced-motion: reduce)').matches || Q.has('still') || EXPORT;
const E = { out: 'cubic-bezier(.16,1,.3,1)', io: 'cubic-bezier(.76,0,.24,1)', soft: 'cubic-bezier(.33,1,.68,1)', lin: 'linear' };
const CLIP = { l: 'inset(0 100% 0 0)', r: 'inset(0 0 0 100%)', u: 'inset(100% 0 0 0)', d: 'inset(0 0 100% 0)' };
const PRESET = {
  fade:  () => [{ opacity: 0 }, { opacity: 1 }],
  rise:  () => [{ opacity: 0, transform: 'translateY(36px)' }, { opacity: 1, transform: 'none' }],
  sink:  () => [{ opacity: 0, transform: 'translateY(-36px)' }, { opacity: 1, transform: 'none' }],
  slide: (el) => [{ opacity: 0, transform: `translateX(${el.dataset.dir === 'r' ? 56 : -56}px)` }, { opacity: 1, transform: 'none' }],
  scale: () => [{ opacity: 0, transform: 'scale(.9)' }, { opacity: 1, transform: 'none' }],
  wipe:  (el) => [{ clipPath: CLIP[el.dataset.dir || 'l'] }, { clipPath: 'inset(0 0 0 0)' }],
  mask:  () => [{ clipPath: 'inset(0 0 100% 0)', transform: 'translateY(40%)' }, { clipPath: 'inset(0 0 -10% 0)', transform: 'none' }],
  grow:  (el) => [{ transform: 'scaleY(0)' }, { transform: 'scaleY(1)' }].map(k => k),
  growx: () => [{ transform: 'scaleX(0)' }, { transform: 'scaleX(1)' }],
};
let slides = [], cur = -1, step = 0, maxStep = [], trans = [], digits = '', digitTimer = 0, errors = 0, startedAt = 0;
let stage, fit, bar, progress, overview, panel, help, black, live, count, dots;

const num = (v, d) => (v === undefined || v === '' || isNaN(+v) ? d : +v);

function animEl(el, slideEl, auto) {
  const m = el.dataset.m;
  if (!m || RM) return null;
  const delay = num(el.dataset.at, auto) * 1000, dur = num(el.dataset.dur, m === 'count' ? 1.1 : .72) * 1000;
  const easing = E[el.dataset.ease] || (m === 'count' || m === 'draw' ? E.soft : E.out);
  const opt = { duration: dur, delay, easing, fill: 'both' };
  try {
    if (m === 'count') return count1(el, opt);
    if (m === 'draw') {
      const L = el.getTotalLength ? el.getTotalLength() : 0;
      el.style.strokeDasharray = L;
      return el.animate([{ strokeDashoffset: L }, { strokeDashoffset: 0 }], opt);
    }
    if (m === 'drift') {
      const [tx, ty, sc] = (el.dataset.drift || '-24,0,1.05').split(',').map(Number);
      return el.animate([{ transform: 'none' }, { transform: `translate(${tx}px,${ty}px) scale(${sc || 1})` }],
        { duration: num(el.dataset.dur, 16) * 1000, delay, easing: E.lin, fill: 'forwards' });
    }
    if (m === 'grow' || m === 'growx') el.style.transformOrigin = el.dataset.origin || (m === 'grow' ? 'bottom' : 'left');
    return el.animate(PRESET[m](el), opt);
  } catch (e) { fail(e); return null; }
}
function count1(el, opt) {
  const final = el.dataset.final || (el.dataset.final = el.textContent);
  const rx = /-?\d[\d,]*\.?\d*/, mm = final.match(rx);
  if (!mm) return null;
  const to = parseFloat(mm[0].replace(/,/g, '')), dec = (mm[0].split('.')[1] || '').length;
  const a = el.animate([{ opacity: 1 }, { opacity: 1 }], opt);  // timing carrier
  const t0 = performance.now() + opt.delay;
  const tick = (now) => {
    const p = Math.min(1, Math.max(0, (now - t0) / opt.duration)), k = 1 - Math.pow(1 - p, 3);
    el.textContent = final.replace(rx, (to * k).toFixed(dec));
    if (p < 1 && a.playState !== 'idle') requestAnimationFrame(tick); else el.textContent = final;
  };
  requestAnimationFrame(tick);
  a.onfinish = () => { el.textContent = final; };
  return a;
}
function settle(s) {
  $$('[data-step]', s).forEach(e => e.classList.remove('dk-pending'));
  s.getAnimations({ subtree: true }).forEach(a => { try { a.cancel(); } catch (_) {} });
  $$('[data-m=count]', s).forEach(e => { if (e.dataset.final) e.textContent = e.dataset.final; });
}
function enterMotion(s) {
  const els = $$('[data-m]:not([data-step])', s);
  const stag = {};
  els.forEach((el, i) => {
    const grp = el.closest('[data-stagger]');
    let auto = .25 + i * .09;
    if (grp) { const g = grp.dataset.stagger; stag[g] = (stag[g] ?? -1) + 1; auto = num(grp.dataset.start, .3) + stag[g] * num(grp.dataset.stagger, .08); }
    animEl(el, s, num(el.dataset.at, auto));
  });
}
function reveal(s, n, anim = true) {
  $$(`[data-step="${n}"]`, s).forEach((el, i) => {
    el.classList.remove('dk-pending');
    if (anim) { if (!el.dataset.m) el.dataset.m = 'rise'; animEl(el, s, num(el.dataset.at, i * .08)); }
  });
}
function hideStep(s, n) {
  $$(`[data-step="${n}"]`, s).forEach(el => { el.getAnimations().forEach(a => a.cancel()); el.classList.add('dk-pending'); });
}
function prepSteps(s, upTo) {
  $$('[data-step]', s).forEach(el => el.classList.toggle('dk-pending', +el.dataset.step > upTo));
}

function transition(from, to, dir, kind, instant) {
  trans.forEach(a => { try { a.finish(); } catch (_) {} }); trans = [];
  slides.forEach(s => { if (s !== to && s !== from) s.classList.remove('is-active', 'is-leaving'); });
  to.classList.add('is-active');
  if (!from || from === to) return;
  if (instant || RM || kind === 'cut') { from.classList.remove('is-active', 'is-leaving'); return; }
  from.classList.remove('is-active'); from.classList.add('is-leaving');
  const D = 760, s = dir >= 0 ? 1 : -1, A = [];
  const a = (el, kf) => A.push(el.animate(kf, { duration: D, easing: E.io, fill: 'both' }));
  if (kind === 'fade') a(to, [{ opacity: 0 }, { opacity: 1 }]);
  else if (kind === 'wipe') a(to, [{ clipPath: s > 0 ? CLIP.r : CLIP.l }, { clipPath: 'inset(0 0 0 0)' }]);
  else if (kind === 'zoom') { a(to, [{ opacity: 0, transform: 'scale(1.08)' }, { opacity: 1, transform: 'none' }]); a(from, [{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'scale(.96)' }]); }
  else if (kind === 'lift') { a(to, [{ opacity: 0, transform: `translateY(${s * 7}%)` }, { opacity: 1, transform: 'none' }]); a(from, [{ opacity: 1 }, { opacity: 0 }]); }
  else { a(to, [{ transform: `translateX(${s * 100}%)` }, { transform: 'none' }]); a(from, [{ transform: 'none', opacity: 1 }, { transform: `translateX(${-s * 26}%)`, opacity: .35 }]); }
  trans = A;
  Promise.allSettled(A.map(x => x.finished)).then(() => {
    if (trans === A) { A.forEach(x => { try { x.cancel(); } catch (_) {} }); from.classList.remove('is-leaving'); trans = []; }
  });
}

function go(i, o = {}) {
  i = Math.max(0, Math.min(slides.length - 1, i | 0));
  if (i === cur && !o.force) return;
  const from = slides[cur], to = slides[i], dir = o.dir ?? (i >= cur ? 1 : -1);
  if (from) settle(from);
  settle(to);
  const back = dir < 0 && !o.instant && cur >= 0;
  step = back || o.final ? maxStep[i] : 0;
  prepSteps(to, step);
  const first = cur < 0;
  cur = i;
  transition(first ? null : from, to, dir, to.dataset.transition || document.documentElement.dataset.transition || 'push', o.instant || first && true);
  slides.forEach((s, k) => { s.inert = k !== i; s.setAttribute('aria-hidden', k !== i); });
  if (!o.final && !back) enterMotion(to);
  ui();
  try { history.replaceState(null, '', '#' + (i + 1)); localStorage.setItem('dk:' + location.pathname, i); } catch (_) {}
  document.dispatchEvent(new CustomEvent('deck:slide', { detail: { index: i, step } }));
}
function next() {
  if (step < maxStep[cur]) { step++; reveal(slides[cur], step); ui(); }
  else if (cur < slides.length - 1) go(cur + 1, { dir: 1 });
}
function prev() {
  if (step > 0) { hideStep(slides[cur], step); step--; ui(); }
  else if (cur > 0) go(cur - 1, { dir: -1 });
}

function ui() {
  const n = slides.length;
  progress.firstElementChild.style.width = ((cur + 1) / n * 100) + '%';
  count.textContent = `${cur + 1} / ${n}`;
  $$('button', dots).forEach((b, k) => b.setAttribute('aria-current', k === cur));
  const s = slides[cur], t = $('h1,h2,[data-title]', s);
  live.textContent = `${cur + 1} / ${n}${t ? ' · ' + t.textContent.trim().replace(/\s+/g, ' ') : ''}`;
  if (panel.classList.contains('is-open')) panelUpdate();
}
function notesOf(s) { const a = $('aside.notes', s); return (a ? a.textContent : s.dataset.notes || '').trim(); }
function titleOf(s) { const t = $('h1,h2,[data-title]', s); return t ? t.textContent.trim().replace(/\s+/g, ' ') : ''; }
function panelUpdate() {
  $('.dk-notes', panel).textContent = notesOf(slides[cur]) || '（本页无备注）';
  const nx = slides[cur + 1];
  $('.dk-next', panel).textContent = nx ? titleOf(nx) : '— 结束 —';
  $('.dk-step', panel).textContent = maxStep[cur] ? `构建 ${step} / ${maxStep[cur]}` : '无构建步骤';
}
function clock() {
  if (!panel.classList.contains('is-open')) return;
  const t = Math.floor((Date.now() - startedAt) / 1000);
  $('.dk-timer', panel).textContent = String(Math.floor(t / 60)).padStart(2, '0') + ':' + String(t % 60).padStart(2, '0');
}
function toggle(el, force) { el.classList.toggle('is-open', force); return el.classList.contains('is-open'); }
function openOverview(force) {
  const open = toggle(overview, force);
  if (!open) return;
  overview.textContent = '';
  slides.forEach((s, k) => {
    const b = document.createElement('button'); b.className = 'dk-thumb'; b.setAttribute('aria-label', `第 ${k + 1} 页 ${titleOf(s)}`);
    b.setAttribute('aria-current', k === cur);
    const w = document.createElement('span'), c = s.cloneNode(true);
    c.className = 'slide dk-clone'; c.removeAttribute('id'); $$('[id]', c).forEach(e => e.removeAttribute('id'));
    $$('.dk-pending', c).forEach(e => e.classList.remove('dk-pending'));
    c.inert = true; w.append(c); b.append(w);
    const em = document.createElement('em'); em.textContent = `${k + 1}  ${titleOf(s)}`; b.append(em);
    b.onclick = () => { openOverview(false); go(k, { dir: k >= cur ? 1 : -1 }); };
    overview.append(b);
  });
  const w = $('.dk-thumb span', overview).clientWidth;
  $$('.dk-clone', overview).forEach(c => c.style.transform = `scale(${w / 1280})`);
  $('.dk-thumb[aria-current=true]', overview)?.focus();
}
async function fullscreen() {
  try { document.fullscreenElement ? await document.exitFullscreen() : await document.documentElement.requestFullscreen(); } catch (_) {}
}
function resize() {
  const fs = !!document.fullscreenElement || EXPORT;
  document.documentElement.style.setProperty('--dk-bar', fs ? '0px' : bar.offsetHeight + 'px');
  const w = innerWidth, h = innerHeight - (fs ? 0 : bar.offsetHeight);
  document.documentElement.style.setProperty('--dk-s', EXPORT ? 1 : Math.max(.1, Math.min(w / 1280, h / 720)));
}
function fail(e) {
  errors++; console.warn('[deck] motion error, falling back to final state', e);
  document.documentElement.classList.add('dk-safe');
  if (slides[cur]) settle(slides[cur]);
}
function key(e) {
  if (e.metaKey || e.ctrlKey || e.altKey || /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)) return;
  const k = e.key;
  if (overview.classList.contains('is-open') && !['Escape', 'o', 'O', 'Enter', ' '].includes(k)) return;
  if (black.classList.contains('is-open')) { toggle(black, false); return; }
  if (/^[0-9]$/.test(k)) { digits += k; clearTimeout(digitTimer); digitTimer = setTimeout(() => { go(+digits - 1); digits = ''; }, 900); return; }
  const m = { ArrowRight: next, ArrowDown: next, PageDown: next, ' ': next, Enter: () => { if (digits) { clearTimeout(digitTimer); go(+digits - 1); digits = ''; } else next(); },
    ArrowLeft: prev, ArrowUp: prev, PageUp: prev, Backspace: prev,
    Home: () => go(0, { dir: -1 }), End: () => go(slides.length - 1, { dir: 1 }),
    f: fullscreen, F: fullscreen, o: () => openOverview(), O: () => openOverview(), p: () => { toggle(panel); startedAt = startedAt || Date.now(); panelUpdate(); }, P: () => { toggle(panel); startedAt = startedAt || Date.now(); panelUpdate(); },
    t: () => { startedAt = Date.now(); }, b: () => toggle(black), B: () => toggle(black), '?': () => toggle(help), h: () => toggle(help),
    Escape: () => { toggle(overview, false); toggle(help, false); toggle(black, false); } }[k];
  if (m) { e.preventDefault(); m(); }
}

function build() {
  const root = $('.deck') || document.body;
  slides = $$('.slide', root);
  if (!slides.length) return;
  document.documentElement.classList.add('dk-js'); if (EXPORT) document.documentElement.classList.add('dk-export');
  const mk = (tag, cls, parent = document.body, html = '') => { const e = document.createElement(tag); e.className = cls; e.innerHTML = html; parent.append(e); return e; };
  const vp = mk('div', 'dk-viewport'); fit = mk('div', 'dk-fit', vp); stage = mk('div', 'dk-stage', fit);
  stage.setAttribute('role', 'region'); stage.setAttribute('aria-roledescription', 'presentation'); stage.setAttribute('aria-label', document.title);
  slides.forEach((s, i) => { stage.append(s); s.setAttribute('role', 'group'); s.setAttribute('aria-roledescription', 'slide'); s.setAttribute('aria-label', `${i + 1} / ${slides.length}`); maxStep[i] = Math.max(0, ...$$('[data-step]', s).map(e => +e.dataset.step || 0)); });
  if (root !== document.body) root.remove();
  progress = mk('div', 'dk-progress', document.body, '<i></i>');
  bar = mk('div', 'dk-bar', document.body, `<button data-a="prev" aria-label="上一步（←）">◀</button><div class="dk-dots" role="tablist" aria-label="页面导航"></div><span class="dk-count" aria-hidden="true"></span><button data-a="next" aria-label="下一步（→）">▶</button><button data-a="ov" aria-label="总览（O）">▦</button><button data-a="pn" aria-label="演讲者面板（P）">✎</button><button data-a="fs" aria-label="全屏（F）">⛶</button><button data-a="help" aria-label="快捷键（?）">?</button>`);
  dots = $('.dk-dots', bar); count = $('.dk-count', bar);
  slides.forEach((s, k) => { const b = document.createElement('button'); b.setAttribute('aria-label', `跳到第 ${k + 1} 页：${titleOf(s)}`); b.onclick = () => go(k, { dir: k >= cur ? 1 : -1 }); dots.append(b); });
  overview = mk('div', 'dk-overview'); overview.setAttribute('role', 'dialog'); overview.setAttribute('aria-label', '页面总览');
  panel = mk('aside', 'dk-panel', document.body, '<h2>演讲者面板</h2><div class="dk-timer">00:00</div><h2>备注</h2><div class="dk-notes"></div><h2>下一页</h2><div class="dk-next"></div><div class="dk-step" style="opacity:.7"></div>');
  help = mk('div', 'dk-help', document.body, '<kbd>→</kbd> <kbd>空格</kbd> 下一步 · <kbd>←</kbd> 上一步<br><kbd>数字+Enter</kbd> 跳页 · <kbd>Home</kbd>/<kbd>End</kbd><br><kbd>O</kbd> 总览 · <kbd>P</kbd> 演讲者 · <kbd>T</kbd> 计时清零<br><kbd>F</kbd> 全屏 · <kbd>B</kbd> 黑屏 · <kbd>Esc</kbd> 关闭');
  black = mk('div', 'dk-black'); live = mk('div', 'dk-live'); live.setAttribute('aria-live', 'polite');
  bar.onclick = (e) => { const a = e.target.closest('button')?.dataset.a; ({ prev, next, ov: () => openOverview(), pn: () => { toggle(panel); startedAt = startedAt || Date.now(); panelUpdate(); }, fs: fullscreen, help: () => toggle(help) })[a]?.(); };
  document.addEventListener('keydown', key);
  addEventListener('resize', resize); document.addEventListener('fullscreenchange', resize);
  addEventListener('error', fail);
  document.addEventListener('visibilitychange', () => { if (!document.hidden && slides[cur]) { trans.forEach(a => { try { a.finish(); } catch (_) {} }); } });
  let tx = 0, ty = 0;
  stage.addEventListener('pointerdown', (e) => { tx = e.clientX; ty = e.clientY; });
  stage.addEventListener('pointerup', (e) => { const dx = e.clientX - tx, dy = e.clientY - ty; if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy) * 1.5) (dx < 0 ? next : prev)(); });
  let idle = 0; const wake = () => { bar.classList.remove('is-idle'); clearTimeout(idle); if (document.fullscreenElement) idle = setTimeout(() => bar.classList.add('is-idle'), 2600); };
  ['pointermove', 'keydown'].forEach(ev => addEventListener(ev, wake));
  setInterval(clock, 1000);
  resize();
  const h = parseInt(location.hash.slice(1), 10), saved = +((() => { try { return localStorage.getItem('dk:' + location.pathname); } catch (_) { return 0; } })() || 0);
  const start = EXPORT ? 0 : (h ? h - 1 : saved);
  (document.fonts?.ready || Promise.resolve()).then(() => { resize(); go(start, { instant: true, force: true }); document.documentElement.dataset.dkReady = '1'; });
  window.Deck = {
    get count() { return slides.length; }, get index() { return cur; }, get errors() { return errors; }, get maxStep() { return maxStep.slice(); },
    go, next, prev,
    /* export/QA: show slide i instantly in its final state (all builds visible, no motion) */
    show(i) { go(i, { instant: true, final: true, force: true }); slides.forEach((s, k) => { if (k !== i) s.classList.remove('is-active', 'is-leaving'); }); settle(slides[i]); return i; },
  };
}
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', build); else build();
})();
