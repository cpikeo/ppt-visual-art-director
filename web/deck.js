/* PPT Visual Art Director · Web Presentation runtime v23.0.0
   Contract: <section class="slide" data-role data-transition data-notes> … elements with data-m (motion) / data-step (build).
   Content is visible by default (no JS → slides stack as a readable page); motion is layered on by WAAPI, so a JS failure never hides content.
   State separation: EXPORT (?export) is URL-only and static; IMMERSIVE follows the real fullscreenchange. They never read each other. */
(() => {
'use strict';
const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
const Q = new URLSearchParams(location.search);
const EXPORT = Q.has('export');
const mqRM = matchMedia('(prefers-reduced-motion: reduce)'), STILL = Q.has('still') || EXPORT;
let RM = mqRM.matches || STILL;
mqRM.addEventListener?.('change', () => { RM = mqRM.matches || STILL; if (RM && slides[cur]) { trans.forEach(a => { try { a.finish(); } catch (_) {} }); settle(slides[cur]); } });
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
let busyUntil = 0, slides = [], cur = -1, step = 0, maxStep = [], trans = [], digits = '', digitTimer = 0, errors = 0, startedAt = 0;
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
  trans.forEach(a => { try { a.finish(); } catch (_) {} }); trans = []; busyUntil = 0;
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
  trans = A; busyUntil = performance.now() + D;
  Promise.allSettled(A.map(x => x.finished)).then(() => {
    if (trans === A) { busyUntil = 0; A.forEach(x => { try { x.cancel(); } catch (_) {} }); from.classList.remove('is-leaving'); trans = []; }
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
  if (!EXPORT) try { history.replaceState(null, '', '#' + (i + 1)); localStorage.setItem('dk:' + location.pathname, i); } catch (_) {}
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
function resize() {
  const imm = isImmersive();
  document.documentElement.style.setProperty('--dk-bar', imm || EXPORT ? '0px' : bar.offsetHeight + 'px');
  if (EXPORT) { document.documentElement.style.setProperty('--dk-s', 1); return; }
  const vp = $('.dk-viewport'), w = vp.clientWidth || innerWidth, h = vp.clientHeight || innerHeight;
  document.documentElement.style.setProperty('--dk-s', Math.max(.1, Math.min(w / 1280, h / 720)));
}
function fail(e) {
  errors++; console.warn('[deck] motion error, falling back to final state', e);
  document.documentElement.classList.add('dk-safe');
  if (slides[cur]) settle(slides[cur]);
}
/* ---------- immersive (fullscreen) · chrome overlay ---------- */
const root = document.documentElement;
const fsEl = () => document.fullscreenElement || document.webkitFullscreenElement || null;
function isImmersive() { return !EXPORT && !!fsEl(); }
const isOpen = (el) => !!el && el.classList.contains('is-open');
let idleT = 0, hotZone = false, pointerDown = false, lastX = -1, lastY = -1, toastEl, toastT = 0, dlMenu, wheel = { acc: 0, last: 0, dir: 0, lastAbs: 0, until: 0 };
const HOT_PX = 72, IDLE_MS = 2000, INTERACTIVE = 'a,button,input,textarea,select,summary,[contenteditable],[data-wheel=native],.dk-bar,.dk-panel,.dk-help,.dk-overview,.dk-dl';
function toast(msg) { toastEl.textContent = msg; toastEl.classList.add('is-open'); clearTimeout(toastT); toastT = setTimeout(() => toastEl.classList.remove('is-open'), 2800); live.textContent = msg; }
async function fullscreen() {
  if (EXPORT) return;
  const de = document.documentElement;
  try {
    if (fsEl()) await (document.exitFullscreen || document.webkitExitFullscreen).call(document);
    else { const rq = de.requestFullscreen || de.webkitRequestFullscreen; if (!rq) return toast('此浏览器不支持全屏'); await rq.call(de, { navigationUI: 'hide' }); }
  } catch (e) { toast('全屏被浏览器拒绝：' + ((e && e.message) || '需要一次用户操作')); }
}
function chromeHeld() {   // 任何一项成立都不得隐藏：面板打开、悬停/键盘焦点在工具栏、按下中、指针在唤出区
  return isOpen(overview) || isOpen(panel) || isOpen(help) || isOpen(dlMenu) || bar.matches(':hover') || !!bar.querySelector(':focus-visible') || pointerDown || hotZone;
}
function showChrome() { for (const [el, c] of [[bar, 'is-idle'], [progress, 'is-idle'], [root, 'dk-idle']]) if (el.classList.contains(c)) el.classList.remove(c); }
function hideChrome(force) {
  if (!isImmersive() || (!force && chromeHeld())) return;
  if (force && (isOpen(overview) || isOpen(panel) || isOpen(help) || isOpen(dlMenu))) return;
  bar.classList.add('is-idle'); progress.classList.add('is-idle'); root.classList.add('dk-idle');
}
function wake() {
  showChrome(); clearTimeout(idleT);
  if (isImmersive()) idleT = setTimeout(function tick() { if (!isImmersive()) return; if (chromeHeld()) { idleT = setTimeout(tick, 400); return; } hideChrome(); }, IDLE_MS);
}
function onFullscreenChange() {
  const imm = isImmersive();
  root.classList.toggle('dk-fs', imm);
  clearTimeout(idleT); hotZone = false; pointerDown = false;
  if (imm) { if (bar.contains(document.activeElement)) document.activeElement.blur(); hideChrome(true); }   // 进入全屏：默认隐藏，不等空闲
  else showChrome();                                                                                         // Esc / 退出：立即恢复工具栏与指针
  resize(); requestAnimationFrame(resize);
  document.dispatchEvent(new CustomEvent('deck:fullscreen', { detail: { immersive: imm } }));
}
function onPointerMove(e) {
  if (EXPORT || !isImmersive()) return;
  if (Math.hypot(e.clientX - lastX, e.clientY - lastY) < 2 && lastX >= 0) return;   // 合成/抖动事件不唤出
  lastX = e.clientX; lastY = e.clientY;
  hotZone = e.clientY > innerHeight - HOT_PX; wake();
}

/* ---------- unified navigation: keyboard / buttons / wheel / touch all land here ---------- */
function nav(dir, src = 'api') {
  if (EXPORT || !slides.length) return false;
  const gesture = src === 'wheel' || src === 'touch';
  if (gesture && (performance.now() < busyUntil || isOpen(overview) || isOpen(black))) return false;   // 转场中手势不叠加
  const before = cur * 1e4 + step;
  dir > 0 ? next() : prev();
  return cur * 1e4 + step !== before;
}
function scrollable(el, dx, dy) {
  for (let e = el; e && e !== document.body && e !== root; e = e.parentElement) {
    const c = getComputedStyle(e);
    if (dy && /(auto|scroll)/.test(c.overflowY) && e.scrollHeight > e.clientHeight + 1 && (dy > 0 ? e.scrollTop + e.clientHeight < e.scrollHeight - 1 : e.scrollTop > 0)) return true;
    if (dx && /(auto|scroll)/.test(c.overflowX) && e.scrollWidth > e.clientWidth + 1 && (dx > 0 ? e.scrollLeft + e.clientWidth < e.scrollWidth - 1 : e.scrollLeft > 0)) return true;
  }
  return false;
}
/* 滚轮：一次手势只翻一页。方向取主轴；累计达阈值触发；触发后上锁，直到“静默间隙”或“明显的新手势”（反向，或增速）才解锁；
   鼠标滚轮单格（|Δ|≥100）立即触发；触控板惯性（持续衰减的小 Δ）不会再触发。 */
const WHEEL = { threshold: 80, gap: 160, cooldown: 520, strong: 40 };
function onWheel(e) {
  if (EXPORT || e.ctrlKey || e.metaKey || e.altKey) return;                 // 缩放（含触控板捏合）不是翻页
  if (e.target.closest?.('input,textarea,select,[contenteditable],[data-wheel=native],.dk-panel,.dk-overview,.dk-help,.dk-dl')) return;
  const k = e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? innerHeight : 1, dx = e.deltaX * k, dy = e.deltaY * k;
  const d = Math.abs(dy) >= Math.abs(dx) ? dy : dx;
  if (!d) return;
  if (scrollable(e.target, dx, dy)) return;                                  // 内部可滚控件先消费
  e.preventDefault();
  const t = e.timeStamp || performance.now(), dir = Math.sign(d), a = Math.abs(d), W = wheel;
  const quiet = t - W.last > WHEEL.gap; W.last = t;
  if (quiet) { W.acc = 0; W.locked = false; W.lastAbs = 0; }
  if (W.locked || t < W.until) {
    const fresh = t >= W.until && (dir !== W.dir ? a >= WHEEL.strong : a > W.lastAbs * 1.8 && a >= WHEEL.strong);   // 冷却结束后的反向/增速 = 新手势
    W.lastAbs = a;
    if (!fresh) return;
    W.locked = false; W.acc = 0;
  }
  if (dir !== W.dir) { W.acc = 0; W.dir = dir; }
  W.acc += d; W.lastAbs = a;
  if (Math.abs(W.acc) >= WHEEL.threshold) {
    const moved = nav(dir, 'wheel');
    W.acc = 0; W.locked = true; W.until = t + (moved ? WHEEL.cooldown : 120);     // 到边界不翻页，也要上锁，避免惯性反复尝试
  }
}
/* 触控：只认单指、水平为主的快速滑动；起点在控件上、多指、取消都放弃。鼠标拖动不翻页。 */
function bindTouch() {
  let tp = null;
  stage.addEventListener('pointerdown', (e) => {
    if (e.pointerType === 'mouse' || !e.isPrimary || e.target.closest(INTERACTIVE) || scrollable(e.target, 1, 0)) { tp = null; return; }
    tp = { id: e.pointerId, x: e.clientX, y: e.clientY, t: e.timeStamp };
  });
  stage.addEventListener('pointerup', (e) => {
    if (!tp || e.pointerId !== tp.id) return;
    const dx = e.clientX - tp.x, dy = e.clientY - tp.y, dt = e.timeStamp - tp.t; tp = null;
    if (dt < 900 && Math.abs(dx) > 48 && Math.abs(dx) > Math.abs(dy) * 1.5) nav(dx < 0 ? 1 : -1, 'touch');
  });
  ['pointercancel', 'lostpointercapture'].forEach((ev) => stage.addEventListener(ev, () => { tp = null; }));
}

function key(e) {
  if (e.metaKey || e.ctrlKey || e.altKey || /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName) || e.target.isContentEditable) return;
  const k = e.key;
  if (overview.classList.contains('is-open') && !['Escape', 'o', 'O', 'Enter', ' '].includes(k)) return;
  if (black.classList.contains('is-open')) { toggle(black, false); return; }
  if (/^[0-9]$/.test(k)) { digits += k; clearTimeout(digitTimer); digitTimer = setTimeout(() => { go(+digits - 1); digits = ''; }, 900); return; }
  const N = () => nav(1, 'key'), P = () => nav(-1, 'key'), togglePanel = () => { toggle(panel); startedAt = startedAt || Date.now(); panelUpdate(); };
  const NAVKEYS = ['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'];
  if (e.repeat && NAVKEYS.includes(k) && performance.now() < busyUntil) { e.preventDefault(); return; }   // 按住不放：转场期间不重复翻
  const m = { ArrowRight: N, ArrowDown: N, PageDown: N, ' ': N, Enter: () => { if (digits) { clearTimeout(digitTimer); go(+digits - 1); digits = ''; } else N(); },
    ArrowLeft: P, ArrowUp: P, PageUp: P, Backspace: P,
    Home: () => go(0, { dir: -1 }), End: () => go(slides.length - 1, { dir: 1 }),
    f: fullscreen, F: fullscreen, o: () => openOverview(), O: () => openOverview(), p: togglePanel, P: togglePanel,
    t: () => { startedAt = Date.now(); }, b: () => toggle(black), B: () => toggle(black), '?': () => toggle(help), h: () => toggle(help),
    Escape: () => { toggle(overview, false); toggle(help, false); toggle(black, false); toggle(dlMenu, false); } }[k];
  if (m) { e.preventDefault(); m(); }
}

/* 静态下载：只在页内带有发布清单（status=PASS）、且文件真实存在（HTTP HEAD）时出现；file:// 无法证明存在，不显示。 */
async function downloads() {
  let rel; try { rel = JSON.parse(document.getElementById('deck-release')?.textContent || 'null'); } catch (_) { return; }
  if (EXPORT || !rel || rel.gate !== 'PASS' || !Array.isArray(rel.files) || !/^https?:$/.test(location.protocol)) return;
  const ok = [];
  for (const f of rel.files) {
    try {
      const r = await fetch(f.path, { method: 'HEAD', cache: 'no-store' }), len = r.headers.get('content-length');
      if (r.ok && (!len || !f.bytes || +len === f.bytes || r.headers.get('content-encoding'))) ok.push(f);
    } catch (_) { /* 不存在就不展示 */ }
  }
  if (!ok.length) return;
  const mb = (n) => (n / 1048576).toFixed(n > 1048576 ? 1 : 2) + ' MB';
  dlMenu.innerHTML = '<h2>下载（构建期生成 · 已校验）</h2>' + ok.map((f) => `<a href="${f.path}" download="${f.name}"><b>${f.label || f.name}</b><span>${f.name}${f.bytes ? ' · ' + mb(f.bytes) : ''}</span><code title="SHA-256 ${f.sha256 || ''}">sha256 ${(f.sha256 || '').slice(0, 12)}…</code></a>`).join('');
  const b = $('[data-a=dl]', bar); b.hidden = false;
}

function build() {
  const root0 = $('.deck') || document.body;
  slides = $$('.slide', root0);
  if (!slides.length) return;
  document.documentElement.classList.add('dk-js'); if (EXPORT) document.documentElement.classList.add('dk-export');
  const mk = (tag, cls, parent = document.body, html = '') => { const e = document.createElement(tag); e.className = cls; e.innerHTML = html; parent.append(e); return e; };
  const vp = mk('div', 'dk-viewport'); fit = mk('div', 'dk-fit', vp); stage = mk('div', 'dk-stage', fit);
  stage.setAttribute('role', 'region'); stage.setAttribute('aria-roledescription', 'presentation'); stage.setAttribute('aria-label', document.title);
  slides.forEach((s, i) => { stage.append(s); s.setAttribute('role', 'group'); s.setAttribute('aria-roledescription', 'slide'); s.setAttribute('aria-label', `${i + 1} / ${slides.length}`); maxStep[i] = Math.max(0, ...$$('[data-step]', s).map(e => +e.dataset.step || 0)); });
  if (root0 !== document.body) root0.remove();
  progress = mk('div', 'dk-progress', document.body, '<i></i>');
  bar = mk('div', 'dk-bar', document.body, `<button data-a="prev" aria-label="上一步（←）">◀</button><div class="dk-dots" role="tablist" aria-label="页面导航"></div><span class="dk-count" aria-hidden="true"></span><button data-a="next" aria-label="下一步（→）">▶</button><button data-a="ov" aria-label="总览（O）">▦</button><button data-a="pn" aria-label="演讲者面板（P）">✎</button><button data-a="dl" aria-label="下载 PDF / PPTX" hidden>⇩</button><button data-a="fs" aria-label="全屏（F）">⛶</button><button data-a="help" aria-label="快捷键（?）">?</button>`);
  dots = $('.dk-dots', bar); count = $('.dk-count', bar);
  slides.forEach((s, k) => { const b = document.createElement('button'); b.setAttribute('aria-label', `跳到第 ${k + 1} 页：${titleOf(s)}`); b.onclick = () => go(k, { dir: k >= cur ? 1 : -1 }); dots.append(b); });
  overview = mk('div', 'dk-overview'); overview.setAttribute('role', 'dialog'); overview.setAttribute('aria-label', '页面总览');
  panel = mk('aside', 'dk-panel', document.body, '<h2>演讲者面板</h2><div class="dk-timer">00:00</div><h2>备注</h2><div class="dk-notes"></div><h2>下一页</h2><div class="dk-next"></div><div class="dk-step" style="opacity:.7"></div>');
  help = mk('div', 'dk-help', document.body, '<kbd>→</kbd> <kbd>空格</kbd> 下一步 · <kbd>←</kbd> 上一步 · 滚轮 / 滑动同样翻页<br><kbd>数字+Enter</kbd> 跳页 · <kbd>Home</kbd>/<kbd>End</kbd><br><kbd>O</kbd> 总览 · <kbd>P</kbd> 演讲者 · <kbd>T</kbd> 计时清零<br><kbd>F</kbd> 全屏（鼠标移动或移到底部唤出工具栏）· <kbd>B</kbd> 黑屏 · <kbd>Esc</kbd> 关闭');
  dlMenu = mk('div', 'dk-dl'); toastEl = mk('div', 'dk-toast'); toastEl.setAttribute('role', 'status');
  black = mk('div', 'dk-black'); live = mk('div', 'dk-live'); live.setAttribute('aria-live', 'polite');
  const togglePanel = () => { toggle(panel); startedAt = startedAt || Date.now(); panelUpdate(); };
  bar.onclick = (e) => { const a = e.target.closest('button')?.dataset.a; ({ prev: () => nav(-1, 'button'), next: () => nav(1, 'button'), ov: () => openOverview(), pn: togglePanel, fs: fullscreen, help: () => toggle(help), dl: () => toggle(dlMenu) })[a]?.(); };
  document.addEventListener('keydown', key);
  if (!EXPORT) {
    new ResizeObserver(resize).observe(vp);                    // 视口真实尺寸变化（含进入/退出全屏、旋转、地址栏）→ 缩放，不依赖事件先后
    addEventListener('resize', resize); addEventListener('orientationchange', () => { resize(); setTimeout(resize, 250); });
    document.addEventListener('fullscreenchange', onFullscreenChange); document.addEventListener('webkitfullscreenchange', onFullscreenChange);
    document.addEventListener('fullscreenerror', () => toast('全屏被浏览器拒绝'));
    addEventListener('wheel', onWheel, { passive: false });
    bindTouch();
    addEventListener('pointermove', onPointerMove, { passive: true });
    addEventListener('pointerdown', (e) => { pointerDown = true; if (isImmersive()) { hotZone = e.clientY > innerHeight - HOT_PX; wake(); } }, { passive: true });
    ['pointerup', 'pointercancel'].forEach((ev) => addEventListener(ev, () => { pointerDown = false; if (isImmersive()) wake(); }, { passive: true }));
    document.addEventListener('pointerleave', () => { hotZone = false; });
    document.addEventListener('keydown', (e) => { if (e.key === 'Tab' && isImmersive()) wake(); });          // 键盘进入工具栏：显示
    document.addEventListener('focusin', (e) => { if (isImmersive() && bar.contains(e.target) && e.target.matches(':focus-visible')) wake(); });
    addEventListener('blur', () => { clearTimeout(idleT); pointerDown = false; hotZone = false; });             // 失焦：不计时、不卡住按下状态
    addEventListener('focus', () => { resize(); if (isImmersive()) wake(); });
    const mo = new MutationObserver(() => { if (isImmersive() && (isOpen(overview) || isOpen(panel) || isOpen(help) || isOpen(dlMenu))) wake(); });   // 只观察面板自身：wake() 改的是工具栏/根节点，不会自触发
    [overview, panel, help, dlMenu].forEach((el) => mo.observe(el, { attributes: true, attributeFilter: ['class'] }));
  }
  addEventListener('error', fail);
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) { clearTimeout(idleT); wheel.acc = 0; return; }
    resize(); trans.forEach(a => { try { a.finish(); } catch (_) {} }); if (isImmersive()) wake();              // 标签页切回：恢复布局，转场收尾，工具栏短暂出现后再淡出
  });
  setInterval(clock, 1000);
  resize();
  const h = parseInt(location.hash.slice(1), 10), saved = +((() => { try { return EXPORT ? 0 : localStorage.getItem('dk:' + location.pathname); } catch (_) { return 0; } })() || 0);
  const start = EXPORT ? 0 : (h ? h - 1 : saved);
  (document.fonts?.ready || Promise.resolve()).then(() => { resize(); go(start, { instant: true, force: true }); document.documentElement.dataset.dkReady = '1'; downloads(); });
  window.Deck = {
    get count() { return slides.length; }, get index() { return cur; }, get errors() { return errors; }, get maxStep() { return maxStep.slice(); }, get immersive() { return isImmersive(); },
    get chromeHidden() { return bar.classList.contains('is-idle'); }, get busy() { return performance.now() < busyUntil; }, get step() { return step; },
    go, next: () => nav(1, 'api'), prev: () => nav(-1, 'api'), nav, fullscreen,
    /* export/QA: show slide i instantly in its final state (all builds visible, no motion) */
    show(i) { go(i, { instant: true, final: true, force: true }); slides.forEach((s, k) => { if (k !== i) s.classList.remove('is-active', 'is-leaving'); }); settle(slides[i]); return i; },
  };
}
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', build); else build();
})();
