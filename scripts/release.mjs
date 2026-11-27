#!/usr/bin/env node
/* 构建期发行：deck.html → 同一批 PNG → PDF + 展示版 PPTX；同一 DOM → 可编辑 PPTX；核验；SHA-256；静态下载页。
   node scripts/release.mjs deck.html --out dist/NAME [--dpr 2|3] [--force] [--allow-unverified libreoffice] [--powerpoint-evidence FILE]
   无 API、无队列、无第二套导出器：只顺序调用 web_export.mjs 与 web_verify.py。
   覆盖保护：out 已存在默认拒绝；--force 时先在同级暂存目录完整构建，成功后才替换，失败不动旧产物并清理暂存。

   发行门禁（gate，唯一口径）：
     BLOCKED               任一核验 BLOCK、任一步骤失败、渲染不确定（两次独立渲染的 PNG 不同）、或存在未获允许的 UNVERIFIED。不生成可下载清单。
     PASS_WITH_LIMITATIONS 无 BLOCK；UNVERIFIED 只剩“允许的例外”，并逐条写入 release.json 的 limitations：
                             · powerpoint（永久例外）：没有 Microsoft PowerPoint 实机证据；
                             · libreoffice（仅当显式 --allow-unverified libreoffice）：本机无 LibreOffice，未做 PPTX 真实渲染对照。
     PASS                  无 BLOCK 且无任何 UNVERIFIED：必须给出与本次可编辑 PPTX SHA-256 相符的 --powerpoint-evidence，并有 LibreOffice 渲染对照。
   普通 UNVERIFIED 不会被汇总成 PASS。 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs'; import path from 'node:path'; import crypto from 'node:crypto'; import { fileURLToPath } from 'node:url';
const HERE = path.dirname(fileURLToPath(import.meta.url)), ROOT = path.join(HERE, '..');
const argv = process.argv.slice(2), deck = argv[0], opt = (k, d) => { const i = argv.indexOf(k); return i >= 0 ? argv[i + 1] : d; }, FORCE = argv.includes('--force');
const die = (m) => { console.error('ERROR ' + m); process.exit(2); };
if (!deck || !opt('--out')) die('usage: release.mjs deck.html --out dist/NAME [--dpr 2|3] [--force] [--allow-unverified libreoffice] [--powerpoint-evidence FILE]');
const deckPath = path.resolve(deck), out = path.resolve(opt('--out')), dpr = opt('--dpr', '2');
const ALLOW = new Set((opt('--allow-unverified', '') || '').split(',').filter(Boolean));
for (const a of ALLOW) if (a !== 'libreoffice') die(`--allow-unverified 只接受 libreoffice（powerpoint 永远是披露例外）: ${a}`);
if (!fs.existsSync(deckPath)) die('deck 不存在: ' + deckPath);
if (fs.existsSync(out) && !FORCE) die(`${out} 已存在，未覆盖；换目录或加 --force`);
const PY = process.env.PYTHON || 'python3', version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
const html = fs.readFileSync(deckPath, 'utf8'), meta = (() => { try { return JSON.parse(/<script[^>]*id="deck-meta"[^>]*>([\s\S]*?)<\/script>/.exec(html)?.[1] || '{}'); } catch { return {}; } })();
const slug = String(meta.deck_id || path.basename(path.dirname(deckPath))).toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5]+/g, '-').replace(/^-|-$/g, '') || 'deck';
fs.mkdirSync(path.dirname(out), { recursive: true });
const stage = fs.mkdtempSync(path.join(path.dirname(out), `.release-${slug}-`));
const sh = (label, cmd, args, env = {}) => { const r = spawnSync(cmd, args, { encoding: 'utf8', cwd: ROOT, env: { ...process.env, ...env } }); return { label, code: r.status, out: (r.stdout || '') + (r.stderr || '') }; };
const sha = (f) => crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex');
let ok = false;
try {
  const dl = path.join(stage, 'downloads'), work = path.join(stage, 'work'), site = path.join(stage, 'site'), rep = path.join(stage, 'reports');
  [dl, work, site, rep].forEach((d) => fs.mkdirSync(d, { recursive: true }));
  const E = (...a) => ['scripts/web_export.mjs', ...a], A = path.join(work, 'frames'), B = path.join(work, 'baseline');
  const F = { pdf: `${slug}.pdf`, image: `${slug}-image.pptx`, editable: `${slug}-editable.pptx` };
  const steps = [
    sh('frames', 'node', E('shots', deckPath, '--shots', A, '--dpr', dpr)),                      // 同一批 PNG：PDF 与展示版 PPTX 的唯一像素来源
    sh('baseline', 'node', E('shots', deckPath, '--shots', B, '--dpr', dpr)),                    // 另起浏览器上下文再渲染一次：独立基准，同时检验渲染确定性
    sh('pdf', 'node', E('pdf', deckPath, path.join(dl, F.pdf), '--from-shots', A)),
    sh('pptx-image', 'node', E('pptx-image', deckPath, path.join(dl, F.image), '--from-shots', A)),
    sh('pptx', 'node', E('pptx', deckPath, path.join(dl, F.editable), '--assets', path.join(work, 'assets')), { PYTHON: PY }),
  ];
  const bad = steps.find((s) => s.code !== 0);
  if (bad) throw new Error(`${bad.label} 失败，未发布任何文件:\n${bad.out.slice(-1500)}`);
  for (const f of fs.readdirSync(dl)) if (/\.(plan(\.manifest)?|pdf-report|image-pptx-report)\.json$/.test(f)) fs.renameSync(path.join(dl, f), path.join(/\.plan/.test(f) ? work : rep, f));
  const stem = (n) => n.replace(/\.(pdf|pptx)$/, '');
  const v = sh('verify', PY, ['scripts/web_verify.py', '--pptx', path.join(dl, F.editable), '--manifest', path.join(work, `${stem(F.editable)}.plan.manifest.json`), '--pdf', path.join(dl, F.pdf), '--pdf-report', path.join(rep, `${stem(F.pdf)}.pdf-report.json`),
    '--image-pptx', path.join(dl, F.image), '--image-pptx-report', path.join(rep, `${stem(F.image)}.image-pptx-report.json`), '--shots', B, '--render-pptx', '--out', path.join(work, 'verify'), '--json', path.join(work, 'verify.json')]);
  if (v.code === 2 || !fs.existsSync(path.join(work, 'verify.json'))) throw new Error('verify 输入错误:\n' + v.out.slice(-1500));
  const vr = JSON.parse(fs.readFileSync(path.join(work, 'verify.json'), 'utf8')), checks = [...vr.checks];
  // 渲染确定性：两次独立渲染的逐页 PNG 必须字节相同，否则“像素一致”无从谈起
  const ba = JSON.parse(fs.readFileSync(path.join(A, 'baseline.json'), 'utf8')), bb = JSON.parse(fs.readFileSync(path.join(B, 'baseline.json'), 'utf8'));
  const diff = ba.pages.filter((p, i) => p.sha256 !== bb.pages[i]?.sha256).map((p) => p.page);
  checks.push({ area: '渲染', name: 'render determinism: two independent renders yield byte-identical PNGs', status: diff.length || ba.pages.length !== bb.pages.length ? 'BLOCK' : 'PASS', detail: diff.length ? `pages differ: ${diff.join(',')}` : `${ba.pages.length} pages identical` });
  if (!checks.some((c) => /PowerPoint/.test(c.name))) checks.push({ area: 'PPTX', name: 'Microsoft PowerPoint real render', status: 'UNVERIFIED', detail: 'not run in this environment (no PowerPoint); LibreOffice is a different layout engine' });
  // PowerPoint 证据：只在与本次可编辑 PPTX 的 SHA-256 相符时才算数
  const editSha = sha(path.join(dl, F.editable)); let pptEvidence = null;
  if (opt('--powerpoint-evidence')) {
    try { const ev = JSON.parse(fs.readFileSync(path.resolve(opt('--powerpoint-evidence')), 'utf8')); if (ev.editable_pptx_sha256 === editSha && ev.application && ev.reviewer && ev.date) pptEvidence = ev; } catch { /* ignored: stays UNVERIFIED */ }
  }
  const kind = (c) => /PowerPoint/.test(c.name) ? 'powerpoint' : /LibreOffice|visual fidelity/.test(c.name) ? 'libreoffice' : 'other';
  for (const c of checks) if (c.status === 'UNVERIFIED' && kind(c) === 'powerpoint' && pptEvidence) { c.status = 'PASS'; c.detail = `evidence: ${pptEvidence.application}, ${pptEvidence.reviewer}, ${pptEvidence.date}`; }
  const blocks = checks.filter((c) => c.status === 'BLOCK'), unv = checks.filter((c) => c.status === 'UNVERIFIED');
  const disallowed = unv.filter((c) => { const k = kind(c); return !(k === 'powerpoint' || (k === 'libreoffice' && ALLOW.has('libreoffice'))); });
  const gate = blocks.length || disallowed.length ? 'BLOCKED' : unv.length ? 'PASS_WITH_LIMITATIONS' : 'PASS';
  const limitations = [
    ...unv.map((c) => ({ kind: 'UNVERIFIED', area: c.area, name: c.name, detail: c.detail, allowed: !disallowed.includes(c) })),
    ...(vr.downgrades || []).map((d) => ({ kind: 'DOWNGRADE', ...d })),
    { kind: 'SCOPE', detail: '像素一致只对应同一渲染环境（同一 Chromium、视口、DPR、字体、图片、动画终态）的基准；不同设备/播放器的显示差异未验证。' },
    { kind: 'SCOPE', detail: '可编辑 PPTX 与 HTML 的换行、字重、行高可能存在差异；不承诺跨渲染器 100% 一致。' },
  ];
  // 站点：deck 副本（runtime 路径改写到 ./web）+ 运行时 + assets；只在 gate≠BLOCKED 时注入下载清单
  fs.cpSync(path.join(ROOT, 'web'), path.join(site, 'web'), { recursive: true });
  const adir = path.join(path.dirname(deckPath), 'assets'); if (fs.existsSync(adir)) fs.cpSync(adir, path.join(site, 'assets'), { recursive: true });
  const labels = { [F.pdf]: 'PDF · 像素一致图像页（文字不可选）', [F.image]: 'PPTX · 展示版（整页图像，不可编辑）', [F.editable]: 'PPTX · 可编辑（原生对象，含披露的局部栅格）' };
  const files = [F.pdf, F.image, F.editable].map((f) => { const p = path.join(dl, f); return { name: f, path: `downloads/${f}`, bytes: fs.statSync(p).size, sha256: sha(p), kind: path.extname(f).slice(1), label: labels[f] }; });
  const pdfRep = JSON.parse(fs.readFileSync(path.join(rep, `${stem(F.pdf)}.pdf-report.json`), 'utf8'));
  const release = { name: slug, version, gate, verify_status: vr.status, powerpoint: pptEvidence ? 'PASS' : 'UNVERIFIED', pdf: { mode: pdfRep.mode, dpr: pdfRep.dpr, claim: pdfRep.claim }, stats: vr.stats, files, counts: { PASS: checks.filter((c) => c.status === 'PASS').length, BLOCK: blocks.length, UNVERIFIED: unv.length }, limitations, checks: checks.map(({ area, name, status, detail }) => ({ area, name, status, detail: String(detail).slice(0, 400) })) };
  fs.writeFileSync(path.join(stage, 'release.json'), JSON.stringify(release, null, 2));
  fs.writeFileSync(path.join(stage, 'verify.json'), JSON.stringify({ ...vr, checks }, null, 2));
  fs.writeFileSync(path.join(stage, 'SHA256SUMS'), files.map((f) => `${f.sha256}  downloads/${f.name}`).join('\n') + '\n');
  const injected = gate !== 'BLOCKED' ? `<script type="application/json" id="deck-release">${JSON.stringify({ gate, files: files.map(({ name, path: p, bytes, sha256, label }) => ({ name, path: p, bytes, sha256, label })) })}</script>` : '';
  fs.writeFileSync(path.join(site, 'deck.html'), html.replace(/(?:\.\.\/)+web\//g, 'web/').replace(/<\/head>/i, injected + '</head>'));
  if (gate !== 'BLOCKED') { fs.mkdirSync(path.join(site, 'downloads')); for (const f of files) fs.copyFileSync(path.join(dl, f.name), path.join(site, 'downloads', f.name)); }
  else fs.rmSync(dl, { recursive: true, force: true });   // BLOCKED：不留可下载文件，避免未通过门禁的产物流出
  fs.rmSync(work, { recursive: true, force: true });
  // 原子放入：旧产物先挪开，新产物到位后才删旧的；任何一步失败都还原
  const old = fs.existsSync(out) ? out + '.old-' + crypto.randomBytes(4).toString('hex') : null;
  if (old) fs.renameSync(out, old);
  try { fs.renameSync(stage, out); } catch (e) { if (old) fs.renameSync(old, out); throw e; }
  if (old) fs.rmSync(old, { recursive: true, force: true });
  ok = true;
  console.log(JSON.stringify({ gate, out, verify_status: vr.status, counts: release.counts, files: files.map((f) => ({ name: f.name, bytes: f.bytes, sha256: f.sha256 })), stats: vr.stats?.by_result, limitations: limitations.filter((l) => l.kind !== 'SCOPE').length }, null, 2));
  if (gate === 'BLOCKED') process.exitCode = 1;
} catch (e) { console.error(String(e.message || e)); process.exitCode = 1; }
finally { if (!ok) fs.rmSync(stage, { recursive: true, force: true }); }
