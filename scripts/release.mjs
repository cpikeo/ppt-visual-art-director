#!/usr/bin/env node
/* 构建期发行：deck.html → 像素一致 PDF + 可编辑 PPTX + 核验报告 + 静态下载页。无 API、无队列、无第二套导出器：
   只是顺序调用 web_export.mjs 与 web_verify.py，再把“已校验”的结果与 SHA-256 写进 release.json，并注入 dist 里的 deck.html。
   node scripts/release.mjs deck.html --out dist/NAME [--dpr 2|3] [--force]
   覆盖保护：out 已存在默认拒绝；--force 时先在同级暂存目录完整构建，成功后才替换，失败不动旧产物并清理暂存。
   gate：核验没有 FAIL / BLOCKED 才为 PASS（PowerPoint 实机恒为 UNVERIFIED，披露在 release.json，不阻断下载，也不冒充通过）。 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs'; import path from 'node:path'; import crypto from 'node:crypto'; import { fileURLToPath } from 'node:url';
const HERE = path.dirname(fileURLToPath(import.meta.url)), ROOT = path.join(HERE, '..');
const argv = process.argv.slice(2), deck = argv[0], opt = (k, d) => { const i = argv.indexOf(k); return i >= 0 ? argv[i + 1] : d; }, FORCE = argv.includes('--force');
const die = (m) => { console.error('ERROR ' + m); process.exit(2); };
if (!deck || !opt('--out')) die('usage: release.mjs deck.html --out dist/NAME [--dpr 2|3] [--force]');
const deckPath = path.resolve(deck), out = path.resolve(opt('--out')), dpr = opt('--dpr', '2');
if (!fs.existsSync(deckPath)) die('deck 不存在: ' + deckPath);
if (fs.existsSync(out) && !FORCE) die(`${out} 已存在，未覆盖；换目录或加 --force`);
const PY = process.env.PYTHON || 'python3', version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
const html = fs.readFileSync(deckPath, 'utf8'), meta = (() => { try { return JSON.parse(/<script[^>]*id="deck-meta"[^>]*>([\s\S]*?)<\/script>/.exec(html)?.[1] || '{}'); } catch { return {}; } })();
const slug = String(meta.deck_id || path.basename(path.dirname(deckPath))).toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5]+/g, '-').replace(/^-|-$/g, '') || 'deck';
const stage = fs.mkdtempSync(path.join(path.dirname(out), `.release-${slug}-`).replace(/^(.*)$/, (p) => { fs.mkdirSync(path.dirname(out), { recursive: true }); return p; }));
const sh = (label, cmd, args, env = {}) => { const r = spawnSync(cmd, args, { encoding: 'utf8', cwd: ROOT, env: { ...process.env, ...env } }); return { label, code: r.status, out: (r.stdout || '') + (r.stderr || '') }; };
const sha = (f) => crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex');
let ok = false;
try {
  const dl = path.join(stage, 'downloads'), work = path.join(stage, 'work'), site = path.join(stage, 'site');
  [dl, work, site].forEach((d) => fs.mkdirSync(d, { recursive: true }));
  const E = (...a) => ['scripts/web_export.mjs', ...a];
  const steps = [
    sh('shots', 'node', E('shots', deckPath, '--shots', path.join(work, 'baseline'), '--dpr', dpr)),
    sh('pdf', 'node', E('pdf', deckPath, path.join(dl, `${slug}.pdf`), '--mode', 'pixel', '--dpr', dpr)),
    sh('pptx', 'node', E('pptx', deckPath, path.join(dl, `${slug}.pptx`), '--assets', path.join(work, 'assets')), { PYTHON: PY }),
  ];
  const bad = steps.find((s) => s.code !== 0);
  if (bad) throw new Error(`${bad.label} 失败，未发布任何文件:\n${bad.out.slice(-1500)}`);
  for (const f of fs.readdirSync(dl)) if (/\.plan(\.manifest)?\.json$/.test(f)) fs.renameSync(path.join(dl, f), path.join(work, f));
  const v = sh('verify', PY, ['scripts/web_verify.py', '--pptx', path.join(dl, `${slug}.pptx`), '--manifest', path.join(work, `${slug}.plan.manifest.json`), '--pdf', path.join(dl, `${slug}.pdf`),
    '--pdf-report', path.join(dl, `${slug}.pdf-report.json`), '--shots', path.join(work, 'baseline'), '--render-pptx', '--out', path.join(work, 'verify'), '--json', path.join(work, 'verify.json')]);
  const rep = JSON.parse(fs.readFileSync(path.join(work, 'verify.json'), 'utf8'));
  const gate = rep.checks.some((c) => c.status === 'FAIL' || c.status === 'BLOCKED') ? 'FAIL' : 'PASS';
  // 站点：deck 副本（runtime 路径改写到 ./web）+ 运行时 + assets；只在 gate=PASS 时注入下载清单
  fs.cpSync(path.join(ROOT, 'web'), path.join(site, 'web'), { recursive: true });
  const adir = path.join(path.dirname(deckPath), 'assets'); if (fs.existsSync(adir)) fs.cpSync(adir, path.join(site, 'assets'), { recursive: true });
  const files = fs.readdirSync(dl).filter((f) => /\.(pdf|pptx)$/.test(f)).map((f) => { const p = path.join(dl, f); return { name: f, path: `downloads/${f}`, bytes: fs.statSync(p).size, sha256: sha(p), kind: path.extname(f).slice(1), label: f.endsWith('.pdf') ? 'PDF（像素一致 · 图像页）' : 'PPTX（可编辑 · 原生对象）' }; });
  const pdfRep = JSON.parse(fs.readFileSync(path.join(dl, `${slug}.pdf-report.json`), 'utf8'));
  const release = { name: slug, version, gate, verify_status: rep.status, powerpoint: 'UNVERIFIED', pdf: { mode: pdfRep.mode, dpr: pdfRep.dpr, claim: pdfRep.claim }, stats: rep.stats, files, counts: rep.counts };
  fs.writeFileSync(path.join(stage, 'release.json'), JSON.stringify(release, null, 2));
  fs.copyFileSync(path.join(work, 'verify.json'), path.join(stage, 'verify.json'));
  fs.writeFileSync(path.join(stage, 'SHA256SUMS'), files.map((f) => `${f.sha256}  downloads/${f.name}`).join('\n') + '\n');
  const injected = gate === 'PASS' ? `<script type="application/json" id="deck-release">${JSON.stringify({ gate, files: files.map(({ name, path: p, bytes, sha256, label }) => ({ name, path: p, bytes, sha256, label })) })}</script>` : '';
  fs.writeFileSync(path.join(site, 'deck.html'), html.replace(/(?:\.\.\/)+web\//g, 'web/').replace(/<\/head>/i, injected + '</head>'));
  fs.mkdirSync(path.join(site, 'downloads'), { recursive: true }); for (const f of files) fs.copyFileSync(path.join(dl, f.name), path.join(site, 'downloads', f.name));
  fs.rmSync(work, { recursive: true, force: true });
  // 原子放入：旧产物先挪开，新产物到位后才删旧的；任何一步失败都还原
  const old = fs.existsSync(out) ? out + '.old-' + crypto.randomBytes(4).toString('hex') : null;
  if (old) fs.renameSync(out, old);
  try { fs.renameSync(stage, out); } catch (e) { if (old) fs.renameSync(old, out); throw e; }
  if (old) fs.rmSync(old, { recursive: true, force: true });
  ok = true;
  console.log(JSON.stringify({ status: gate, out, verify_status: rep.status, verify_exit: v.code, files: files.map((f) => ({ name: f.name, bytes: f.bytes, sha256: f.sha256 })), stats: rep.stats?.by_result }, null, 2));
  if (gate !== 'PASS') process.exitCode = 1;
} catch (e) { console.error(String(e.message || e)); process.exitCode = 1; }
finally { if (!ok) fs.rmSync(stage, { recursive: true, force: true }); }
