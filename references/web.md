# Web Presentation：工程契约与跨格式发行

新建整稿的默认交付是一个 `deck.html`（+ `assets/`）。Python 引擎（`plan.json → .pptx`）保留，用于**原生 PPTX 直做**与**导出落地**。

## 一、流程（每步只回答自己层级的问题）

内容理解 → 叙事 → 视觉概念（`design.md`）→ 视觉系统（字体、色彩、母题）→ 逐页构图 → 动态艺术指导（`motion.md`）→ 写 `deck.html` → `audit` → 看截图修正 → 导出 PDF/PPTX → `web_verify.py` → 签收。`audit` 通过只证明工程项，不证明好看。

## 二、deck.html 契约

```html
<link rel="stylesheet" href="web/deck.css">                      <!-- 路径按实际 -->
<script type="application/json" id="deck-meta">{deck_id,title,audience,goal,core_claim,brief:{insight,concept,veto[]},fonts,palette}</script>
<main class="deck">
  <section class="slide" data-role="open|build|turn|breathe|hold|close" data-focus data-concept data-rationale data-transition>
    <h1|h2 data-m="rise" data-at=".2">…</h1> … <aside class="notes">演讲备注</aside>
  </section>
</main>
<script src="web/deck.js"></script>
```
- 画布 **1280×720 px**（= 13.333×7.5 in，96 px = 1 in，字号 pt = px×0.75）。运行时整体缩放，布局只在 1280×720 内思考；安全边 ≥ 64 px，右下角 ≥ 1151,651 留给页码类信息。
- 每页一个 `h1/h2`（大纲、读屏、PPTX 标题占位符）；`data-role` `data-focus` 必填，对应 plan 的 `role` `focus`；`aside.notes` 进入演讲者面板与 PPTX 备注。
- 自由构图：`.slide` 内元素用 `position:absolute` + px 坐标；本页的视觉概念写进 `data-concept`，取舍写进 `data-rationale`。风格由 deck 自带 `<style>`，运行时只管引擎。
- 运行时（只读 `web/deck.js`，勿改）：键盘 ←→/空格/Home/End/数字⏎，O 总览，P 演讲者面板（备注、计时、下一页），F 全屏，B 黑屏，? 帮助；触摸滑动；`#页码` + localStorage 刷新恢复；`inert`/`aria-live`；`Deck.show(i)` 终态（QA 与导出用）。
- 图表：`<div data-chart='{"kind":"bar|line|table|…","data":…}'>`，**HTML 视图由 deck 自己用同一份 JSON 画**；导出时整块变成原生图表/表格。字段同 `scaffold.py fields chart`。

## 三、命令（需 Node ≥ 18 + `npm i` + `npx playwright install chromium`；Python 依赖见 requirements.txt）

```bash
node scripts/web_export.mjs audit deck.html --shots DIR     # 静态审计 + 真实键盘通关 + 375/768/1440 + 帧间隔，输出 PASS/WARN/FAIL
node scripts/web_export.mjs pdf   deck.html out.pdf         # print 样式，矢量文字，1280×720 逐页
node scripts/web_export.mjs pptx  deck.html out.pptx --assets DIR   # 终态 DOM → plan.json（+ manifest）→ compile.py --web
python scripts/web_verify.py --pptx out.pptx --manifest out.plan.manifest.json --pdf out.pdf --shots DIR --out DIR2
```
CJK 字体缺失时 audit 报 `WARN`（PDF 会替换；PPTX 取决于目标机）。无头 Linux 缺系统库时先装 Chromium 依赖。

## 四、元素级导出策略（唯一口径）

| HTML | PPTX | PDF | 说明 |
| --- | --- | --- | --- |
| 文字（含 `<b>` `<em>` `<span>` 逐 run 颜色/字重/字号/字体） | 原生文本框 | 矢量 | `<br>` = 段落；行高换算为倍数；框宽放大 2%（单行 8%）防换行差异 |
| 背景色 / 边框 / 圆角 / 椭圆 | 原生形状 | 矢量 | 单边边框 → 细矩形；四边同色 → 描边 |
| 线性渐变（≥3 色取首末色，写 `W-GRADIENT`） | 原生渐变 | 矢量 | |
| `<img>` cover + `object-position` / CSS `background-image:url` | 原生图片（cover 裁切） | 位图 | 须是 png/jpg/webp；svg 图片先转位图 |
| `data-chart` | 原生图表/表格 | 矢量 | |
| `<svg>` `<canvas>` `filter` `backdrop-filter` `mix-blend-mode` `clip-path` `mask` 缩放/斜切变换 圆角图片 复杂渐变 `data-x="raster"` | **仅该元素**栅格（2× PNG，透明底）+ `W-RASTER` | 矢量/位图 | 不可编辑；报告逐项列出 |
| `data-x="skip"` | 跳过 | 保留 | 仅 HTML 的装饰 |
| `box-shadow` `text-shadow` `text-decoration` | 丢弃 + `W-SHADOW`/`W-DECOR` | 保留 | 无原生等价 |
| 动画、转场、构建步骤、交互 | 终态 | 终态 | 不支持，不假装 |

**绝不整页截图当 PPTX**（`web_verify.py` 检查：无 ≥90% 页面的图片）。`--web` 只把 `E-OVERLAP` `E-PAGE` 降为 SOFT（版面权威是已验收的 HTML），页码默认隐藏，其余 compile 校验不变；12pt 以下等 `W-PT` 仍提示。

## 五、能力边界（必须原样告知用户）

- PPTX 与 HTML 的**换行、字重、行高**可能有像素级差异；未装字体被替换。本包只能做对象级（几何、文本、类型）核验；**渲染还原需 LibreOffice 或 PowerPoint 实机**，否则 `UNVERIFIED`。
- PDF 里 Chromium 把 CJK 字体嵌为 Type3 轮廓：可选中、不可换字体编辑。
- 无：动画/转场导出、嵌入字体、SVG 原生矢量导入、阴影/发光/混合模式原生化。

## 六、验收口径

PASS｜FAIL｜BLOCKED（环境缺失，写原因）｜UNVERIFIED（没执行或无法判定）。没跑的不写通过。交付回复列：做了什么、`audit` 与 `web_verify` 结果、栅格/降级清单、UNVERIFIED 项。
