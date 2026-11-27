---
name: ppt-visual-art-director
description: 做 / 评审演示（汇报、路演、发布、复盘、融资）。先定受众、洞察、叙事与原创视觉概念，以 HTML 为唯一创作源：交付动态网页演示（16:9，含运动语法与演讲者工具、沉浸式全屏），并由同一份 HTML 导出像素一致 PDF 与原生可编辑 PPTX（元素级分级，不整页截图）。不适用于纯文案、Word 排版，也不提供独立的原生 PPTX 起稿或既有 PPTX 改稿入口。
version: 23.0.0
---

# PPT 视觉艺术指导

**HTML 是唯一创作源**，三种终态：**Web 演示**（`deck.html`）、**PDF**（像素模式：逐页无损 PNG 图像页，与屏幕逐像素一致；文字不可选中）、**PPTX**（可编辑：文字/形状/图片/表格/图表为原生对象，复杂 CSS 仅对该元素局部栅格并披露）。`compile.py` 是 PPTX 导出引擎，不是第二个创作入口；既有 PPTX 改稿与 Outline→Plan 起稿已移出本包（见 CHANGELOG 23.0.0，旧版在 git 标签 `v22.1.0-pre-trim`）。内容决定形式，关系决定构图，系统建立秩序，动态服务叙事。

**AI 判断**（概念、取舍）；**脚本算术与事实**（网格、对比度、框高、对象、字节）；**人签收**。信息不全按最小假设推进，假设写进 `open_questions`，不补造数字、不中途提问。

## 一、先判任务，只读需要的

| 任务 | 做法 | 只读 |
| --- | --- | --- |
| 评审、专项（配色、某页、动效） | 只答被问的 | `review.md`；动效加 `motion.md` |
| 新建演示（默认） | 判断链 → 写 `deck.html` → `audit` → 看图修正 | `design.md` `web.md` `motion.md`；有图加 `imagery.md`；图表/排版加 `craft.md` |
| 要 PDF / PPTX / 网页下载 | `release.mjs`（= shots + pdf + pptx + web_verify + SHA-256） | `web.md` 三～六节 |

按需：图表字段、素材、诊断码 `elements.md`；诊断码、退出码、能力边界 `ops.md`。字段以 `scripts/plan.schema.json` 为准，图表缺省值只看 `scripts/limits.json` 的 `chart.contract`。

## 二、判断链（新建）

内容理解 → 叙事判断 → 视觉概念 → 视觉系统 → 页面架构 → 视觉编排 → 真实渲染 → 整稿校准 → **停止**（系统成立即停；两轮无实质改善交给人）。每层只回答自己的问题，上层问题不用下层装饰修补：

| 层级 | 回答什么 | 落点字段 |
| --- | --- | --- |
| 视觉世界 | 受众与主张；洞察（观众原以为什么，证据说明什么）；概念、否决项；光、景别、材质、色域 | `audience` `goal` `core_claim` `brief.insight/concept/veto/art` |
| 章节节奏 | 谁最重、何处停顿与转折；放大、退隐、删除什么 | `role` `rhythm` `brief.editorial/motif` |
| 页面构成 | 先看到什么、凭什么、放弃了什么；空白的功能 | `focus` `composition`（`temperament` `negative_space` `hero_ratio`） `composition_rationale` |
| 局部执行 | 字体、墨阶、裁切、标注、材质 | 元素字段（`craft.md`） |

风格词（高级、极简、科技感）不是方案；概念须能追溯到洞察、换主题就不成立。十一种配方与 `free` 只是起点，关系不在配方里就用 `free`。**交付前四问**（答不出回上一层，不靠加装饰）：第一眼是什么、凭什么？删掉每个图形元素后重心与读序还成立吗？图、背景、字是否同一个“场”（光向、明度结构、边界语言）？概念换主题还成立吗——成立就是风格词，重写。

## 三、命令

画布唯一 **16:9 = 1280×720 px = 13.333…×7.5 in（96 px = 1 in）**，字号 pt = px×0.75。`plan.json` 是 HTML 与 PPTX 引擎之间的内部契约，由 `web_export.mjs` 生成，不手写。

```bash
pip install -r requirements.txt                 # python-pptx 等；渲染核验另见 requirements-render.txt
npm ci && npx playwright install chromium       # Node ≥ 20；Playwright 精确锁 1.65.1
node scripts/web_export.mjs audit deck.html --shots DIR        # 静态审计 + 真实键盘通关 + 375/768/1440
node scripts/release.mjs deck.html --out dist/NAME             # 构建期一次发行：PDF（像素）+ PPTX + 核验 + SHA256SUMS + 静态下载页
node scripts/web_export.mjs pdf deck.html out.pdf [--mode pixel|vector] [--dpr 2|3]   # 缺省 pixel
node scripts/web_export.mjs pptx deck.html out.pptx --assets DIR
python scripts/web_verify.py --pptx … --manifest … --pdf … --shots DIR --render-pptx --out DIR2 --json r.json
python scripts/verify.py plan.json out.pptx                    # 可选：对象层深度核验（读 release 暂存的 plan.json）
node scripts/sync_contract.mjs [--check]                       # limits.json → web/contract.js
```

运行时 `web/`（`deck.css` `contract.js` `chart.js` `deck.js`）随 deck 拷贝、不要改；`deck.html` 末尾按 `contract.js → chart.js → deck.js` 顺序加载。动态按 `motion.md`；动画只进 HTML，PDF/PPTX 是终态。成品原生可编辑（占位符标题、真项目符号、原生图表与表格），不整页栅格化，不自动缩字、换色、补数据、联网取图。

## 四、验收

三层各自成立，不合并成分数：对象核对、渲染核对、人的判断。状态口径唯一：**PASS｜FAIL｜BLOCKED｜UNVERIFIED**。BLOCKED / UNVERIFIED 永远不是 PASS：`web_verify.py` 汇总 FAIL > BLOCKED > UNVERIFIED > PASS，退出码 0/1/2(输入错)/3/4 与之一一对应；`powerpoint` 恒为 `NOT VERIFIED`。没做的核验不说成做了。诊断码与退出码见 `ops.md`；看图顺序、“不能说”清单与交付回复六项见 `review.md`。
