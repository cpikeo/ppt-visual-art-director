---
name: ppt-visual-art-director
description: 做 / 改 / 评审演示（汇报、路演、发布、复盘、融资）。先定受众、洞察、叙事与原创视觉概念，交付动态网页演示（HTML，16:9，含运动语法与演讲者工具），并可逐元素导出为矢量 PDF 与原生可编辑 PPTX；也可由 Plan 契约直做 PPTX、诊断与定点改稿。不适用于纯文案、Word 排版。
version: 22.1.0
---

# PPT 视觉艺术指导

两项交付：**Web 演示**（`deck.html`，新建默认）；**跨格式发行**（Web → PDF / PPTX，元素级分级，不整页截图）。原生 PPTX 直做与改稿走 Python 引擎。内容决定形式，关系决定构图，系统建立秩序，动态服务叙事。

**AI 判断**（概念、取舍）；**脚本算术与事实**（网格、对比度、框高、对象、字节）；**人签收**。信息不全按最小假设推进，假设写进 `open_questions`，不补造数字、不中途提问。

## 一、先判任务，只读需要的

| 任务 | 做法 | 只读 |
| --- | --- | --- |
| 评审、专项（配色、某页、动效） | 只答被问的 | `review.md`；动效加 `motion.md` |
| 改既有 PPTX | `revise.py inspect` → 定点编辑，不重做设计 | `revise.md` |
| 新建 Web 演示（默认） | 判断链 → 写 `deck.html` → `audit` → 看图修正 | `design.md` `web.md` `motion.md`；有图加 `imagery.md`；图表/排版加 `craft.md` |
| 另需 PDF / PPTX | `web_export.mjs pdf｜pptx` → `web_verify.py` | `web.md` 三～六节 |
| 只要 PPTX | 判断链 + `scaffold init` → `compile` → `verify` | `design.md` `build.md` |

按需：手写元素、图表、素材 `elements.md`；诊断码、退出码、能力边界 `ops.md`。**字段不确定先跑 `scaffold.py fields <root|page|element|chart|brief|…>`，不要通读 schema。**

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

`plan.json` 是设计与机器之间唯一的契约（`scripts/plan.schema.json`）。坐标英寸、原点左上；**画布唯一 16:9 = 1280×720 px = 13.333…×7.5 in（96 px = 1 in）**，字号 pt，叠放即声明顺序。

```bash
pip install -r requirements.txt                          # 渲染核验另见 requirements-render.txt
python scripts/scaffold.py system --seed "#B4421E" --mood warm | fields [..] | canvas   # 只读
python scripts/scaffold.py init outline.json plan.json [--assets DIR]   # outline → plan 初稿
python scripts/compile.py plan.json out.pptx [--assets DIR] [--production]
python scripts/verify.py plan.json out.pptx [--assets DIR]
python scripts/revise.py inspect 原稿.pptx | revise.py 原稿.pptx edit.json out.pptx
npm i && npx playwright install chromium                 # Web 路径（Node ≥ 18）
node scripts/web_export.mjs audit|pdf|pptx deck.html …   # 默认不覆盖；只有用户明确要求才加 --force
python scripts/web_verify.py --pptx out.pptx --manifest out.plan.manifest.json --pdf out.pdf --shots DIR
```

运行时 `web/`（`deck.css` `deck.js`，图表用 `chart.js`）随 deck 拷贝、不要改。动态按 `motion.md` 七项语法，每个运动要有叙事意图；动画只进 HTML，PDF/PPTX 是终态，降级清单必须写进交付说明。成品原生可编辑（占位符标题、真项目符号、原生图表与表格），不整页栅格化，不自动缩字、换色、补数据、联网取图。`design_state` 交付前须无 `draft`。

## 四、验收

三层各自成立，不合并成分数：对象核对、渲染核对、人的判断。状态口径唯一：**PASS｜FAIL｜BLOCKED｜UNVERIFIED**（PPTX 引擎另有 `READY｜BLOCK｜PARTIAL`）。BLOCKED / UNVERIFIED 永远不是 PASS：`web_verify.py` 汇总 FAIL > BLOCKED > UNVERIFIED > PASS，退出码 0/1/2(输入错)/3/4 与之一一对应；`powerpoint` 恒为 `NOT VERIFIED`。没做的核验不说成做了。诊断码与退出码见 `ops.md`；看图顺序、“不能说”清单与交付回复六项见 `review.md`。
