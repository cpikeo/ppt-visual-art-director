---
name: ppt-visual-art-director
description: 做 / 改 / 评审 PPT（汇报、路演、发布、复盘、融资）。先定受众、洞察与原创视觉概念，再由 Plan 契约确定性编译为原生可编辑 PPTX，看真实渲染后定点修正；也用于诊断与定点改稿。不适用于纯文案、Word/PDF 排版。
version: 20.6.0
---

# PPT 视觉艺术指导

**AI 判断**（概念、取舍）；**脚本算术与事实**（网格、对比度、框高、对象、字节）；**人签收**。信息不全按最小假设推进，假设写进 `open_questions`，不补造数字、不中途提问。

## 一、先判任务，只读需要的

| 任务 | 做法 | 只读 |
| --- | --- | --- |
| 评审、专项（配色、某页） | 只答被问的 | `review.md` |
| 改既有稿 | `revise.py inspect` → 定点编辑，不重做设计 | `revise.md` |
| 新建，页少、信息单一 | 第二节 1–4 步 + 编译核验 | `design.md`；图表/排版/色彩加 `craft.md` |
| 新建，整稿、路演、发布 | 第二节全流程 + 整稿节奏复核 | `design.md` `build.md`；有图加 `imagery.md` |

按需：手写元素、图表、素材登记 `elements.md`；诊断码、退出码、能力边界 `ops.md`。**字段不确定先跑 `scaffold.py fields <root|page|element|chart|brief|…>`（约 1KB），不要通读 schema。**

## 二、判断链（新建）

内容理解 → 叙事判断 → 视觉概念 → 视觉系统 → 页面架构 → 视觉编排 → 真实渲染 → 整稿校准 → **停止**（系统成立即停；两轮无实质改善交给人）。每一步只回答自己层级的问题，上层问题不用下层装饰修补：

| 层级 | 回答什么 | 落点字段 |
| --- | --- | --- |
| 整体视觉世界 | 受众与主张；洞察（观众原以为什么，证据说明什么）；概念、否决项；光、景别、材质、色域 | `audience` `goal` `core_claim` `brief.insight/concept/veto/art` |
| 章节叙事节奏 | 谁最重、何处停顿与转折；放大、退隐、删除什么 | `role` `rhythm` `brief.editorial/motif` |
| 页面空间构成 | 先看到什么、凭什么、放弃了什么；空间气质、空白的功能；网格每页自选 | `focus` `composition`（含 `temperament` `negative_space` `hero_ratio`） `composition_rationale` |
| 局部视觉执行 | 字体、墨阶、裁切、标注、材质 | 元素字段（`craft.md`） |

风格词（高级、极简、科技感）不是方案；概念须能追溯到洞察、换主题就不成立。十一种配方与 `free` 只是起点，关系不在配方里就用 `free`（`free` 是原创构图的主要出口）。原则与判断框架见 `design.md`。

## 三、命令

`plan.json` 是设计与机器之间唯一的契约（`scripts/plan.schema.json`；语义见 `build.md` `elements.md`）。坐标英寸、原点左上（16:9 = 13.333 × 7.5），字号 pt，叠放即声明顺序。

```bash
pip install -r requirements.txt                                       # 渲染核验另见 requirements-render.txt
python scripts/scaffold.py system --seed "#B4421E" --mood warm        # 色彩系统 + 字号阶梯（只读）
python scripts/scaffold.py fields [root|page|element|chart|…]        # 紧凑字段表（只读）
python scripts/scaffold.py init outline.json plan.json [--assets DIR] # outline → plan 初稿（写出前先过 schema；seed/mood/fonts 写在 outline 的 design 下）
python scripts/scaffold.py migrate old.json new.json                 # 旧 plan → schema_version 2
python scripts/compile.py plan.json out.pptx [--assets DIR] [--production]   # 校验+编译；--production：draft 即阻断
python scripts/verify.py plan.json out.pptx [--assets DIR]           # 对象核对 + 渲染核对
python scripts/revise.py inspect 原稿.pptx                            # 改稿：只读解析
python scripts/revise.py 原稿.pptx edit.json out.pptx                 # 改稿：定点编辑 + 保留核对
```

整页空间合成（图、背景、字如何共同建立一个“场”）见 `design.md` 六。图像可用 `crop_focus` + `crop_zoom`（1–8）从同一素材裁出“整体 + 局部”，配合 `mask`、层叠与出血做不规则构成（`design.md` 五）。成品原生可编辑（标题占位符、真项目符号、原生图表与表格），不整页栅格化；不自动缩字、换色、补数据、联网取图。`design_state` 交付前须无 `draft`。只产出 `.pptx`。

## 四、验收

三层各自成立，不合并成分数：对象核对、渲染核对（LibreOffice → PDF → PNG）、人的判断。`verify.py` 状态 `READY`｜`BLOCK`｜`PARTIAL`（渲染没跑、字体被替换或字面无法判定）都不是质量结论；`powerpoint` 恒为 `NOT VERIFIED`。没做的核验不说成做了。退出码、诊断码见 `ops.md`；看图顺序、“不能说”清单与交付回复六项见 `review.md`。
