---
name: ppt-visual-art-director
description: 做 PPT / 改 PPT / 评审 PPT 时使用。以视觉艺术总监的判断完成汇报、路演、发布、复盘、融资材料：读懂受众与证据，定原创视觉概念，再决定构图、排印、色彩、图像、图表与整稿节奏；由 Plan 契约确定性编译为原生可编辑 PPTX（原生标题占位符、列表、图表、表格，主题与母版跟随本稿），看真实渲染后定点修正；也用于诊断与定点改稿。不适用于纯文案撰写、Word/PDF 排版。
version: 18.2.0
---

# PPT 视觉艺术指导

让一个判断被看见，并准确地做出来。**AI 负责判断**（概念、取舍）；**脚本负责算术与事实**（网格、对比度、框高、对象、字节），不给审美结论；**人负责签收**。

## 一、先判任务大小，只读需要的

| 任务 | 做法 | 只读 |
| --- | --- | --- |
| 评审、专项（配色、某一页） | 只答被问的，不跑全流程 | `review.md` |
| 改既有稿 | `revise.py inspect` → 定点编辑，不重做设计 | `revise.md` |
| 新建（页少、信息单一） | 第二节 1–4 步 + 编译核验 | `design.md`；涉及图表、排版、色彩加 `craft.md` |
| 新建（整稿、路演、发布） | 第二节全流程 + 整稿节奏复核 | `design.md` `build.md`；有图像加 `imagery.md`；看图读 `review.md` |

按需：图表、排版、色彩 `craft.md`；手写元素、图表、素材登记 `elements.md`；诊断码、退出码、能力边界 `ops.md`。模式不明时按字面最小范围做，开头说明模式与假设。**字段不确定先跑 `scaffold.py fields <root|page|element|chart|…>`（约 1KB），不要通读 `plan.schema.json`。**

## 二、判断流程（新建）

风格词（高级、极简、科技感）不是方案。答得出五问，才有资格选字体与颜色：

1. **受众与主张**：谁看、已知什么、怀疑什么、在哪看；离场记住哪一句。→ `audience` `goal` `core_claim`
2. **洞察与概念**：观众原以为什么，证据说明什么；用什么**视觉关系**让差距被看见，**否决**什么。→ `brief.insight` `concept` `veto`
3. **编辑取舍**：放大什么、退隐什么、删除什么。→ `brief.editorial`
4. **艺术指导与系统**：光、景别、材质、色域（含唯一强调色）、颗粒；字体、色彩、空间各一句理由；一到两个从内容里长出来的跨页母题。→ `brief.art`（有图时）`brief.system` `brief.motif`
5. **叙事**：每页 `role`（open build turn breathe hold close）与一句“做什么、多重”；最重的一页是最重要的一页吗？→ `role` `rhythm`

再落到几何：`scaffold.py system` 看色彩与字号推导（`design.ground / accent / data` 可覆盖，文字对比仍由脚本保证）；每页一种讲法——十一种配方与 `free` 只是起点，关系不在配方里就用 `free`，网格每页自选；页内声明 `focus_element` 与 `depth`，在 `composition_rationale` 写清“先看到什么、凭什么、放弃了什么”；编译、看图、修正。

**概念四问**：能追溯到洞察吗？能否否掉一个具体的东西？换掉主题还成立吗？缩略图上认得出是这一套吗？　**每页三问**：先看到什么？凭什么信？删掉哪件会损失什么？

## 三、判断原则

1. **概念先于风格；原创来自关系。** 换配色、字体、圆角改的是表皮；原创是只属于这份内容的尺度、裁切、并置或读法。
2. **关系先于版式。** 先判内容是比较、时间、因果、流程、组成还是层级，再选图表与空间。相关与先后都不是因果。
3. **主次先于均衡。** 每页一个视觉主导，靠尺度、位置、明暗、留白建立，灰度下也成立。固定位置的重复装饰不是母题。
4. **编辑先于缩放。** 装不下就删、合并、改写、拆页，不压字号。
5. **色彩是关系系统。** 先定明度与职责，再谈色相；强调色只指向结论；同一对象跨页同色。
6. **真实先于完整。** 没证据的漂亮页不如留白；生成图只承担氛围，不替代产品、人物、数据与事实；缺的数字写进 `open_questions`，不补造。
7. **判断先于规则。** 密集比较守网格、字阶与色彩职责；尺度、裁切或留白本身是论据时破格，在 `composition_rationale` 写明破了哪条、换来什么、谁承担代价；整稿只破少数几次。
8. **停止先于完美。** 目标达成、系统成立即停；两轮无实质改善，交给人决策。

顺序（上游优先）：内容与信息架构 → 主导关系与构图 → 空间与节奏 → 排印与层级 → 图像与色彩 → 工艺。上层问题用下层装饰修补，必然失败。

## 四、命令

`plan.json` 是设计与机器之间唯一的契约：字段以 `scripts/plan.schema.json` 为准，语义见 `build.md` `elements.md`。坐标英寸、原点左上，16:9 为 13.333 × 7.5；字号 pt；叠放顺序即声明顺序。

```bash
pip install -r requirements.txt                                      # 渲染核验另见 requirements-render.txt
python scripts/scaffold.py system --seed "#B4421E" --mood warm      # 色彩系统 + 字号阶梯（只读）
python scripts/scaffold.py fields [root|page|element|chart|…]       # 紧凑字段表（只读）
python scripts/scaffold.py init outline.json plan.json [--assets DIR] # outline → plan 初稿
python scripts/scaffold.py migrate old-plan.json new-plan.json      # 旧 plan → schema_version 2
python scripts/compile.py plan.json out.pptx [--assets DIR] [--production]   # 校验+编译；--production：draft 即阻断；同环境同输入 → 同字节
python scripts/verify.py plan.json out.pptx [--assets DIR]          # 对象核对 + 渲染核对
python scripts/revise.py inspect 原稿.pptx                           # 改稿：只读解析
python scripts/revise.py 原稿.pptx edit.json out.pptx                # 改稿：定点编辑 + 保留核对
```

成品原生：页内第一个顶层 `title` 是原生标题占位符；`list` 是真项目符号；图表与表格是原生对象；主题与母版跟随本稿；不整页栅格化；不自动缩字、换色、补数据、联网取图。类别 ≤ 8 的柱/饼默认直标读数。`design_state`：`draft`／`judged`／`confirmed`，交付前须无 draft。

边界：只产出 `.pptx`；路径守卫、素材根、外部 PPTX 的 XML 安全（DTD/实体/损坏 XML 以 `E-FORMAT` 拦下）见 `ops.md`。测试、案例与评测不属于本 Skill 包。

## 五、验收

三层各自成立，不合并成分数：**对象核对**、**渲染核对**（LibreOffice → PDF → PNG）、**人的判断**。`verify.py` 退出码：`READY` 0 ｜ `BLOCK` 1 ｜ `PARTIAL` 4（渲染没跑、字体被替换）｜ 输入或契约错误 2；都不是质量结论，`powerpoint` 恒为 `NOT VERIFIED`。**没做的核验不能说成做了**（没渲染不说“已预览”“无溢出”，`PARTIAL` 不说“通过”，`READY` 不说“设计通过”）。看图顺序、缺陷分级、“不能说”全表与**交付回复六项**见 `review.md` 十一、十二。
