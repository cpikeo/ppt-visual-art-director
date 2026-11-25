---
name: ppt-visual-art-director
description: 做 / 改 / 评审 PPT 时使用。以视觉艺术总监的判断完成汇报、路演、发布、复盘、融资材料：读懂受众与证据，定原创视觉概念，再定构图、排印、色彩、图像、图表与整稿节奏；由 Plan 契约确定性编译为原生可编辑 PPTX（原生标题、列表、图表、表格，主题与母版跟随本稿），看真实渲染后定点修正；也用于诊断与定点改稿。不适用于纯文案撰写、Word/PDF 排版。
version: 18.3.0
---

# PPT 视觉艺术指导

**AI 判断**（概念、取舍）；**脚本算术与事实**（网格、对比度、框高、对象、字节），不给审美结论；**人签收**。

## 一、先判任务，只读需要的；不中途提问

信息不全就按最小假设推进，把假设写在开头与 `open_questions`，不补造数字、不停下来问。

| 任务 | 做法 | 只读 |
| --- | --- | --- |
| 评审、专项（配色、某一页） | 只答被问的 | `review.md` |
| 改既有稿 | `revise.py inspect` → 定点编辑，不重做设计 | `revise.md` |
| 新建，页少、信息单一 | 第二节 1–4 步 + 编译核验 | `design.md`；图表、排版、色彩加 `craft.md` |
| 新建，整稿、路演、发布 | 第二节全流程 + 整稿节奏复核 | `design.md` `build.md`；有图加 `imagery.md` |

按需：手写元素、图表、素材登记 `elements.md`；诊断码、退出码、能力边界 `ops.md`。**字段不确定先跑 `scaffold.py fields <root|page|element|chart|…>`（约 1KB），不要通读 `plan.schema.json`。**

## 二、判断流程（新建）

风格词（高级、极简、科技感）不是方案。答得出五问，才选字体与颜色：

1. **受众与主张**：谁看、已知什么、怀疑什么、在哪看；离场记住哪一句。→ `audience` `goal` `core_claim`
2. **洞察与概念**：观众原以为什么，证据说明什么；用什么**视觉关系**让差距被看见，**否决**什么。→ `brief.insight` `concept` `veto`
3. **编辑取舍**：放大、退隐、删除什么。→ `brief.editorial`（可选）
4. **艺术指导与系统**：光、景别、材质、唯一强调色、颗粒；字体、色彩、空间各一句理由；一两个从内容里长出来的跨页母题。→ `brief.art`（有图时）`brief.system`（可选）`brief.motif`
5. **叙事**：每页 `role`（open build turn breathe hold close）与一句“做什么、多重”；最重的一页是最重要的一页吗？→ `role` `rhythm`

再落到几何：`scaffold.py system` 看色彩与字号推导；十一种配方与 `free` 只是起点，关系不在配方里就用 `free`，网格每页自选；页内声明 `focus_element` 与 `depth`，`composition_rationale` 写清“先看到什么、凭什么、放弃了什么”。

**自检**：概念能追溯到洞察、能否掉具体的东西、换主题就不成立；每页先看到什么、凭什么信、删掉哪件会损失什么。

## 三、判断原则

八条原则与上游优先顺序见 `design.md` 开头，新建前必读。

## 四、命令

`plan.json` 是设计与机器之间唯一的契约（`scripts/plan.schema.json`；语义见 `build.md` `elements.md`）。坐标英寸、原点左上，16:9 为 13.333 × 7.5；字号 pt；叠放顺序即声明顺序。

```bash
pip install -r requirements.txt                                       # 渲染核验另见 requirements-render.txt
python scripts/scaffold.py system --seed "#B4421E" --mood warm        # 色彩系统 + 字号阶梯（只读）
python scripts/scaffold.py fields [root|page|element|chart|…]        # 紧凑字段表（只读）
python scripts/scaffold.py init outline.json plan.json [--assets DIR] # outline → plan 初稿
python scripts/scaffold.py migrate old.json new.json                 # 旧 plan → schema_version 2
python scripts/compile.py plan.json out.pptx [--assets DIR] [--production]   # 校验+编译；--production：draft 即阻断（同环境同输入 → 同字节）
python scripts/verify.py plan.json out.pptx [--assets DIR]           # 对象核对 + 渲染核对
python scripts/revise.py inspect 原稿.pptx                            # 改稿：只读解析
python scripts/revise.py 原稿.pptx edit.json out.pptx                 # 改稿：定点编辑 + 保留核对
```

成品原生：页内第一个顶层 `title` 是原生标题占位符；`list` 是真项目符号；图表与表格是原生对象；不整页栅格化；不自动缩字、换色、补数据、联网取图。`design_state`：`draft`／`judged`／`confirmed`，交付前须无 draft。只产出 `.pptx`。开发材料（测试、案例、打包）在仓库 `dev/`，不属于本 Skill 包，不要读入上下文。

## 五、验收

三层各自成立，不合并成分数：对象核对、渲染核对（LibreOffice → PDF → PNG）、人的判断。`verify.py`：`READY` 0｜`BLOCK` 1｜`PARTIAL` 4（渲染没跑、字体被替换）｜输入或契约错误 2，都不是质量结论，`powerpoint` 恒为 `NOT VERIFIED`。没做的核验不说成做了；看图顺序、“不能说”清单与交付回复六项见 `review.md`。
