# ppt-visual-art-director

**v5.0.0** — 让 AI 以视觉设计总监的方式做 PPT：先把判断写成 brief，再把设计写成 plan，最后确定性地落成原生可编辑 PPTX，并看真实渲染定点修正。

它不做模板套用，也不从行业关键词推风格。工作方式是：读懂受众与证据 → 找一条洞察 → 定一个能约束全篇的概念（含它否决了什么）→ 定艺术指导 → 排叙事节奏 → 逐页构成 → 编译 → 看真实渲染 → 定点修正。

> 一句话原则：设计判断归你，确定性执行归脚本，审美与签收归人。机器检查不等于好看，预览不等于目标环境通过。

## 结构

```
SKILL.md                  判断原则、判断顺序、brief、任务路线、执行、验收、交付
references/
  design.md               概念与原创性、艺术指导、叙事、视觉系统、构图、版式语法、
                          信息设计、排版、色彩、图像与色域压制、诊断、复核维度
  build.md                plan 语义、坐标、叠放、文本与 runs、图表口径、素材、页码、
                          网格、命令、诊断码、机器边界
scripts/
  compile.py              plan.json → 原生可编辑 PPTX（校验 + 编译 + 整稿设计信号）
  verify.py               plan ↔ PPTX ↔ 真实渲染（对象核对、渲染核对、证据报告）
  inspect_pptx.py         只读解析既有 PPTX（改稿与评审第一步）
  grid.py                 列数 → 英寸坐标，标出页码保留框
  plan.schema.json        数据契约（字段形状与类型条件）
  limits.json             技术阈值、默认值、字体别名
requirements.txt
```

## 安装

```bash
python -m venv .venv && . .venv/bin/activate      # Windows：.venv\Scripts\activate
pip install -r requirements.txt
```

Python 3.9+。脚本不联网。真实渲染核验另需 LibreOffice；缺失时 `verify.py` 会跳过渲染并说明哪些判断因此没有依据。

## 三步

```bash
python scripts/grid.py --span 1:6 --span 8:5      # 先算坐标
python scripts/compile.py plan.json out.pptx      # 再编译（已有同名文件需 --force）
python scripts/verify.py plan.json out.pptx       # 最后核验（写 render/）
```

plan 的最小骨架见 `references/build.md`「最小骨架」。改稿先跑 `python scripts/inspect_pptx.py 原稿.pptx`。

## 能做什么

- **判断有落点**：`brief` 里写洞察、概念、**否决项**、艺术指导五参数与母题；`veto` 至少一条——否决不了任何东西的概念没有指导力。
- **原生可编辑**：文本（逐段逐 run 混排，`runs` 可跨段换色换字重）、图片（裁切、透明度、`duotone` 压进整稿色域）、形状（圆角半径、描边、旋转、渐变与半透明）、路径（折线、箭头、闭合多边形、锚点连接）、组合、表格、图表。
- **图表**：柱（横／竖／堆叠／百分比）、线、面积（含堆叠）、饼、环、散点、柱＋线组合，直标分柱层／折线层／饼环三套规则，轴线与轴名可控。
- **整稿信号**：连续同构、标题下细线、无功能英文标签、素材复用、待主张页、满幅图开场——只指出位置，判断归人。
- **确定性**：叠放顺序、页码、几何、样式全部由 plan 决定；编译前阻断硬错误，不静默改设计。

## 四条硬事实

- **坐标单位是英寸**，原点在页面左上角；16:9 画布 13.333 × 7.5，4:3 是 10 × 7.5；字号单位是 pt。`path` 与 `group` 的几何由点集或子元素推出。
- **编译器总是生成新文件**，不就地修改原稿；目标文件已存在时默认不动，输出与报告都不得指向 plan 自己。
- **视觉质量只有在真实渲染且确实看过图像时才可能被判断**。缺任一条件，如实说未核验，不要说"已通过"。
- **`duotone` 在 PowerPoint 生效，LibreOffice 只做去饱和**：渲染核验能证明结构与明暗，不能证明色相。

## 从 v4 到 v5

删掉单句 `visual_concept`，换成结构化的 `brief`（洞察／概念／否决／艺术指导／母题）——判断不再只活在对话里。新增 `duotone`、`corner_in`、`pad_in`、`fonts.number`、跨段 `runs`，修正组合图与饼环直标的分层规则；去掉 `shared_basis` 这个可以绕过口径检查的开关；阻断输出路径覆盖 plan 的数据丢失路径；诊断码与级别对齐（`E-` / `W-` / `I-`）。
