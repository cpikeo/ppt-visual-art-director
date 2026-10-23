# 交付审计记录 · ppt-visual-art-director

本文件属开发治理产物，**严禁带入日常 PPT 设计任务的运行上下文**。所有结论严格区分四类证据：
**代码实测**（可重跑的命令与输出）、**渲染观察**（看图得到的结论）、**设计判断**（需人作取舍）、
**未验证**（条件缺失，不得写作通过）。

---

## 1. 审计基线与环境边界

| 项 | 事实 |
| --- | --- |
| 审计起点 | `main @ 5db2573`（12 个受管文件 / 2,376 行），即上一轮交付记录所称的 `8cb85a8` 之后的最新提交 |
| 运行环境 | Python `3.13.16`、`python-pptx 1.0.2`、`Pillow 12.3.0`、`pytest 9.0.3` |
| 渲染链 | Linux 无头 `LibreOffice Impress 25.2.3.2` + `poppler-utils (pdftoppm/pdftotext/pdfinfo/pdffonts)` |
| 中文字体 | `Noto Sans CJK SC` / `Noto Serif CJK SC`（仓库无内嵌字体，运行时依赖目标机字体） |
| 测试规模 | 优化前 **29** 项 → 优化后 **69** 项（`test_compile` 53 + `test_render` 13 + `test_install` 3） |

**客观证据缺口（本轮仍未闭合）**

1. 上游无任何历史 `.pptx` 基线产物，无法做跨版本盲评的"视觉非退化"证明。
2. 无 Windows/macOS 原生 **Microsoft PowerPoint**（微软雅黑/宋体真实字距与会议室投影）终验。
3. LibreOffice 的字体回落与行高与 PowerPoint 不完全一致：渲染回归只能证明"在该 Linux 渲染链下成立"。

---

## 2. 缺陷清单（本轮新发现，按风险排序）

### P1-A｜E008 语义核对基本失效（4 条真实缺陷，全部已复现）

**复现命令**：`python3 /tmp/probe_e008.py`（7 类陷阱，优化前 7/7 全部"PASS"即全部漏检）。

| # | 缺陷 | 优化前行为（实测） | 根因 |
| --- | --- | --- | --- |
| 1 | **增长/下降方向混淆** | 「获客成本下降 25%」对上升序列 800→900→1000 **不报警** | `_chart_verifiable_numbers()` 只收集数值集合，`abs()` 抹掉了方向；
| 2 | **任一数字匹配即整句通过** | 「Q2 获客 900 人，同时获客成本下降 30%」**不报警** | `any(...)` 短路：900 命中即放弃核对 30% |
| 3 | **单位与统计期间错配** | 「本季度收入 1.2 亿元」对 `unit=万元, value=1.2` **不报警**；「Q4 获客 1000 人」对 Q1–Q3 序列 **不报警** | 解析只取裸数值，单位/期间从未进入比较；`unit`/`period` 字段只用于标题 |
| 4 | **派生值无依据膨胀** | 「规模达到 80,000 人」因 `v×100` 候选值命中而**通过** | 为"容错"加入的 `v*100`/`abs(v)*100` 候选，使任意数字都可能碰巧匹配 |

**修复方案（`scripts/compile_pptx.py` 重构语义层，约 600 行）**

- `Quantity` 数据类显式区分 `kind`(level/percent/pct_points)、`direction`(±1/0)、`unit`、`period`、
  `origin`(raw/total/share/diff/change/cell)、`comparative`(同比/环比)、`pair_role`(从…到…)。
- 主张侧：逐条数值独立解析，保留符号、百分号、百分点、量级（万/亿/千）、小句方向动词与期间。
- 数据侧：原始值、系列合计、多系列同类别合计、饼/环切片占比、相邻/首末差值、变化率**分别建模并带出处**。
- 匹配：先按 kind → 单位量纲 → 方向 过滤，再比较数值；`12000 万元 == 1.2 亿元` 通过，`1.2 万元 vs 1.2 亿元` 报量纲错配。
- 期间：`Q3 / 26Q3 / 2026Q3 / H1 / 第X季度 / X月` 交叉核对；命中值仍需与主张期间自洽；图注/来源期间与
  全部图表期间无交集时提示。
- 输出：整页聚合为**一条** `E008`，逐项列出「可核对 / 待核」与原因（任务指令要求逐项列出可观察指标）。
- **取舍**：`E008` 仍为 `SOFT`。理由：它是数值一致性筛查而非事实核验，同值不同指标（毛利率 vs 营收增速）、
  因果方向、口径定义合理性无法由程序判定；把它升为 HARD 会诱使模型为过编译而改写事实，违背"设计判断优先"。

**回归**：新增 13 项 E008 测试（方向、单数字救句、单位换算/错配、期间错配与跨期、派生值污染、
分数尺度歧义、百分点差、从…到…顺序、图注期间、整页聚合、不阻断编译、裸量级单位）。

### P1-B｜真实渲染验证缺失（已建立）

- **优化前**：`dev/DELIVERY.md` 声称"7 张真实渲染 PNG 目视复核"，但**仓库内没有任何固定案例、
  渲染脚本或可重跑的渲染测试**；任何后续改动都无法复现该结论。渲染链缺失时也没有任何机制提示"未验证"。
- **优化后**：`dev/tests/fixtures/render_deck.json`（4 页固定案例：深底封面 / 原生柱状图 / 无竖框三线表 /
  16:9→竖框裁切）+ `dev/tests/test_render.py`（13 项）：
  页数与 PDF 页数一致、页面非空白、深底真的是深底、CJK 字体在 PDF 中**嵌入**、每个文本 run 有 `<a:ea>`、
  逐个文本元素**墨迹包围盒不越出声明框**（容差 4pt）、文字不越出页面、**避头尾**（行首/行尾禁则字符）、
  **数字与单位不跨行**、图片裁切后有效比例误差 <1%、柱状图/表格必须是原生对象、表格无竖线且数值右列对齐。
- **环境缺失处置**：渲染链缺失时整模块显式 `skip`，跳过即意味着 `UNVERIFIED`，**不会**被读成 PASS。

### P1-C｜`E010`：文本装不下时无人提示（新增，SOFT）

- **实测**：一张 `32pt` 中文标题写进 `6.0×1.0 in` 的框，LibreOffice 渲染后**溢出框底**（
  `pdftotext -bbox` 实测 `yMax=129.2pt` vs 框底 `126pt`）。仓库原有检查无任何相关提示。
- **处置**：新增 `estimate_text_lines()` + `text_fit_risk()`，以"字宽近似 + 显式换行"估算所需高度，
  仅当估算高度 > 框高 × 1.05 时给 `E010 SOFT`。
- **取舍**：**不自动缩字号、不改写文案**（有测试 `test_e010_does_not_shrink_font_or_rewrite_content` 锁定）。
  估算模型的误差在 4pt 容差内已被接受，真正结论以渲染墨迹包围盒为准。

### P2-A｜图表配色契约未落地（已复现）

| 缺陷 | 实测 | 根因 |
| --- | --- | --- |
| 饼/环图只给 `series[0].format` 上色 | 三分类饼图渲染成**单一颜色** | `if palette and kind in {"bar","line"}` 直接排除了 pie/donut |
| 饼图无任何标注 | 类别只能靠颜色区分，违反"复杂状态图表不得单靠颜色编码" | 未设置 data labels |
| 未设 `palette` 时静默回退 Office 默认蓝 | 与 AI 确定的视觉语法冲突且无提示 | 无回退策略 |
| 同色调色板不报 | `palette[0]` 与页面底色同色时系列消失 | 只按相对亮度粗筛，无对比度门槛、无披露 |

**修复**：`_pick_series_colors()` 按契约顺序消费调色板，`bar/line` 逐系列、`pie/donut` **逐扇区**；
线/标记按 `≥3:1`、面按 `≥2:1` 过滤；被跳过的颜色产生 `E009 SOFT` 如实披露（"发生了换色"），不静默替换；
饼/环显示**仅百分比**标注（修复了渲染中出现的 `58; 58%` 重复），标注墨色按"最差扇区也可读"选择
（`_label_ink`）；单系列柱/线图省略与标题重复的图例，饼/环保留并置底。

**仍存在（未修）**：未设 `palette` 时仍回退到 python-pptx 主题默认蓝。已修的是"有契约却不执行"，
"无调色板时如何配色"属于设计判断，交由 AI 在 `plan.palette` 中决定；本轮不擅自发明默认色板。

### P2-B｜评测不可复现（已修复）

- **优化前**：`eval_cases.json` 有 8 条 case，但无 `input_facts` 固定口径、无 `vetoes`（一票否决）、
  无盲评与渲染环境约定；`rubric.md` 的 `PASS` 条件与"必须有旧版对照"自相矛盾（无基线即永远无法 PASS，
  但没有说明如何补齐）。断言与设计判断混在一起，无法区分"机器检查通过"与"设计合格"。
- **优化后**：`schema_version 2.0` + `protocol`（固定输入 / 相近模型条件 / 一致渲染环境 / 去标识保留产物 /
  盲评 / 自评不得作为唯一依据）+ 每条 case 的 `input_facts`、`observable_criteria`、`vetoes`、
  `human_review_required`；`rubric.md` 新增"各维度取证方式"表，明确机器检查不折算为设计分。

### P3｜安装与 CI（已复现并修复）

| 缺陷 | 实测证据 | 修复 |
| --- | --- | --- |
| `pip install .` 直接失败 | `error: Multiple top-level packages discovered in a flat-layout: ['dev', 'references']` | 补 `[build-system]`，`[tool.setuptools] packages = ["scripts"]` |
| 无安装测试 | — | `dev/tests/test_install.py`（3 项）：构建契约、入口可调用、端到端 `pip install --target` + 跑 CLI 编译 |
| 无 CI | — | `.github/workflows/ci.yml`：`unit`（3.10/3.13，纯 Python）+ `render`（装 LibreOffice/poppler/CJK 字体跑渲染回归并上传产物） |

`pip install --target` 实测成功，`ppt-vad plan.json out.pptx` 退出码 0、产物 28,438 B。
`pyproject` 的 `requires-python = ">=3.10"` 已用 `ast.parse(feature_version=(3, 10))` 逐文件核对语法兼容。

---

## 3. 修改文件与关键变更

| 路径 | 变更 | 判定依据 |
| --- | --- | --- |
| `scripts/compile_pptx.py` | +E008 语义层（重写）、+E009、+E010、图表调色板/标注/图例、E007 根因透传、删除死代码 | 上述 P1-A/P1-C/P2-A 复现 |
| `references/contract.md` | 补 `E008` 核对范围与边界、`E009`、`E010`、新增"图表呈现契约"表 | 契约必须与实现一致 |
| `SKILL.md` | 补充 `E009` 说明、明确 `E008` 是筛查而非核验、VERIFY 阶段"必须真的看图" | 防止把程序结果当设计结论 |
| `dev/tests/test_compile.py` | 29 → 53 项 | P1-A/P1-C/P2-A/P3 |
| `dev/tests/test_render.py` | 新增 13 项 | P1-B |
| `dev/tests/fixtures/render_deck.json` | 新增固定案例 | P1-B |
| `dev/tests/test_install.py` | 新增 3 项 | P3 |
| `dev/evals/eval_cases.json` | 2.0 schema、protocol、input_facts/vetoes/human_review | P2-B |
| `dev/evals/rubric.md` | 取证方式表、盲评与结论口径 | P2-B |
| `dev/evals/baseline/render/*.png` | 新增去标识基准渲染产物（5 张）+ 复现说明 | P2-B 盲评需要可比对的历史产物 |
| `pyproject.toml` | build-system / packages / optional-dependencies | P3 |
| `.github/workflows/ci.yml` | 新增 | P3 |
| `dev/DELIVERY.md` | 本文件重写 | 交付一致性 |

**新增基准产物**：`dev/evals/baseline/render/`（5 张 110dpi PNG + `README.md`）。文件名不含版本/模型信息，
可供下轮盲评直接对照；这是本轮唯一新增的二进制产物（284 KB），不进入运行时上下文。

**未改动**：`references/art-direction.md`、`color-art-direction.md`、`image-prompt.md`。
本轮审计未在其中发现实际失效引用或与实现冲突的条文；按"最小必要改动"原则不动。

---

## 4. 验证证据

### 4.1 代码实测（可重跑）

```bash
python -m pytest dev/tests                    # 69 passed
python -m pytest dev/tests/test_render.py     # 13 passed（需 LibreOffice + poppler + CJK 字体）
python -m pytest dev/tests/test_install.py    # 3 passed（需 pip）
python -m pyflakes scripts/compile_pptx.py dev/tests/*.py   # 无输出
```

| 检查 | 结果 |
| --- | --- |
| 全量测试 | **69 passed in 3.1s**（优化前 29） |
| `pip install --target` + CLI | 成功，`ppt-vad` 退出码 0，产物非空 |
| 固定案例 `--strict` | 零错误零告警 |
| `pyflakes` | 无告警（并修掉了一处真实 bug：`_add_image` 引用了不存在的 `unusable`） |

### 4.2 渲染观察（LibreOffice 25.2 → PDF → 110dpi PNG，逐页看图）

| 页面 | 观察到的现象 | 处置 |
| --- | --- | --- |
| 1 深底封面 | 衬线标题断面清晰、发丝线细且不断、页脚可读 | 通过 |
| 2 柱状图 | 首次渲染：图注写 `2026 Q1–Q3` 但图表期间为 `2025 Q2–2026 Q3`（**我自己案例的口径错误**）；6 根细方块柱之间留白不足 | 修正图注；改为 6 个季度连续序列使"连续增长"有据；并把"图注期间与图表期间无交集"补进 `E008` |
| 2 柱状图（修后） | 6 类目宽柱间距正常、无图例重复、网格线为浅发丝 | 通过 |
| 3 三线表 | 无竖线、无斑马纹、表头上下粗线、数值列右对齐、单位随数值 | 通过 |
| 4 图文页 | 源图 1600×900 裁切为 5.03×7.5 竖框，实测 `crop_left=right=0.3113`，圆形主体仍为正圆 | 通过 |
| 饼/环验证页 | 首轮：所有扇区**同色**、标注显示 `58; 58%`、深蓝扇区上深色字不可读 | 逐扇区着色 + 仅百分比 + 按最差对比度选墨色，三轮后通过 |

**结论**：自动检查全部通过**之后**，看图仍发现 3 类真实缺陷（口径不一致、构图失衡、扇区对比度）。
这正是"编译成功 ≠ 视觉合格"的实证；也说明机器检查必须与目视并存。

### 4.3 未覆盖 / 仍需人工终验

| 项 | 状态 | 说明 |
| --- | --- | --- |
| 微软雅黑/宋体在真实 PowerPoint 的字距与断行 | `UNVERIFIED` | 本环境无 Microsoft PowerPoint，Noto CJK 与雅黑度量不同 |
| 会议室投影下的对比度与字号 | `UNVERIFIED` | 无投影设备；仅按 WCAG 相对亮度做了客观计算 |
| 历史版本视觉非退化 | `UNVERIFIED` | 无历史 `.pptx` 基线 |
| 8 条 eval case 的实际跑分 | `UNVERIFIED` | 本轮只建立协议与客观指标，未做模型盲评（需要多模型/多版本对照） |
| `E008` 同值不同指标、因果方向、口径定义 | 程序不覆盖 | 属 AI/人工判断；已在契约与 `SKILL.md` 显式声明 |
| 无 `palette` 时的默认配色 | 未改 | 回退 python-pptx 主题蓝；属设计判断，交由 AI 指定 |
| 4:3 画幅的渲染回归 | 未覆盖测试 | 几何校验有单测，但固定渲染案例只有 16:9 |
| JPEG/PNG/WebP 混合素材与 EXIF 旋转 | 未覆盖 | 图片裁切只测了 PNG 正立图 |

---

## 5. 优化前后差异

| 维度 | 优化前 | 优化后 |
| --- | --- | --- |
| E008 能发现的问题 | 数值完全对不上（且被派生值污染大量漏检） | 方向反转、单位量级错配、期间不符、无来源数字、裸量级换算、尺度歧义、图注期间矛盾 |
| E008 判定粒度 | 整句 `any()` | 逐条定量表述独立判定并逐项报告 |
| 渲染验证 | 无固定案例、无脚本、结论不可复现 | 固定案例 + 13 项渲染回归 + 产物可上传 |
| 图表色彩 | 饼/环单色、无标注、静默回退 | 逐扇区着色、仅百分比标注、对比度筛选 + `E009` 披露 |
| 文本溢出 | 无提示 | `E010` 估算提示（不缩字号、不改文案） |
| 安装 | `pip install .` 失败 | 可构建、可安装、CLI 可用，有安装测试 |
| CI | 无 | unit（3.10/3.13）+ render 两条流水线 |
| 评测 | 断言混杂、无盲评协议 | schema 2.0、固定输入、一票否决、盲评与产物保留 |

**行数变化**：`scripts/compile_pptx.py` 1,115 → 1,852；测试 522 → 1,258（含固定案例与两个新测试文件）；
运行时文档 `SKILL.md + references/*.md` 488 → 517（只补必要的风险声明，未扩写规则清单）。

---

## 6. 发布状态

**`UNVERIFIED`**

- **不判 `BLOCK` 的依据**：本轮复现的 8 类缺陷（E008 四项、渲染验证缺失、E010、图表色彩、安装）均已修复，
  69 项测试通过，4 页固定案例 `--strict` 零告警并通过渲染目视复核。
- **不判 `PASS` 的依据**：缺少历史 `.pptx` 基线做盲评非退化对照；缺少目标 PowerPoint 与真实投影环境的
  人工终验；8 条评测案例尚未做多版本盲评打分。
- **判 `PASS` 的前置条件**：① 至少一次目标 PowerPoint 逐页核验（数据口径 / 中文断行 / 投影可读性）；
  ② 用本次 `eval_cases.json` 协议跑一轮去标识盲评并归档产物；③ 补 4:3 与混合素材的渲染用例。
