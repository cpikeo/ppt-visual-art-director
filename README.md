# ppt-visual-art-director

**版本 v1.0.0**（契约版本：Schema `4.0`、Policy `2.5`、报告 `1.5`，各自独立，与包版本不联动）

让 AI 以**视觉艺术指导**的方式制作、改稿和评审 PPT/PPTX：先理解受众、目标、主张、证据与洞察，推导整稿视觉概念、叙事节奏、构图、中文排版、色彩与素材；再用脚本确定性编译为**原生可编辑**的 PPTX；最后以真实渲染、看图复核、根因定点修正与具名人工验收闭环。

> 一句话原则：设计判断归 AI，确定性执行归脚本，审美与签收结论归人。机器检查不等于好看，预览不等于目标环境通过。

它不是 PPT 模板库、自动排版工具或简单的生成脚本：不从行业关键词推导风格，设计由内容、受众、传播目的与视觉概念共同决定。

## 设计决策链

Audience → Objective → Message → Evidence → Insight → Visual Concept → Narrative → Composition → Art Direction → Execution → Verification。每一步及其落入的 plan 字段见 `SKILL.md`；视觉概念的推导与自检见 `references/art-direction.md` §1.1；十一个评审维度（含概念成立、整稿成立、设计效率）见同文件 §7；看图复核、根因修正与停止原则见 `references/review-protocol.md` §4。

## 包结构

```
SKILL.md                     技能入口：设计决策链、原则、BUILD / VERIFY 流程与按需加载表
README.md                    本文件
requirements.txt             直接依赖：python-pptx、jsonschema、Pillow
references/
  art-direction.md           概念推导、整稿识别与节奏、构图、中文排版、素材、关系图形、十一维度评审、模板味信号
  color-art-direction.md     颜色职责、明度结构、数据色、深浅场切换
  image-prompt.md            图像简报、提示词、真实调用闭环、三级评审
  layout-system.md           坐标单位、网格、字号尺度、文本框尺寸估算、页码保留区
  review-protocol.md         任务模式、改稿路线、评审报告、看图复核与根因修正、AI 预审 vs 人工签收、降级做法
  contract.md                设计判断与机器执行的边界、plan 最小骨架、诊断码、质量门、核验独立性
scripts/
  compile_pptx.py            plan.json → 原生 PPTX；写文件前做契约检查、数据语义筛查、装载估算
  verify_delivery.py         plan ↔ PPTX ↔ 真实渲染 ↔ 质量门，输出证据报告
  inspect_pptx.py            只读解析既有 PPTX（改稿 / 评审第一步）
  grid.py                    网格 → 英寸坐标，标出页码保留区
  plan.schema.json           plan 与验收配置的唯一数据契约
  policy.json                技术阈值、默认值、诊断路由、门与发布状态的唯一来源
```

本发布包不含示例、测试与变更历史。plan 的最小骨架在 `references/contract.md` §2.1；人工验收配置的形状以 `scripts/plan.schema.json` 的 `$defs.acceptanceConfig` 为准，骨架见 `references/review-protocol.md` §5。

## 安装

```bash
python -m venv .venv && . .venv/bin/activate         # Windows：.venv\Scripts\activate
pip install -r requirements.txt
```

Python 3.9+。脚本不会联网。真实预览另需 LibreOffice Impress、poppler（`pdftoppm` / `pdftotext` / `pdffonts`）、fontconfig 与相应字体；安装参考见 `references/review-protocol.md` §6。

## 命令

```bash
# 1) 把 references/contract.md §2.1 的骨架保存为 plan.json，严格编译（未具名处置的 SOFT 问题视为失败）
python scripts/compile_pptx.py plan.json out.pptx --assets . --strict

# 2) 验证（有 LibreOffice + poppler 时真实渲染；渲染目录必须为空；--assets 请与编译时一致）
mkdir render
python scripts/verify_delivery.py plan.json out.pptx --assets . --render-dir render --report report.json
#   可选 --config 验收配置.json

# 3) 解析一份已有 PPTX（只读）；--json 另存完整结果
python scripts/inspect_pptx.py 原稿.pptx [--json inspect.json]

# 4) 把网格换算成 plan 用的英寸坐标
python scripts/grid.py --span 1:5 --span 7:6 [--aspect 4:3] [--cols 12] [--corner top_right] [--json]
```

各脚本均可用 `--help` 查看参数（`compile_pptx.py`、`verify_delivery.py`、`inspect_pptx.py` 需先安装依赖）。

## 重要事实

- **坐标单位是英寸**，原点在页面左上角；16:9 画布 13.333 × 7.5 in，4:3 为 10 × 7.5 in；`font_size` 为 pt。
- **编译器总是生成新文件，不就地修改原稿。** 原稿含自定义母版、动画、SmartArt 等时，改稿路线见 `references/review-protocol.md` §2。
- **默认交付状态常为 `BLOCK — Delivery`**：`G-DECK`（整稿设计）与 `G-TGT`（目标 PowerPoint）是具名人工门。AI 可以预审，但不能自填 reviewer；签收步骤见 `references/review-protocol.md` §5。
- **视觉质量只在“真实渲染 + 能查看图片”时才可能被判定。** 缺任一条件，相关维度如实记 `NOT VERIFIED`。
- **渲染门依赖外部工具**：缺失时 `G-PRE` 如实判 `NOT VERIFIED`。
- 默认字体“微软雅黑”是 Windows 字体；跨平台放映请按 `references/art-direction.md` §4 选择并实测。编译器只写字体名，不嵌入字体文件。

## 编译器现在不能做什么

不生成演讲者备注（plan 的页 `notes` 只是设计备注）、动画与页面切换。没有 XY 散点、桑基、自动甘特、决策树和绑定端点的连接器；专用原生图表只有柱、线、饼、环与表格。其余关系可用文字、基本形状和分隔线手工组合，位置不与数据绑定。页标题是普通文本框而非标题占位符，因此 PowerPoint 的大纲视图与读屏软件识别不到页标题——对无障碍有要求时需要在交付说明中提示或后续手工补充。详见 `references/contract.md` §3。

## 许可

仓库未声明许可证。复用或再分发前请先向作者确认；作者可在根目录补充 `LICENSE`。
