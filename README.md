# ppt-visual-art-director

让 AI 以**视觉艺术指导**的方式制作、改稿和评审 PPT/PPTX：先由受众、主张与证据推导整稿概念、构图、中文排版、色彩、素材与叙事节奏，再用脚本确定性编译为**原生可编辑**的 PPTX，并用真实渲染与具名人工验收核验。

> 一句话原则：设计判断归 AI，确定性执行归脚本，审美结论归人。机器检查不等于好看，预览不等于目标环境通过。

## 包结构

```
SKILL.md                     核心流程（THINK → BUILD → VERIFY）与按需加载表；技能入口
references/
  art-direction.md           整稿概念、节奏、构图、中文排版、素材、关系图形、诊断、模板味信号
  color-art-direction.md     颜色职责、明度结构、数据色、深浅场切换
  image-prompt.md            图像简报、提示词、真实调用闭环、三级评审
  layout-system.md           坐标单位、网格、字号尺度、文本框尺寸估算、页码保留区
  review-protocol.md         任务模式、改稿路线、评审报告、AI 预审 vs 人工签收、无渲染工具时的降级
  contract.md                设计判断与机器执行的边界、诊断码、质量门
scripts/
  compile_pptx.py            plan.json → 原生 PPTX；写文件前做契约检查、数据语义筛查、装载估算
  verify_delivery.py         plan ↔ PPTX ↔ 真实渲染 ↔ 质量门，输出证据报告
  inspect_pptx.py            只读解析既有 PPTX（改稿 / 评审第一步）
  grid.py                    网格 → 英寸坐标，标出页码保留区
  plan.schema.json           plan 的唯一数据契约
  policy.json                技术阈值、默认值、诊断路由、门与发布状态
examples/                    minimal_plan.json、showcase_plan.json、acceptance.json、assets 生成脚本
tests/                       回归测试（无需 LibreOffice）
CHANGELOG.md                 变更历史
```

## 快速开始

```bash
python -m venv .venv && . .venv/bin/activate         # Windows：.venv\Scripts\activate
pip install -r requirements.txt

# 1) 跑通示例（严格模式：未具名处置的 SOFT 问题视为失败）
python scripts/compile_pptx.py examples/showcase_plan.json out.pptx --assets examples/assets --strict

# 2) 验证（有 LibreOffice + poppler 时会真实渲染；渲染目录必须为空）
mkdir render
python scripts/verify_delivery.py examples/showcase_plan.json out.pptx \
  --assets examples/assets --render-dir render --report report.json

# 3) 解析一份已有 PPTX（只读）
python scripts/inspect_pptx.py 原稿.pptx

# 4) 把网格换算成 plan 用的英寸坐标
python scripts/grid.py --span 1:5 --span 7:6

# 5) 回归测试
python -m unittest discover -s tests -v
```

## 重要事实

- **坐标单位是英寸**，原点在页面左上角；16:9 画布 13.333 × 7.5 in，4:3 为 10 × 7.5 in；`font_size` 为 pt。
- **编译器总是生成新文件，不就地修改原稿。** 原稿含自定义母版、动画、SmartArt 等时，改稿路线见 `references/review-protocol.md` §2。
- **默认交付状态常为 `BLOCK — Delivery`**：`G-DECK`（整稿设计）与 `G-TGT`（目标 PowerPoint）是具名人工门。AI 可以预审，但不能自填 reviewer；签收步骤见 `references/review-protocol.md` §4 与 `examples/acceptance.json`。
- **渲染门依赖外部工具**：LibreOffice Impress、poppler（`pdftoppm`/`pdftotext`/`pdffonts`）、fontconfig。缺失时 `G-PRE` 如实判 `NOT VERIFIED`，安装指引见 `references/review-protocol.md` §5。
- 默认字体“微软雅黑”是 Windows 字体；跨平台放映请按 `references/art-direction.md` §4 选择并实测。编译器只写字体名，不嵌入字体文件。
- 示例中的数据均为**演示数据**，不是真实业务数据。

## 编译器现在不能做什么

不生成演讲者备注（plan 的页 `notes` 只是设计备注）、动画与页面切换。没有 XY 散点、桑基、自动甘特、决策树和绑定端点的连接器；专用原生图表只有柱、线、饼、环与表格。其余关系可用文字、基本形状和分隔线手工组合，位置不与数据绑定。页标题是普通文本框而非标题占位符，因此 PowerPoint 的大纲视图与读屏软件识别不到页标题——对无障碍有要求时需要在交付说明中提示或后续手工补充。详见 `references/contract.md` §3。

## 许可

仓库未声明许可证。复用或再分发前请先向作者确认；作者可在根目录补充 `LICENSE`。
