# ppt-visual-art-director

以国际一线视觉艺术总监的判断制作、改稿或评审 PPT。AI 负责设计判断，脚本负责确定性执行与事实核验，真实渲染与人眼复核共同把关，目标环境由人最终签收。

交付原生可编辑的 `.pptx`，并用对象核对与真实渲染给出可验证证据。机器检查不等于好看，LibreOffice 预览不等于 Microsoft PowerPoint 目标环境通过。

## 目录结构

```
LICENSE                  MIT 开源许可证
README.md                安装、命令、验证分级与能力边界说明
SKILL.md                 判断原则、判断顺序、brief、任务路线、验收与边界（AI 主读文件）
requirements.txt         Python 依赖声明
references/
  design.md              概念、艺术指导、叙事、空间、信息设计、排版、色彩、图像、诊断、复核维度
  build.md               plan 语义、几何、文本与 runs、素材与生成图、图表、命令、诊断码、定点编辑、兼容性清单
scripts/
  core.py                共享契约、度量、几何、色彩、文本样式、设计信号与包校验
  pdf_backend.py         PDF 解析与 PNG 渲染适配器（PyMuPDF 优先，自动回退 Poppler）
  compile.py             plan.json → 原生 PPTX（校验、编译、成品包校验）
  verify.py              plan ↔ PPTX ↔ 真实渲染（对象核对、逐元素文字归属、字体替换检查、状态判定）
  revise.py              既有 PPTX：inspect 只读解析；定点编辑 + 保留核对（富文本、超链接、OOXML 结构）
  plan.schema.json       plan 的 JSON Schema 数据契约
  limits.json            技术阈值与默认值的唯一来源
tests/                   工程回归用例（不属于运行时必需，用于开发与 CI 回归）
```

## 安装与环境要求

- **Python**: 3.9+
- **Python 依赖**: `pip install -r requirements.txt`（`python-pptx`、`jsonschema`、`Pillow`、`lxml`）
- **渲染核验环境**: LibreOffice（`soffice`），配合以下任一 PDF 后端：
  - `pymupdf`（`pip install pymupdf`），或
  - 系统 Poppler 工具链（`pdftoppm`、`pdftotext`、`pdffonts`、`pdfinfo`）

缺少渲染工具时，`verify.py` 不会伪造通过，而是如实降级为 `PARTIAL` 并说明未核验项。

## 常用命令

```bash
python scripts/compile.py plan.json out.pptx          # 新建：校验 + 编译
python scripts/verify.py plan.json out.pptx           # 新建稿核验，写 render/<名称>/
python scripts/revise.py inspect 原稿.pptx             # 改稿／评审：只读解析
python scripts/revise.py 原稿.pptx edit.json out.pptx # 改稿：定点编辑（保留核对，原稿不动）
python -m unittest discover -s tests -v               # 回归用例
```

## 验证分级与状态诚实度

- **Tier 1（对象核对）**：校验 PPTX 内部每个对象的几何、文本、run 样式、颜色、图表载荷、表格单元格、图片字节与包结构。
- **Tier 2（LibreOffice 渲染核对）**：通过 LibreOffice → PDF → PNG 核查页尺寸、越界、页码、空白页、字体替换与逐元素文字归属。
- **Tier 3（目标 Microsoft PowerPoint 兼容）**：在 Windows / macOS 真实 PowerPoint 中核查字体、渐变、双色调、裁切与溢出。

状态三档：`BLOCK`（存在硬错误）、`PARTIAL`（渲染未跑或声明字体被替换）、`READY`（对象一致、LibreOffice 渲染完成且声明字体均参与渲染）。**任何自动状态都不代表“设计通过”或“PowerPoint 实机已验证”**——未在目标 PowerPoint 实际打开即标记为 `NOT VERIFIED`。

## 能力边界

- **完整支持**：`16:9` / `4:3` 原生 PPTX、富文本与 runs、矢量形状/渐变/路径/组合、原生图表（柱/条/堆叠/组合/折线/面积/饼/环/散点/表格）、既有稿定点编辑与 OOXML 保留核对。
- **有限支持**：`duotone`（PowerPoint 生效，LibreOffice 仅去饱和）、无障碍读屏（页标题为普通文本框，非占位符）、曲线/复杂关系（仅直线与折线，编译期锚点）。
- **外部依赖 / 不支持**：图片生成采用工具无关的简报与提示词工作流并强制登记 SHA-256；新建模式不生成备注、动画、页面切换、双数值轴、桑基、甘特或雷达图，不嵌入字体。

## 未验证与已知限制

- 在无 Windows / Microsoft PowerPoint 的 Linux 环境下，Tier 3（PowerPoint 实机渲染、Windows 专有字体如微软雅黑的精确折行、原生 `duotone` 着色效果）均标记为 `NOT VERIFIED`。
- 详见 `references/build.md` 中的「Microsoft PowerPoint 真实环境兼容核对清单」。
