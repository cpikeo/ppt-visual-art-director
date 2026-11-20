# ppt-visual-art-director

以视觉艺术总监的判断制作、改稿或评审 PPT。**AI 负责设计判断，脚本负责确定性执行与事实核验，真实作品负责证明质量，人负责最终签收。**

交付原生可编辑的 PPTX（只输出 `.pptx`），并用对象核对与真实渲染给出证据。机器检查不等于好看，预览不等于目标环境通过。

## 仓库结构

```
技能包（交付给 AI，只含运行必需文件）
  SKILL.md               判断原则、判断顺序、三种整体判断、任务路线、验收、交付
  references/design.md   设计师方法、概念、艺术指导、叙事、空间、信息设计、排印、色彩、诊断、复核
  references/imagery.md  图片与视觉资产：资产职责、简报、提示词、验收
  references/build.md    plan 语义、几何、素材、图表、命令、诊断码、定点编辑、输入安全、能力边界
  scripts/               model · ooxml · charts · checks · compile · verify · revise
                         plan.schema.json（数据契约，唯一）  limits.json（技术阈值，唯一）
  requirements.txt       必需依赖
  requirements-render.txt  可选：渲染核验（PyMuPDF 为 AGPL-3.0）

开发材料（不属于技能包；打包给 AI 时排除整个 dev/）
  dev/tests/             回归测试（12 项：编译、安全、原生结构、定点改稿、灰度信号）
  dev/evals/             art-direction-demo（可复现的 6 页示例稿）与 FINDINGS.md
  dev/compat/COMPAT.md   已验证与未验证清单
```

## 本次修订做了什么

- **设计判断**：新增“三种整体判断”（整体性、连续性、有效性）作为看图必做项；新增“向大师学什么”（只提炼判断方法）；色彩改为“从事实推导的关系系统”（环境/结构/强调/数据四种职责，区分各类中性色），去掉预设色板起点；数据视觉补上四层分工与“不夸大”。
- **工程**：页标题编译为**原生标题占位符**（大纲/读屏/导航可识别），无标题页不留空占位符；新增 `W-GRAY`（多系列只靠色相区分时报警）；图表标注默认 14pt。
- **精简**：SKILL.md 去掉与 references 重复的段落；design.md 删除重复的图像节，评测经验由 12 条散记改为 16 条可执行的判断。
- **证据**：补齐 README 一直声明、但仓库里缺失的 `dev/`（测试、评测样本、兼容记录）。

## 安装与命令

- Python 3.9+（在 3.11 上测试）。必需：`pip install -r requirements.txt`。
- 可选渲染核验：`pip install -r requirements-render.txt` + LibreOffice（`soffice` 在 PATH，或设 `SOFFICE`）。不装时 `verify.py` 给出 `PARTIAL`。
- 字体只写名称、不嵌入。脚本不联网。

```bash
python scripts/compile.py plan.json out.pptx [--assets DIR]
python scripts/verify.py plan.json out.pptx [--assets DIR]
python scripts/revise.py inspect 原稿.pptx
python scripts/revise.py 原稿.pptx edit.json out.pptx
python -m unittest discover -s dev/tests -v          # 开发回归
python dev/evals/art-direction-demo/build_plan.py     # 重建示例 plan
```

## 能力边界

| 级别 | 内容 |
| --- | --- |
| 完整支持 | 16:9 / 4:3；原生文本、形状、折线路径、组、图片裁切；原生图表（柱、线、面积、饼、环、散点）与表格；原生页标题；备注；替代文字；页码；定点改稿 |
| 有限支持 | 无双数值轴、桑基、甘特；表格无行高亮；无逐系列线宽；渐变与 duotone 在 LibreOffice 只是近似；连线是坐标而非连接符 |
| 外部依赖 | 图片生成（工具无关，只登记）；LibreOffice；PowerPoint 实机；目标机字体 |
| 不支持 | 动画、切换、嵌入字体、新建 SmartArt、宏 |

## 验证等级

| 等级 | 内容 | 状态 |
| --- | --- | --- |
| L0 对象 | plan ↔ PPTX 结构与对象核对 | 已做（示例稿 + 12 项回归） |
| L1 LibreOffice | 无头渲染、断行、字体、越界 | 已做：7.4.7.2 / Debian 12，示例稿 READY |
| L2 PowerPoint 实机 | 真实 Microsoft PowerPoint | **NOT VERIFIED** |
| L3 人工签收 | 目标环境下人确认 | **NOT REVIEWED** |

示例稿的评价是作者自评，数据为虚构；详见 `dev/evals/FINDINGS.md`（含没做好的部分）。

## 许可证

本仓库目前没有 LICENSE 文件，未声明开源许可证；选择哪种由作者决定。第三方依赖：python-pptx（MIT）、jsonschema（MIT）、Pillow（HPND）、lxml（BSD）；可选的 PyMuPDF 为 AGPL-3.0 / 商业双许可。示例稿里的生成图权利状态为 `pending`，只作评测素材。
