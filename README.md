# PPT Visual Art Director

视觉艺术总监式的 PPT 工作流：先判断（受众、洞察、视觉概念、色彩方向、叙事），再把判断写成 `plan.json`，由脚本确定性地编译成原生可编辑的 PPTX，并核验、改稿。

## 文件

| 路径 | 作用 |
| --- | --- |
| `SKILL.md` | 身份、判断原则、执行顺序、任务路由 |
| `references/design.md` | 概念、艺术指导、叙事、空间、信息设计、排版、色彩、诊断 |
| `references/imagery.md` | 摄影、背景图、插图、图标、生成素材与验收 |
| `references/build.md` | plan 语义、图表、诊断码、定点改稿、能力边界 |
| `scripts/plan.schema.json` | 数据结构唯一契约 |
| `scripts/limits.json` | 技术阈值与默认值唯一来源 |
| `scripts/{model,checks,compile,charts,chart_audit,ooxml,verify,revise}.py` | 共用模型 / plan 校验 / 编译 / 图表 / 成品图表审计 / OOXML 写入 / 核验 / 改稿 |

## 命令

```bash
pip install -r requirements.txt            # 编译、核验（对象层）、改稿
pip install -r requirements-render.txt     # 可选：渲染核验（PyMuPDF 为 AGPL-3.0；另需 LibreOffice，PATH 或 SOFFICE）
python scripts/compile.py plan.json out.pptx [--assets DIR] [--report r.json] [--force]
python scripts/verify.py plan.json out.pptx [--assets DIR] [--render-dir DIR] [--report v.json] [--force]
python scripts/revise.py inspect 原稿.pptx [--json i.json] [--force]
python scripts/revise.py 原稿.pptx edit.json out.pptx [--assets DIR] [--report r.json] [--force]
```

## 验证边界

脚本只做对象核验与（可选的）LibreOffice 渲染核验；构图、主次、概念是否成立，以及真实 PowerPoint 下的表现，必须由人判断。没有渲染就是 `PARTIAL`，字体被替换也是 `PARTIAL`，`READY` 不等于设计合格。

## 变更

本节末。

### 13.3.0
- 图表：组合折线层显式不平滑；非负柱/面积默认 0 起点（`verify` 拦成品被改起点）；含负值柱预留直标空间；带类别名的饼内缩绘图区；中文单位紧贴数字。
- 核验：修复组合图折线层未开直标时误报 `V-TEXT-RENDER`。
- 清理：删除 `revise.base_style`、`A_NS`、`verify._render_lines` 等无调用代码。
- 文档：新增色彩方向、光影材质、前中后景判断；修正 `highlight_category` 能力表矛盾、中西文空格口径矛盾、不存在的 `dev/`、示例稿与历史版本引用。
