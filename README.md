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
| `dev/` | 测试、生产验证、独立审计（不进安装包） |
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

## 开发、测试与评测（不属于 Skill 安装包）

`dev/` 存放测试、生产验证与独立审计，**不随正式包分发**（`python dev/package.py` 只打包运行必需文件）。

```bash
python -m unittest discover -s dev/tests -v   # 回归测试
python dev/run_evals.py                       # 5 个生产案例：编译 → verify → 独立审计；改稿含负例
python dev/package.py                         # 生成 dist/ppt-visual-art-director-<version>.zip
```

结果与环境见 `dev/evals/RESULTS.md`、`dev/evals/environment.json`，详细口径见 `dev/README.md`。对象通过、LibreOffice 渲染通过、PowerPoint 实机通过、人工设计验收是四件不同的事，报告里分列，未执行的写 `NOT RUN`。

## 变更

本节末。

### 13.4.0
- 核验：散点 marker 成品审计补齐——`verify` 逐系列核对 marker 填充/边线与调色板；`chart_audit` 经 marker 读取散点颜色（此前读被关闭的连线，散点的对比度与灰度审计始终取不到色）；新增 `V-MARKER`（marker 不可见、过小、连线未关、折线点线异色）。
- 改稿：`chart_data` 数值有限性与新建图表同一规则（共用 `model.non_finite_series`）；拒绝布尔/字符串/NaN/Infinity、系列名重复、越过原图固定轴范围的数值、散点图。
- 编译：异常诊断 PPTX 走与交付物相同的 `guard_paths`（不经符号链接、不写目录、已存在无 `--force` 不覆盖），并改为原子写入。
- 契约：scatter 系列允许单点（原要求至少 2 点，无法表达“单独标出一个主角”）。
- 新增 `dev/`（不入安装包）：回归测试、四个新建稿生产案例 + 一个复杂既有 PPTX 改稿案例、与编译器无共用逻辑的独立成品审计与改稿保留审计、环境记录、打包脚本。

### 13.3.0
- 图表：组合折线层显式不平滑；非负柱/面积默认 0 起点（`verify` 拦成品被改起点）；含负值柱预留直标空间；带类别名的饼内缩绘图区；中文单位紧贴数字。
- 核验：修复组合图折线层未开直标时误报 `V-TEXT-RENDER`。
- 清理：删除 `revise.base_style`、`A_NS`、`verify._render_lines` 等无调用代码。
- 文档：新增色彩方向、光影材质、前中后景判断；修正 `highlight_category` 能力表矛盾、中西文空格口径矛盾、不存在的 `dev/`、示例稿与历史版本引用。
