# ppt-visual-art-director

以视觉艺术总监的判断制作、改稿或评审 PPT。**AI 负责设计判断，脚本负责确定性执行与事实核验，真实作品负责证明质量，人负责最终签收。**

交付原生可编辑的 PPTX（只输出 `.pptx`），并用对象核对与真实渲染给出证据。机器检查不等于好看，预览不等于目标环境通过。

## 目录

```
SKILL.md                 判断原则、判断顺序、brief、任务路线、验收、交付（AI 读的主文件）
references/
  design.md              概念、艺术指导、摄影指导、叙事、空间、信息设计、排版、色彩、图像、诊断、复核、评测教训
  build.md               plan 语义、几何、素材与生成图、图表、命令、诊断码、定点编辑、能力边界、模块依赖
scripts/                 ── 运行包 ──
  model.py               共用：诊断、plan 载入、度量、颜色、几何、成品包校验
  ooxml.py  charts.py    OOXML 低层写入；原生图表与表格
  checks.py              plan 校验与整稿模式信号
  compile.py             plan.json → 原生 PPTX
  verify.py              plan ↔ PPTX ↔ 真实渲染（含环境事实）
  revise.py              既有 PPTX：inspect 只读解析；定点编辑 + 保留核对
  plan.schema.json       数据契约（唯一）
  limits.json            技术阈值（唯一）
requirements.txt  LICENSE

```

## 环境

- Python 3.9+；`pip install -r requirements.txt`（Windows：`.venv\Scripts\activate`）。
- 渲染核验（可选）：LibreOffice（`soffice` 在 PATH，或用环境变量 `SOFFICE` 指定）与 PyMuPDF；缺失时 `verify.py` 给出 `PARTIAL`。
- 字体：只写字体名、不嵌入。样本使用 Noto Sans/Serif CJK SC（OFL，跨平台可装）；换成微软雅黑等字体后的断行需在目标机复核。
- 脚本不联网。

## 常用命令

```bash
python scripts/compile.py plan.json out.pptx          # 新建：校验 + 编译
python scripts/verify.py plan.json out.pptx           # 核验，写 render/<名称>/（PDF、PNG、缩略总览）
python scripts/revise.py inspect 原稿.pptx             # 改稿／评审：只读解析
python scripts/revise.py 原稿.pptx edit.json out.pptx # 改稿：定点编辑（保留核对，原稿不动）
python -m unittest discover -s dev/tests -v           # 回归用例
python dev/evals/run_evals.py --version vN            # 真实作品评测（编译 + 核验 + 渲染 + 汇总）
```

## 能力边界

| 级别 | 内容 |
| --- | --- |
| 完整支持 | 16:9 / 4:3；原生文本、形状、折线路径、组、图片裁切；原生图表（柱、线、面积、饼、环、散点）与表格；演讲者备注；替代文字；页码；既有稿定点编辑 |
| 有限支持 | 页标题是普通文本框（读屏识别不到）；无双数值轴、桑基、甘特；渐变与 `duotone` 在 LibreOffice 只是近似；连线是坐标而非连接符 |
| 外部依赖 | 图片生成（工具无关：只登记工具、提示词、文件、SHA-256、权利状态，不调用任何模型）；LibreOffice；PowerPoint 实机验证；目标机字体 |
| 不支持 | 动画、切换、嵌入字体、新建 SmartArt、宏（`.pptm` 拒绝） |

## 验证等级与状态

`verify.py` 的状态：`BLOCK` 有硬问题；`PARTIAL` 渲染没跑或声明字体被替换；`READY` 对象一致、渲染完成、字体都真实参与。三档都不是设计结论，`human` 永远是 `NOT REVIEWED`。

环境等级 L0 对象 / L1 LibreOffice / L2 PowerPoint 实机采集 / L3 人工签收，见 `dev/compat/COMPAT.md`。**本包目前只证明到 L1；PowerPoint 实机为 NOT VERIFIED。**

