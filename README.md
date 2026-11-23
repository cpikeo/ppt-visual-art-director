# ppt-visual-art-director

以视觉艺术总监的判断制作、改稿或评审 PPT。**AI 负责判断，脚本负责确定性执行与事实核验，真实渲染证明表现，人负责签收。** 交付原生可编辑 `.pptx`。机器检查不等于好看，预览不等于目标环境通过。

## 结构

```
SKILL.md                  判断原则、顺序、路线、验收、交付（入口，约 100 行）
references/
  design.md               概念 · 艺术指导 · 叙事 · 空间 · 信息设计 · 排印 · 色彩 · 诊断 · 复核 · 真稿判断
  imagery.md              图片与视觉资产、镜头单
  build.md                plan 语义 · 图表 · 命令 · 诊断码 · 定点改稿 · 能力边界
scripts/                  model · ooxml · charts · checks · audit · compile · verify · revise
                          plan.schema.json（契约）  limits.json（阈值）
requirements.txt          必需：python-pptx · jsonschema · Pillow · lxml
requirements-render.txt   可选：PyMuPDF（AGPL-3.0）· numpy；另需 LibreOffice

```

打包技能时只取 `SKILL.md`、`references/`、`scripts/`、`requirements*.txt`。

## 命令（在仓库根目录运行）

```bash
pip install -r requirements.txt                # 渲染核验：pip install -r requirements-render.txt
python scripts/compile.py plan.json out.pptx [--assets DIR]
python scripts/verify.py  plan.json out.pptx [--assets DIR]
python scripts/revise.py inspect 原稿.pptx
python scripts/revise.py 原稿.pptx edit.json out.pptx
```