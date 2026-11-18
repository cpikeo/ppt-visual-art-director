# ppt-visual-art-director

**v2.0.0** — 让 AI 以视觉设计总监的方式做 PPT：先判断，再表达，最后确定性地落成原生可编辑 PPTX。

它不做模板套用，也不从行业关键词推风格。工作方式是：读懂受众与证据 → 找一条洞察 → 定一个能约束全篇的视觉概念 → 排叙事节奏 → 逐页构成 → 编译 → 看真实渲染 → 定点修正。

> 一句话原则：设计判断归你，确定性执行归脚本，审美与签收归人。机器检查不等于好看，预览不等于目标环境通过。

## 结构

```
SKILL.md                  判断链、原则、路线、执行与验收
references/
  design.md               概念、叙事、构图、版式语法、信息图形、诊断、复核维度
  craft.md                排版、尺度、色彩、图像与提示词
  build.md                plan 契约、图表与素材、页码、网格、命令、诊断码、机器边界
scripts/
  compile.py              plan.json → 原生可编辑 PPTX（校验 + 编译）
  verify.py               plan ↔ PPTX ↔ 真实渲染（对象核对、渲染核对、证据报告）
  inspect_pptx.py         只读解析既有 PPTX（改稿与评审第一步）
  grid.py                 列数 → 英寸坐标，标出页码保留框
  plan.schema.json        数据契约（字段形状）
  limits.json             技术阈值与默认值的唯一来源
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
python scripts/compile.py plan.json out.pptx      # 再编译
python scripts/verify.py plan.json out.pptx       # 最后核验（写 render/）
```

plan 的最小骨架见 `references/build.md`「最小骨架」。改稿先跑 `python scripts/inspect_pptx.py 原稿.pptx`。

## 三条硬事实

- **坐标单位是英寸**，原点在页面左上角；16:9 画布 13.333 × 7.5；字号单位是 pt。
- **编译器总是生成新文件**，不就地修改原稿；原稿含自定义母版、动画、SmartArt 时，重建会丢失这些特性，必须提前说明。
- **视觉质量只有在真实渲染且确实看过图像时才可能被判断**。缺任一条件，如实说未核验，不要说"已通过"。
