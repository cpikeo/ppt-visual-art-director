# ppt-visual-art-director

以视觉艺术总监的判断制作、改稿或评审 PPT。AI 负责设计判断，脚本负责确定性执行，人负责最终视觉签收。

交付原生可编辑的 PPTX，并用对象核对与真实渲染给出证据。机器检查不等于好看，预览不等于目标环境通过。

## 结构

```
SKILL.md                  判断原则、判断顺序、brief、任务路线、执行、验收、交付（AI 读的主文件）
references/
  design.md               概念、艺术指导、叙事、空间、信息设计、排版、色彩、图像、诊断、复核维度
  build.md                plan 语义、坐标、叠放、文本与 runs、图表口径、素材、命令、诊断码、定点编辑、机器边界
scripts/
  compile.py              plan.json → 原生 PPTX（校验 + 编译 + 整稿模式信号）
  verify.py               plan ↔ PPTX ↔ 真实渲染（对象核对、逐元素文字归属、图表直标、结构核对）
  revise.py               既有 PPTX 的定点编辑（操作白名单 + 授权比对 + 四层保留核对）
  inspect_pptx.py         只读解析既有 PPTX（页面、对象、样式、部件指纹）
  grid.py                 列数或比例 → 英寸坐标；三分线与黄金位
  plan.schema.json        plan 的数据契约
  limits.json             技术阈值、默认值、字体别名
tests/                    仓库开发用的回归与样例（不属于技能包运行所需）
requirements.txt
```

## 安装

```bash
python -m venv .venv && . .venv/bin/activate      # Windows：.venv\Scripts\activate
pip install -r requirements.txt
```

Python 3.9+。脚本不联网。真实渲染核验需要 LibreOffice；缺失时 `verify.py` 给出 `PARTIAL`，并说明哪些判断因此没有依据。

## 常用命令

```bash
python scripts/grid.py --span 1:6 --split 5,7 --guides thirds
python scripts/compile.py plan.json out.pptx
python scripts/verify.py plan.json out.pptx
python scripts/inspect_pptx.py 原稿.pptx
python scripts/revise.py 原稿.pptx edit.json out.pptx
python -m unittest discover -s tests -v
```

## 状态

`BLOCK` 有硬问题；`PARTIAL` 对象核对通过但渲染没跑；`READY` 对象与渲染都一致。三档都不表示"设计通过"，`human` 永远是 `NOT REVIEWED`。

## 测试

`tests/test_regression.py` 每条负向用例都先人为破坏一个属性，再确认核对会响：未知键被拒、未授权属性被改、隐藏页特性被删、图表第 13 个数据点被改、页序与部件名不一致、重名对象、LibreOffice 另存文件的误报、单元素缺字归属、直标单位不一致。渲染相关用例在没有 LibreOffice 时跳过（会在输出中显示）。

## 未验证的部分

成稿没有在真实 Microsoft PowerPoint 中打开过，只在 LibreOffice 下渲染；`duotone` 在 LibreOffice 只体现为去饱和；没有做人工视觉评审，所以本包任何地方都不声称"设计通过"。
