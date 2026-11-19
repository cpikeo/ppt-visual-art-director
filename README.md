# ppt-visual-art-director

以视觉艺术总监的判断制作、改稿或评审 PPT。AI 负责设计判断，脚本负责确定性执行与事实核验，人负责最终视觉签收。

交付原生可编辑的 PPTX（只输出 `.pptx`），并用对象核对与真实渲染给出证据。机器检查不等于好看，预览不等于目标环境通过。

## 结构

```
SKILL.md                 判断原则、判断顺序、brief、任务路线、执行、验收、交付（AI 读的主文件）
references/
  design.md              概念、艺术指导、叙事、空间、信息设计、排版、色彩、图像、诊断、复核维度
  build.md               plan 语义、几何、文本与 runs、素材与生成图、图表、命令、诊断码、定点编辑、报告、机器边界
scripts/
  compile.py             plan.json → 原生 PPTX（校验、编译、成品包校验）
  verify.py              plan ↔ PPTX ↔ 真实渲染（对象核对、逐元素文字归属、字体替换检查、状态判定）
  revise.py              既有 PPTX：inspect 只读解析；定点编辑 + 保留核对（富文本、超链接、OOXML 结构）
  plan.schema.json       plan 的数据契约
  limits.json            技术阈值与默认值的唯一来源
tests/                   回归用例（仓库开发用，不属于技能包运行所需）
requirements.txt
```

## 安装

```bash
python -m venv .venv && . .venv/bin/activate      # Windows：.venv\Scripts\activate
pip install -r requirements.txt
```

Python 3.9+。脚本不联网。真实渲染核验需要 LibreOffice 与 PyMuPDF；缺失时 `verify.py` 给出 `PARTIAL`，并说明哪些判断因此没有依据。

## 常用命令

```bash
python scripts/compile.py plan.json out.pptx          # 新建：校验 + 编译
python scripts/verify.py plan.json out.pptx           # 新建稿核验，写 render/<名称>/
python scripts/revise.py inspect 原稿.pptx             # 改稿／评审：只读解析
python scripts/revise.py 原稿.pptx edit.json out.pptx # 改稿：定点编辑（保留核对，原稿不动）
python -m unittest discover -s tests -v               # 回归用例
```

## 状态

- `BLOCK`：有硬问题。
- `PARTIAL`：对象核对通过但渲染没跑；或声明字体被替换，断行只代表替代字体。
- `READY`：对象一致、渲染完成、声明字体都真实参与渲染。

三档都不表示“设计通过”，`human` 永远是 `NOT REVIEWED`。

## 测试

`tests/test_regression.py` 每条负向用例都先人为破坏一个属性，再确认核对真的会响，覆盖：未知键与格式错误被拒、未授权属性或 XML 被改、超链接与富文本丢失、页面结构与关系丢失、隐藏页特性被删、PPTM 与宏被拒、素材 `generated` 与 `generation` 矛盾、生成图作证据、哈希不符、图表第 13 个数据点被改、页序与部件名不一致、重名对象、字体替换不得判 READY、LibreOffice 另存文件的误报。需要 LibreOffice 的渲染用例在缺失时跳过（输出中会显示）。

## 未验证的部分

成稿没有在真实 Microsoft PowerPoint 中打开过，只在 LibreOffice 下渲染；沙箱内可用的字体与目标机不同，断行与宽度的结论只对替代字体有效；`duotone` 在 LibreOffice 只体现为去饱和；没有做人工视觉评审，因此本包任何地方都不声称“设计通过”。
