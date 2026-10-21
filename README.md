# PPT Visual Art Director

**内容 → 逐页判断 → 原生可编辑 PPTX。** 判断透明、执行确定、证据可核。

这是一个技能包（Skill），不是模板库、组件库或设计系统。它做的是**判断**：
这一页要让观众记住什么、证据是什么形状、为什么这样构图、哪些东西必须删；
以及「这些判断到底有没有落到 PPTX 里」的可核对凭证。

```
plan（判断）  →  build（落几何）  →  check（硬门 + PPTX + 回读对账 + 渲染对照）
```

## 快速开始

```bash
pip install -r requirements.txt          # Python 3.10+；无系统级依赖

python scripts/pvad.py plan  examples/demo.yaml --out plan.json
python scripts/pvad.py build plan.json --brief examples/demo.yaml --out deck.json
python scripts/pvad.py check deck.json out.pptx --mode release

# 图还没有的时候：先出生产说明，再谈交付
python scripts/pvad.py media deck.json --out media_briefs.json   # 提示词 / 画幅 / 目标像素
make briefs                          # 需要生产一张图的完整示例（examples/needs-image.yaml）
```

产物：

```
plan.json                     逐页判断（依据 / 未决问题 / 策略指纹）
deck.json                     几何 + 字号 + 数据（作者可直接改；改过的元素标 layout_source=author）
out.pptx                      原生文本 + 原生图表（可编辑）
pvad-report/manifest.json     全部证据与结论（PASS / BLOCKED + 未验证项）
pvad-report/preview-*.png     渲染对照拼图（spec 版 / PPTX 回读版）
```

## 目录

```
SKILL.md                 判断协议：顺序、原则、反例、冲突优先级、验收标准
policy/
  strategy.yaml          阈值 / 枚举 / 字号阶梯 / 颜色 / 词表与探针（唯一数值来源）
  composition.yaml       构图配方 + 为什么 / 何时不成立
  prompts.yaml           提示词语言片段与顺序（素材生产说明；含风格词自审清单）
schema/
  brief.schema.json      输入契约
  plan.schema.json       判断输出形状
  deck.schema.json       执行契约（几何 + 数据）
  policy.schema.json     策略形状
  composition.schema.json
  prompts.schema.json    提示词片段形状
scripts/
  pvad.py                唯一入口：plan / build / check / media
  policy.py              策略加载 + 语义校验（引用完整性）
  brief.py               输入校验与归一（缺失就记，不补编）
  judgment.py            判断协议的可执行部分（带 source / why / 未决问题）
  compose.py             判断 → 几何（配方推导，作者可覆盖）
  typography.py          中文度量与断行（真实字宽、单位绑定、禁则、墨迹几何）
  assets.py              素材度量（像素/亮度/纹理/对比度/裁切损失；只测不判）
  prompt.py              素材生产说明（拍什么/为什么/给谁让位/交付多大；纯翻译）
  guard.py               硬门（只拦确定性错误）
  compile.py             确定性 PPTX 编译（含 OOXML 时间戳归零）
  fidelity.py            spec 与 PPTX 回读逐元素对账
  render.py              同一渲染器吃两种真相 + 像素对照 + 拼图
tests/                   181 个用例：策略 / 输入 / 判断 / 排版 / 素材 / 提示词 / 边界 / 硬门 / 产物 / 端到端
evals/
  judgment_eval.py       离散判断基准（含无判断基线对照与词表消融）
  typography_eval.py     中文排版对照（朴素实现 vs 本实现）
  prompt_eval.py         提示词对照（说明完备性 vs 朴素提示词；不评图）
  blind_pack.py          盲评包与计分（未回收就是 NOT RUN）
  cases.yaml             32 个判断用例，每条对应一条协议纪律
  prompt_cases.yaml      8 条素材生产说明用例（拍什么/为什么/让位/交付尺寸）
examples/                demo.yaml（已有素材）+ needs-image.yaml（待生产）+ 已跑通的产物
references/              contract / policy / verification（按需读取）
```

## 质量边界（请认真读）

- **`check` 只拦硬错误**：越界、相撞、装不下、字段名写错、图表载荷非法、来源缺失。
  输出只有 PASS / BLOCK；**PASS 不等于设计通过**。
- **证据分层**：判断 → 结构硬门 → 真实产物 → 回读对账 → 渲染对照 → 人工盲评。
  各层结论分开写，不互相代替。
- **目标 PowerPoint 未验证**：本包不调用 PowerPoint，也不用其他引擎冒充它渲染。
  字体回落、真实断行、投影可读性必须人工在目标软件里确认。
  发布结论恒为两句：`PASS — Structural Preview` + `NOT VERIFIED — Target PowerPoint`。
- **视觉质量靠人**：`evals/blind_pack.py` 提供打散、去标识的盲评包与计分；
  没有回收评分表时输出 `NOT RUN`，不会把「没人评」记成通过或 0 分。
- **素材有硬门，但仍非审美**：缺图、格式、分辨率、裁切、文字压图的实测对比度、单页预算都会拦；
  「这张图好不好、合不合叙事」仍然只能人判。
- **边界是测试，不是口号**：`tests/test_boundaries.py` 机械检查「策略有没有人消费、阻断清单与错误码是否一致、
  枚举是否只有一份来源、判断字段有没有进执行流程、词表命中会不会改变几何」。越界会让测试失败。
- **提示词管「可执行」，不管「好看」**：缺图时会产出说清拍什么/为什么/给谁让位/交付多大的说明，
  并禁止坐标与风格词；但本包不调用任何图像模型，**没有验证过这些提示词在真实模型上出图的效果**。

## 四层边界

| 层 | 载体 | 管什么 |
|---|---|---|
| 判断协议 | `SKILL.md` | 为什么这样判断、什么时候不成立 |
| 结构化策略 | `policy/` | 多少：枚举、阈值、字号、预算、配方与词表 |
| 确定性执行 | `scripts/` | 算：解析、推导、度量、编译、对账 |
| 契约与证据 | `schema/`、`tests/`、`evals/` | 形状与证据 |

**散文管为什么，策略管多少，Schema 管形状，Python 管算，测试管证据。**

代码里不允许出现新的魔法数字或判断词表：阈值去 `policy/`，判断依据去 `SKILL.md`。

## 测试与评测

```bash
python -m unittest discover -s tests -p "test_*.py"     # 181 用例，约 47s
python evals/judgment_eval.py --baseline --ablate       # 判断基准 + 基线 + 消融
python evals/typography_eval.py                         # 中文排版对照
python evals/prompt_eval.py                             # 提示词说明完备性对照
python evals/blind_pack.py build --group candidate=... --group baseline=... --out pack/
```

当前实测见 [REPORT.md](REPORT.md)：判断基准 32/32、无判断基线 7/32、排版对照三项指标全部收敛到 0，
提示词说明完备性 45/45 条款（朴素提示词 16/19）、真实产物 41 元素回读全对齐、1 张素材 sha256 与包内字节一致。

## 版本

v1.0 · 策略指纹、Schema、评测数字与产物哈希均记录在 `examples/output/report/manifest.json`。
License: MIT
