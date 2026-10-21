# 证据与验收（怎么诚实地说「做完了」）

## 1. 六个层次，结论分开写

| # | 层 | 命令 / 工具 | 能证明 | 不能证明 |
|---|---|---|---|---|
| 1 | 判断 | `pvad.py plan` | 角色/证据/主张来源/构图理由/未决问题 | 判断是否好看 |
| 2 | 结构硬门 | `pvad.py check --mode spec` | 越界、相撞、装不下、字段名、图表载荷 | 层级成立 |
| 3 | 产物 | `check --mode draft` | 编译成功、原生文本与图表 | 打开后长什么样 |
| 4 | 回读对账 | `fidelity.compare` | 判断被完整写进 PPTX（几何/字/色/数据/字体角色） | 目视质量 |
| 5 | 渲染对照 | `render.pixel_diff` | 预览与回读的一致（同一渲染器、同一字体度量） | 目标 PowerPoint 的排版 |
| 6 | 人工盲评 | `evals/blind_pack.py` | 焦点/层级/构图/排版/可信度/节奏 | —（未跑就写未跑） |

前五层是机器证据，第六层只能由人给。任何一层都不能替另一层签字。

## 2. 为什么渲染对照是「同一渲染器吃两种真相」

`render.py` 把 spec 与「从 PPTX 回读的事实」构造成同一种页面模型，用同一个渲染器画两遍。
因此像素差只可能来自**编译或字段丢失**，不会来自两个渲染引擎的口径差异——
这正是「预览与最终 PPTX 不一致」的可测形式。

它仍然不是目标 PowerPoint 的渲染：字体回落、真实行距、投影可读性都不在其中。
manifest 里 `verification.target_powerpoint` 与 `target_renderer` 恒为 `NOT VERIFIED`。

## 3. 判断基准（`evals/judgment_eval.py`）

```bash
python evals/judgment_eval.py                  # 29 例离散判断
python evals/judgment_eval.py --baseline       # 与「无判断基线」对照
python evals/judgment_eval.py --ablate         # 逐条词表消融
```

- 用例集与本技能包同步编写，**不是独立盲标**；被看过之后只能算回归集，不能当泛化证明。
- 无判断基线 = 「有架构、没有设计判断」的系统：每页都当一段文字排版。
- 消融回答的是「用例集测到没测到这条规则的独立作用」，不回答「这条规则有没有用」。
  未覆盖 ≠ 可删：删之前先补一条会失败的用例。

## 4. 中文排版对照（`evals/typography_eval.py`）

同一语料、同一字体度量、同一宽度集合，跑两条实现：

| 指标 | 含义 | 方向 |
|---|---|---|
| `unit_split_rate` | 数字与单位被拆到两行的比例 | 越低越好 |
| `forbidden_start_rate` | 以禁则字符开头的行占比 | 越低越好 |
| `char_wrap_rate` | 没有标点可依、纯逐字断开的行占比 | 越低越好 |

基线是「逐字贪心断行」的朴素实现（无单位绑定、无禁则、无断点偏好）。
结果写在 `evals/reports/typography-baseline.json`；这只证明度量模型里的排版改善，
**不证明目标 PowerPoint 的断行**。

## 5. 盲评流程（`evals/blind_pack.py`）

```bash
# 1. 打包：多份 deck 的渲染图打散、去标识、随机编号
python evals/blind_pack.py build --group candidate=renders/candidate --group baseline=renders/baseline \
    --out pack/ --label "Q1 复盘 v2"

# 2. 交给评审者：pack/pages + pack/rubric.md，回收 sheet.yaml
# 3. 计分：均值、组间最大分差、（≥2 人时）一致性
python evals/blind_pack.py score --pack pack/ --sheet sheet-alice.yaml --sheet sheet-bob.yaml
```

纪律：

- 量表 1–7，只接受整数；量表外分数直接报错（不夹紧、不四舍五入）。
- 没有有效评分时输出 `NOT RUN`，并以退出码 1 结束——**不得把「没人评」记成通过或 0 分**。
- 分差 < 0.5 时不得宣称任一组更好。
- 一致性高 ≠ 质量高：一致地低分也是共识。

## 6. 发布结论模板

```
PASS — Structural Preview
  guard:  PASS（0 error / N warnings）
  fidelity: MATCH（matched 43 elements）
  pixel:   worst 0.0000 ≤ limit 0.004
  sha256:  <output_sha256>
NOT VERIFIED — Target PowerPoint
  未在目标 PowerPoint 中打开 / 编辑 / 另存 / 投影；字体回落与真实断行未验证。
NOT RUN — Independent Visual Review
  盲评包已生成（pack/），尚无回收的评分表。
```

## 7. 报告里必须出现的四件事

1. 每条结论对应哪一层证据（不许把结构 PASS 写成「视觉验收通过」）。
2. 未执行/未验证的项目，逐条列出原因。
3. 基线对照的数字与口径（基线是什么、样本有多大、什么条件下可比）。
4. 已知风险与下一步优先级。
