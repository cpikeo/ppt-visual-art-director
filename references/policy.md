# 策略层（调「多少」，不改「为什么」）

`policy/strategy.yaml` + `policy/composition.yaml` + `policy/prompts.yaml` 是全部数值、
枚举与**提示词语言片段**的唯一来源。三份文件都进策略指纹。
改完必须跑：`python -m unittest discover -s tests` 与 `python evals/judgment_eval.py --ablate`。

## 1. 加载与校验

```
schema/policy.schema.json        形状（类型、必填、取值范围）
schema/composition.schema.json   配方形状
schema/prompts.schema.json       提示词片段与顺序（含 style_words_banned 自审）
scripts/policy.py::validate_semantics  跨字段关系（Schema 管不到的部分）
```

`schema/policy.schema.json` 拒绝：缺键、类型错、`#HEX` 不合法、枚举空。
`validate_semantics` 拒绝：

1. `floors/line_height/tracking_pt/max_lines` 与 `scale` 键不一致（「有字号没下限」）；
2. floor 大于 scale；
3. `colors.tokens` 与 `colors.defaults` 不一致（多写 token 不写值＝渲染时静默回落）；
4. `enums.composition` 与 `composition.yaml` 的配方不一致；
5. 配方槽位引用了不存在的字号档 / 颜色 token / 槽位类型；
6. 槽位 zone 越出内容区；
7. 有词表没有探针（命中后无处可问＝噪声）；
8. 缺判断必需的词表（`comparison_end` / `part_whole` / `sequence` / `structure` /
   `claim_cue` / `assertion_cue` / `change_cue` / `imperative_cue`）；
9. `verification.default_mode` 不在 `verification.modes` 里；
10. 配方缺 `must_not` / `ladder` / `focus`，或 `ladder` 不在 `typography.scale`、`focus` 不在 `enums.focus`；
11. `prompts.phrases.function / .role` 与 `enums.media_function / media_role` 不一一对应
    （有枚举没短语 = 提示词静默缺一句职责）；
12. `prompts.clause_order` 缺必需段（role / subject / function / framing / ratio）；
13. 提示词短语里出现 `style_words_banned` 中的词（策略偷渡风格 → 整批稿子趋同）。

策略文件带 SHA-256：`plan.json` 与 manifest 都记录它，回放时能确认「同一套策略」。

**死配置是缺陷**：`tests/test_boundaries.py` 会遍历 `strategy.yaml` 里的每个数值叶子，
凡是在 `scripts/` 里找不到消费者的都算失败。历史上这样抓出并处理了 6 个键
（`colors.contrast.*` 没有对比度检查、`assets.safe_area.min_width_ratio` 没接线、
`assets.qc.min_negative_ratio` 算了不用、`verification.exit_codes.usage` 写死 2、
`thresholds.contrast_ratio_floor` / `asset_luma_delta` 与素材段重复且无人读）。

## 2. 常用调参位

| 想做 | 改哪里 | 注意 |
|---|---|---|
| 字号更克制/更大 | `typography.scale` | 必须同步 `floors`；阶梯要单调递减 |
| 内容更密 | `thresholds.line_width_ratio`（文本可用宽度）、`typography.line_height` | 不要用缩字号换密度 |
| 更严的相撞判据 | `thresholds.ink_overlap_px` | 太小会把正常间隙判成相撞 |
| 强调色更少 | `colors.accent_budget_per_page` | 预算 0 = 全稿单色 |
| 更严的回读对账 | `thresholds.position_drift_px` / `size_drift_px` | 受 EMU 取整影响，别小于 0.5 |
| 渲染对照更严 | `thresholds.fidelity_pixel_mismatch_max` | 0 表示逐像素一致 |
| 图表类别上限 | `thresholds.chart_max_categories` | 超过应改表格或拆分 |
| 高潮页更强调 | `typography.emphasis_ladder_boost.lead` | 1 档太大时，宁可改内容也不要缩字 |
| 断行更保守 | `typography.line_break.min_line_fill` | 调到 0.7 会让行更短、断点更多 |
| 混量纲图更早拦下 | `thresholds.chart_scale_skew_ratio` | 100 倍是告警线；`CHART_MIXED_UNITS` 是硬错误 |
| 哪个错误码才阻断 | `verification.blocking_codes` | 必须与 guard 实际抛出的码一一对应（测试守） |
| 图片解码预算 | `assets.pixels.max_decode` | 1GB 内存下的物理边界，不是审美 |
| 提示词更克制 | `prompts.phrases`、`prompts.limits.max_chars` | 改短语要跑 `make eval`（提示词对照会立刻反映） |
| 画幅吸附更宽/更紧 | `prompts.ratio.tolerance` | 太宽会把 0.9 的图位说成 1:1，导致出图被裁 |
| 交付像素更保守 | `assets.oversample`、`assets.min_long_side_px` | 同步看 `ASSET_RESOLUTION` 的用例 |
| 换字体 | `typography.families` + `typography.font_files` | 家族数 ≤ 2；`font_files` 要含 ttc index |

## 3. 新增/修改构图配方

每条配方必须有 `why`（为什么这种证据形状配这种构图）、`when`（什么条件下适用）、
`fails`（什么条件下不成立，应改什么）、`slots`（有理由的候选几何）。

```yaml
  my_recipe:
    why: ...            # 至少一句能说服同行的话，而不是「更均衡」
    when: ...           # 证据形状 + 页面任务
    fails: ...          # 反例与替代配方
    slots:
      - {role: claim, type: text, zone: [0.0, 0.08, 0.86, 0.22], size: h1, color: ink, bold: true}
```

规则：

- 槽位 `size` 只能引用 `typography.scale` 的键；`color` 只能引用 `colors.tokens`；
  `type` 只能引用 `enums.slot_type`（`kpi/steps/groups` 是 composite，由 `compose.py` 展开）。
- `enums.composition` 与配方必须一一对应（加了枚举没配方＝运行时崩，加了配方没枚举＝死配置）。
- 配方不是模板库：新增一条之前先问「现有配方为什么不够」，答不出来就不要加。

## 4. 词表纪律（最容易退化成「规则堆」的地方）

- 词表只产生**提示与一条待确认问题**（`judgment.probes`），不产生风格、不直接决定构图。
- 每条词表必须在 `evals/cases.yaml` 里有一条「删掉就会失败」的用例；
  `--ablate` 报「未覆盖」的，要么补用例，要么删掉。
- 不要按行业堆形容词（「科技感」「高级感」）：那正是视觉同质化的来源。
- 判断顺序与陷阱清单在 `../SKILL.md §3`；词表只喂给它们，不替代它们。

## 5. 反例：这些改动应被拒绝

| 改动 | 为什么拒绝 |
|---|---|
| 在 `judgment.py` 里加一个关键词 if 分支 | 判断会藏进代码，无法在本文件与 SKILL 里追溯 |
| 给某个行业加一套配色/字体 | 风格套用；没有品牌依据时应保持中性 |
| 把 `TEXT_OVERFLOW` 改成告警 | 装不下就是装不下，缩字号/隐内容才是真事故 |
| 把 `MEASUREMENT_DEGRADED` 静默忽略 | 估算不能当测量 |
| 给「相加 100%」加例外放行 | 分母陷阱是结构错误，不是边界情况 |
| 新增 `force_pass` 开关 | 交付可靠性优先于「先过再说」 |
