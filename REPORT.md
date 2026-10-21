# ppt-visual-art-director · 交付报告（v1.0）

日期：2026-10-21 ｜ 范围：独立创建 `ppt-visual-art-director` 技能包 ｜ 策略指纹：`d3029cad…6a3e9fa5`（strategy + composition + prompts 三份）

---

## 0. 一句话结论

技能包已建成并跑通全链路：**判断 → 几何 →（缺图时）生产说明 → 原生 PPTX → 回读对账 → 渲染对照 → manifest 凭证**。
可以交付，但**视觉终验尚未进行**：目标 PowerPoint 未验证、独立盲评未回收评分。
以下所有「更好」都只在被测量的维度上成立，未测量的维度不做宣称。

---

## 1. 交付物

```
ppt-visual-art-director/
├── SKILL.md                    判断协议（324 行）
├── README.md                   使用与质量边界
├── REPORT.md                   本文件
├── Makefile                    test / eval / demo / report / clean
├── requirements.txt
├── policy/{strategy,composition,prompts}.yaml   结构化策略（931 行，唯一数值 / 枚举 / 语言来源）
├── schema/*.json               6 份契约（brief / plan / deck / policy / composition / prompts，2544 行）
├── scripts/                    12 个模块，5088 行
├── tests/                      10 个文件、181 用例、2177 行（含 test_boundaries.py）
├── evals/                      judgment_eval / typography_eval / blind_pack / cases.yaml
├── examples/                   demo.yaml + 已跑通的 plan/deck/pptx/report
└── references/                 contract / policy / verification（按需读取）
```

复现：

```bash
pip install -r requirements.txt
make test          # 134 用例（make assets 只跑素材 27 例）
make eval          # 判断基准（含基线对照与消融）+ 排版对照
make demo          # 端到端跑通 examples/demo.yaml，产物在 examples/output/
make briefs        # 需要生产一张图的最小闭环：plan → build → media（examples/needs-image.yaml）
```

---

## 2. 架构与职责划分

```
brief.yaml ──plan──▶ plan.json ──build──▶ deck.json ──check──▶ out.pptx + manifest.json
                       │                    │                      │
                 判断 + 依据          配方推导的几何         硬门 / 归因 / 对账 / 对照
                 + 未决问题           （作者可覆盖）          （PASS 或 BLOCK，不评分）
```

| 层 | 文件 | 负责 | 不负责 | 越界防线 |
|---|---|---|---|---|
| 判断协议 | `SKILL.md` | 判断顺序、陷阱清单、冲突优先级、人工验收标准 | 阈值、坐标 | 代码里出现判断词表或魔法数字即缺陷 |
| 结构化策略 | `policy/*.yaml` | 枚举、字号阶梯与下限、间距容差、强调色预算、配方 why/fails、词表与探针 | 审美判断 | Schema + `validate_semantics` 双向校验 |
| 确定性执行 | `scripts/*.py` | 解析、归一、推导、度量、编译、对账 | 审美偏好 | 每个模块只返回事实，不返回好坏 |
| 契约与证据 | `schema/`、`tests/`、`evals/` | 形状、边界回归、真实产物、对照与盲评 | 用测试替代设计 | manifest 里未验证项恒为 NOT VERIFIED |

模块级分工：

| 模块 | 只做 | 关键设计 |
|---|---|---|
| `pvad.py` | 编排三步、退出码、指纹校验 | 一趟批处理，无逐页重试循环 |
| `policy.py` | 加载 + 语义校验 + 带指纹访问 | 缺键即抛错，**没有隐式默认值** |
| `brief.py` | 校验 + 归一 + 记录「补了什么默认值」 | 缺字段记 note，不补内容 |
| `judgment.py` | 12 步判断链、证据形状、配方选择、探针 | 每条判断带 `source`/`why`/命中片段 |
| `compose.py` | 配方 → 像素几何、复合槽位展开 | 数据形状不符时降级并写 `composition_downgraded` |
| `typography.py` | 真实字宽度量、token 化、禁则、墨迹几何 | 数字+单位合成不可分单元 |
| `guard.py` | 硬错误拦截 | 只输出 PASS/BLOCK，附根因组与 hint |
| `compile.py` | 确定性 PPTX（原生文本 + 原生图表） | OOXML 与内嵌 xlsx 时间戳归零 |
| `fidelity.py` | spec ↔ 回读逐元素对账 | 区分 dropped / drift / payload / extra |
| `render.py` | 同一渲染器吃两种真相 | 像素差只能来自编译或字段丢失 |
| `prompt.py` | 素材生产说明（提示词/反向词/画幅/目标像素） | 纯翻译：读已结算的 `decision.media` + 几何，不改判 |

### 2.5 四层职责映射（本轮补齐后的状态）

| 层 | 文件 | 只负责 | 明确不做 | 越界防线（机械可判） |
|---|---|---|---|---|
| 判断协议 | `SKILL.md` | 判断顺序、原则、反例、冲突优先级、验收标准 | 阈值、坐标、字段转换 | 代码里出现禁令文案或设计映射即判越界（`test_boundaries`） |
| 结构化策略 | `policy/strategy.yaml` | 阈值、枚举、字阶、预算、词表、探针、阻断清单、解码预算 | 审美判断 | 每个数值必须有消费者；枚举与 Schema 双向一致 |
| | `policy/composition.yaml` | 配方几何 **+ 配方的设计属性**（`must_not` / `ladder` / `focus`） | 组件库、风格主题 | `plan.production.must_not` 必须逐字来自配方 |
| | `policy/prompts.yaml` | 提示词语言片段与顺序、反向词、风格词黑名单 | 判断要不要图 | 短语表与枚举一一对应；策略自身不得含风格词 |
| 确定性执行 | `scripts/judgment.py` | 解析、证据形状、职责、决策 + 依据 | 几何、阈值、禁令清单 | 决策只能读出策略值，不能持有设计表 |
| | `scripts/compose.py` | 配方 → 像素几何、判断 → 字号 | 第二份字阶映射 | 字阶只从配方读（本轮删掉了重复的 `_claim_ladder` 映射） |
| | `scripts/guard.py` | 硬错误判定与归因 | 评分、审美 | 阻断与否由策略的 `blocking_codes` 决定 |
| | `scripts/assets.py` / `prompt.py` | 度量 / 翻译 | 判断与拦截 | 度量只返回事实；提示词不改判 function/role |
| | `scripts/compile.py` / `render.py` / `fidelity.py` | 编译 / 同渲染器双真相 / 逐元素对账 | 设计决策 | 渲染与度量共用 `visible_image()` 一个口径 |
| 契约与证据 | `schema/*.json`、`tests/`、`evals/` | 形状、边界回归、真实产物、对照 | 用测试代替设计 | 判断字段未进执行 → 测试失败；词表改几何 → 测试失败 |

---

## 3. 关键设计决策（含取舍）

**D1 判断与几何分两次落，且都可审计。**
`plan.json` 只有判断与依据，`deck.json` 才是几何。取舍：多一次文件，换来「为什么这样排」可追溯到
`source=author|rule:<id>|absent` 与命中片段；也换来作者可以只改几何而不动判断。

**D2 结构化策略独立成层，代码里不允许出现阈值。**
所有数值（字号阶梯、下限、行高、字距、容差、预算、类别上限）与枚举只在 `policy/`。
取舍：配置读取成本 + 一次校验开销，换来「调参不碰代码」和「策略指纹可回放」。

**D3 判断的措辞线索进策略、进消融，不进代码。**
`claim_cue / change_cue / imperative_cue / comparison_end …` 都在策略里，每条必须有
一条「删掉就会失败」的用例。开发中因此删掉了两条死配置（`topic_tag`、`meta_only`），
并把 `proposition` 从判断路径中移除（未被消费）。

**D4 词表只提示，不决定形态。**
命中只产生一条 open question（`PROBE_*`），构图由证据形状决定。这是与「关键词 → 模板」路线的
根本分歧：后者会产生行业风格同质化。

**D5 形状跟着证据走，缺数据就降级。**
声明 `series_chart` 但没有时间读数 → 降级为 `metric_row` 并写理由，不画空轴；
环图合计 ≠ 100 → 告警「量不该画环」；不同单位 → 禁止相减。取舍：少一些自动出图，多一些不出假图。

**D6 渲染对照用「同一渲染器吃两种真相」。**
把 spec 与 PPTX 回读都构造成同一页面模型，用同一渲染器画两遍。像素差因此只可能来自编译或
字段丢失，而不是两个引擎的口径差异——这是「预览与最终 PPTX 不一致」的可测形式。
取舍：不能替代 PowerPoint 渲染，所以 `target_renderer` 恒为 NOT VERIFIED。

**D7 编译字节确定。**
ZIP 时间戳、内嵌图表工作簿时间戳、核心属性时间全部归零，条目排序。
证据：同一 deck 两次 `check` 的 `output_sha256` 完全一致（开发中正是这个断言发现了内嵌 xlsx 的
时间戳漂移）。

**D8 失败只给根因组，不给逐条重试。**
`fix_groups` 按 code 归并 + 可执行 hint；`BLOCK` 后一次批量改完再跑同一条命令。

**D10 配方的设计属性也归策略（本轮迁移）。**
`must_not`（这个形状下不该做什么）、`ladder`（主张字号档）、`focus`（第一落点）
原本写在 `judgment.py` 的 Python 字典里。它们不是计算，是设计判断的可配置部分，
现在在 `composition.yaml` 每个配方里，`validate_semantics` 校验字阶存在、焦点在枚举内、禁令非空。
取舍：策略文件变长 40 行，换来「改口径不用改代码」和「plan 与策略逐字一致」可测。

**D11 阻断清单进策略，且必须与真实错误码一一对应（本轮迁移）。**
`guard.BLOCKING` 原本是代码里的一个集合。现在它是 `verification.blocking_codes`，
并由测试双向比对：策略里登记但从不抛出的码 = 死配置；抛出了却没登记的码 = 该拦不拦。
迁移时正是这条测试发现了 `OVERLAP` 与 `ELEMENT_OVERLAP` 的名字漂移。

**D12 判断必须落到可测量的执行差异，装不下就诚实降级（本轮新增）。**
`emphasis=lead` 过去只写进 `slide.meta`，没有人受影响。现在它把主张字号档上移一档
（`typography.emphasis_ladder_boost`），并用同一套度量检查「放大一档还装得下吗」：
装得下 → `claim.size` 真的变大、`meta.emphasis_applied = true`；
装不下 → 保持基准档并写 `meta.emphasis_notes`。**不缩字、不撑框、不假装。**

**D13 死配置由测试来抓（本轮新增）。**
遍历 `strategy.yaml` 的每个数值叶子，在 `scripts/` 里找不到消费者的算失败。
一次跑出 6 个：`colors.contrast.*`（没有文字/底板对比度检查）、
`assets.safe_area.min_width_ratio`（没接线）、`assets.qc.min_negative_ratio`（算了不用）、
`verification.exit_codes.usage`（写死 2）、`thresholds.contrast_ratio_floor` 与
`asset_luma_delta`（与素材段重复且无人读）。前三项补上了检查，后三项按「修复 > 删除」分别处理。

**D9 未验证项写死在凭证里。**
manifest 的 `target_powerpoint`、`target_renderer`、`independent_visual_review` 恒为
NOT VERIFIED / NOT RUN，除非真的跑过盲评。盲评工具在无有效评分时输出 `NOT RUN` 且退出码 1。

---

## 4. 复用与新增（诚实说明）

**没有复制上游仓库的代码。** 两个参考仓库只用于审计与思想比对：

- `cpikeo/ppt-visual-art-director-os`（约 2.25 万行、11 个脚本 + 7 组冻结 JSON）
- `cpikeo/skill-creator`（技能包结构、评测与盲评流程、打包规范）

**审计到的、被采纳的设计思想**（重写实现，不是搬运）：

| 思想 | 来源 | 本包落地 |
|---|---|---|
| THINK → BUILD → VERIFY 三段式，一次判断一次收口 | OS 的 SKILL/README | `plan / build / check`，一趟批处理 |
| 判断链与「判断陷阱」清单（分母、单位、顺序、元信息、相关≠因果…） | OS §3 | `SKILL.md §3` + `judgment.py` 的可执行判据 + 29 条回归用例 |
| 「词表只提示」，命中产生待确认问题 | OS §3 | `judgment.probes` + 消融评测 |
| 证据分层，结构预览不等于渲染，发布清单双结论 | OS §6 / contract | `references/verification.md` + manifest 恒 NOT VERIFIED |
| 墨迹跨度按锚点算，不用框高冒充 | OS contract | `typography.ink_height/ink_top`，guard 按墨迹判相撞 |
| 字段名写错 = 静默丢弃，必须拦 | OS contract | `FIELD_UNKNOWN`（含 `font_size`→`size`、`maxlines`→`max_lines` 的近似名识别） |
| 盲评量表 1–7 与「一致性 ≠ 质量」 | OS `bench.py blind` | `evals/blind_pack.py`（明确不再冒充 ICC 统计检验） |
| 技能包结构与渐进披露、评测流程 | skill-creator | 目录约定、`references/` 按需读取、evals 与 tests 分离 |

**审计到的、被刻意不采纳的做法**：

| 上游做法 | 为什么不采纳 |
|---|---|
| 单模块 2000–6000 行（`intelligence.py` 3793、`selftest.py` 5921） | 判断与度量、编译与自检混在大文件里，越界无法用结构约束 |
| 判断词表散落在 Python 常量中 | 改成策略层 + 消融用例，词表可审计、可删 |
| 无 Schema 目录（字段契约写在 md 里） | 改成 5 份 JSON Schema + 语义校验，形状错误在入口就拒绝 |
| 无策略层（阈值硬编码在脚本） | 改成 `policy/` 单一来源 + 指纹 |
| 词表消融后仍有未被覆盖的维度 | 本包要求「每条词表一条会失败的用例」，否则删（已删 2 条） |
| 用 `ghost.py` 结构预览作为主要视觉证据 | 本包改为「同渲染器双真相对照」，并明确它仍不是 PowerPoint |

**新增的、上游没有的能力**（每条都能指出为什么必要）：

1. `policy/` + `policy.schema.json` + `validate_semantics`：让「多少」成为可配置、可校验、可指纹化的一层。
2. `fidelity.py`：spec 与 PPTX 回读逐元素对账（dropped / drift / payload / extra），
   把「判断有没有落进包里」变成机器可判事实。
3. 字节级确定性编译（含内嵌 OOXML 归一）：可回归、可哈希、可复现。
4. `typography.py` 的 token 化：数字+单位合成不可分单元、禁则、标点优先断点、墨迹几何，
   并有可复现的排版对照（`evals/typography_eval.py`）。
5. `evals/judgment_eval.py` 的**无判断基线**：给出「有架构、没有设计判断」的对照分数。
6. `evals/blind_pack.py`：打散去标识、量表校验、`NOT RUN` 纪律。
7. `prompts.yaml` + `prompt.py` + `media` 子命令：把「这页要图」翻译成可执行的生产说明
   （拍什么/为什么/给谁让位/交付多大），并禁止坐标与风格词进提示词——上游的做法是
   2155 行单文件里六层风格叠加，本包只要「能被照着做」的那四句。

### 4.5 能力继承矩阵（上游 → 本包，逐项交代去向）

审计对象：`cpikeo/ppt-visual-art-director-os`（11 个脚本、21,273 行；`SKILL.md` + `references/`）
与 `cpikeo/skill-creator`（技能包结构与评测流程）。**没有复制代码**，只继承能力。

| 上游能力 | 本包落点 | 结论 | 证据 |
|---|---|---|---|
| brief 归一与证据解析（`intelligence.py`） | `brief.py` + `judgment.py` | 保留（重写） | 契约 20 例 + 判断 35 例 |
| 判断陷阱：分母 / 单位 / 顺序 / 元信息 / 相关≠因果 | `SKILL.md §3` + `judgment.py` | 提升 | 判断基准 32/32 vs 基线 7/32 |
| 词表只提示、命中转成待确认问题 | `judgment.probes` + `--ablate` | 保留并加严 | 14 条词表全部有「删掉即失败」的用例 |
| 素材：清单 → 提示词 → QC → 链核验（`assets.py` 2155 行） | `assets.py` + `prompt.py` + `guard` + `media` 子命令 | 保留并拆分 | 素材 27 例 + 提示词 28 例 + 45/45 条款 |
| 原生 PPTX 编译（文字 / 图表 / 图片 / 裁切） | `compile.py` | 保留 + 补图片 | 41 元素回读对齐；字节可复现 |
| 结构预览 `ghost.py`（截断 / 越界 / 字体回落） | `render.py` | 提升为「同渲染器双真相」 | 像素差 0.0000 ≤ 0.004 |
| 字段名写错即静默丢弃 → 必须拦（`verify.consumed_fields`） | `guard.FIELD_UNKNOWN` + **决策消费测试** | 提升 | `test_boundaries::test_every_decision_field_is_consumed…` |
| 原子写盘 / 内容指纹（`primitives.py`） | `pvad.write_json`、`compile_deck` | 本轮补齐 | `test_compiled_pptx_never_left_half_written` |
| 评测：cases / 消融 / 回归 / 盲评包（`bench.py`） | `judgment_eval` + `blind_pack` | 保留；**不采纳 ICC 统计** | 盲评未回收时输出 `NOT RUN`（n 太小，算区间只会显得精确） |
| 未见用例基准（`unseen_bench.py`） | — | **未继承（如实标注）** | 本包用例集与实现同步编写，只作回归，不代表泛化 |
| 风格层 `medium / lighting / material / texture / fusion` | — | **不采纳** | 策略层自动叠加风格 → 整批稿子趋同；改由作者写在 `asset.subject` |
| 单文件数千行（`intelligence.py` 3,793 行） | 12 个模块，最大 700+ 行 | **不采纳** | 越界无法用结构约束；但也不为分层而拆 |

**开发中被测试/评测逼出来的真实修复**（每条都有对应用例）：

| 缺陷 | 现象 | 修复 | 证据 |
|---|---|---|---|
| token 化单元过宽 | 「万元」被拆成「万/元」、「18%，」被吞成一个单元 | 数字+白名单单位合成单元；标点独立成 token | `test_typography`（2 例） |
| 词表在归一化文本上匹配 | 条目被切成 bullets 后，原文「第一步」线索消失，顺序判断失效 | 词表在原文上匹配，归一化只服务排版 | `test_judgment::test_real_sequence_with_order_words` |
| 槽位 zone 二次变换 | 复合槽位（KPI/步骤/分组）坐标被乘以内容区宽高，元素飞到画布外 | 复合槽位统一走像素框 `add_text_box` | `test_guard::test_baseline_passes_in_draft` |
| 单位串丢失 | 行内写法「门店数量 128 家」把单位并进标签，release 报缺单位 | `_split_value_unit` 拆值/单位；反序配对「经销 / 占 51%」 | `test_guard::test_metric_without_unit_is_blocked_in_release` |
| 焦点竞争误报 | 同尺寸的多指标并行数字被判成「多个主角」 | 同 family（`metric_value.*`）与对称两端（left/right）不算竞争 | `test_guard::test_baseline_passes_in_draft` |
| 主张缺席页留白 | 无结论句的页只显示眉标，正文凭空消失 | 主张缺席时用标题作话题标签、降到 lede 档、muted 色 | demo s05 渲染 |
| 环图合计 ≠ 100 静默 | 量被画成占比环 | `DONUT_NOT_PERCENT` 告警 | `test_guard::test_donut_not_summing_to_100_warns` |
| 内嵌 xlsx 时间戳 | 同一 deck 两次编译哈希不同 | OOXML 时间戳归零 | `test_pipeline::test_compiled_bytes_are_reproducible_across_runs` |

### 4.6 本轮变更清单（P1 边界治理 + P2 定向优化）

每条都能指到证据；「行为是否改变」是实测的，不是估计的。

| # | 变更 | 类别 | 原因 | 证据 |
|---|---|---|---|---|
| 1 | `MUST_NOT` / `ladder` / `focus` 从 `judgment.py` 迁到 `composition.yaml` | 迁移 | 设计判断写在 Python 里，改口径要改代码 | `test_boundaries::test_recipe_design_fields_come_from_policy` |
| 2 | 删除 `compose._claim_ladder` 的第二份字阶映射 | 删除 | 同一映射两个模块各写一份，必然漂移 | 同上 + demo 字节未变 |
| 3 | `guard.BLOCKING` → `verification.blocking_codes`，并双向校验 | 迁移 | 什么算阻断是产品决定，不是代码细节 | `test_blocking_codes_match_real_errors`（发现 `OVERLAP`/`ELEMENT_OVERLAP` 漂移） |
| 4 | `assets.MAX_PIXELS` / `_SAMPLE_MAX_SIDE` → `assets.pixels.max_decode` | 迁移 | 阈值硬编码，且 384 与策略重复 | `test_decode_budget_comes_from_policy` |
| 5 | `fidelity.SIZE_PT_TOL` → `thresholds.size_pt_tolerance` | 迁移 | 同上 | 死配置测试 |
| 6 | `typography._BREAK_AFTER_PUNCT` + 55% 行宽 → `typography.line_break.*` | 迁移 | 断行规则是可调策略 | 死配置测试 |
| 7 | `judgment._derive_world` 的中文设计文案 → `judgment.default_world` | 迁移 | 设计语言在代码里，进不了策略指纹 | 死配置测试 |
| 8 | 补齐 4 个死配置：文字/底板对比度、安全区最小宽度、安静区平坦占比、usage 退出码 | 修复 | 写了配置却没人执行 = 两套真相 | `TEXT_CONTRAST_LOW`、`ASSET_SAFE_AREA`、`ASSET_QUIET_AREA`、退出码用例 |
| 9 | 删除 2 个重复死键（`contrast_ratio_floor`、`asset_luma_delta`）与 `text_padding_px` | 删除 | 与 `assets.text_over_image.*` 重复且无人读 | 死配置测试 |
| 10 | 新增 `CHART_MIXED_UNITS`（硬错误）+ `CHART_SCALE_SKEW`（告警） | 新增 | 混合量纲被画成单一图表是用户点名的失效模式 | `test_mixed_units_in_one_chart_is_blocked` |
| 11 | `emphasis=lead` → 字号档上移，装不下就如实降级 | 新增 | 判断只写进 meta 等于装饰 | `test_lead_emphasis_actually_changes_the_type_size` / `..._degrades_honestly` |
| 12 | PPTX 编译改为临时文件 + 原子替换 | 新增 | 中断会留下「看起来完整、其实半截」的交付物 | `test_compiled_pptx_never_left_half_written` |
| 13 | 新增 `tests/test_boundaries.py`（19 例） | 新增 | 让四层边界变成会失败的测试 | 全部通过 |

**行为不变性的证据**：`examples/output/demo.pptx` 的 `output_sha256` 在本轮前后**完全相同**
（`61eb6c18…56601675`，327,303 字节）。也就是说第 1–9 条是纯迁移 + 补检查，
没有改变真实产物的一个字节；改变输出的只有第 10–12 条，且各有专门用例。

---

## 5. 测试与真实产物验证

### 5.1 自动化（全部在本机实跑）

| 项 | 命令 | 结果 |
|---|---|---|
| 单元 + 集成 + 端到端 | `python -m unittest discover -s tests` | **181 passed / 0 failed，47.1s** |
| 四层边界自检 | 同上（`test_boundaries.py`） | **19 passed**：死配置 / 阻断清单一致性 / 枚举单一来源 / 判断进执行 / 词表不改几何 / 原子写盘 |
| 素材专项 | `python -m unittest discover -s tests -p "test_assets.py"` | **27 passed**（含 3 条反假通过用例） |
| 判断基准 | `python evals/judgment_eval.py` | **32/32（100%）** |
| 无判断基线对照 | 同上 `--baseline` | **7/32（21.9%）**；差距 **+78.1pt** |
| 词表消融 | 同上 `--ablate` | 14 条词表**全部被覆盖**（含 `visual_subject → J31`），未覆盖 0 |
| 提示词说明完备性 | `python evals/prompt_eval.py` | **45/45 条款（8/8 用例）**；朴素提示词 16/19 条款（0/8） |
| 中文排版对照 | `python evals/typography_eval.py` | 180 组样本，见 5.3 |
| 端到端示例 | `make demo` | `PASS`（release 模式），见 5.2 |

覆盖分布：策略/契约 20 · 判断 35 · 排版 15 · 素材 27 · 提示词 28 · 边界 19 · 硬门 35 · 产物 15 · 端到端 16（含重叠）。

### 5.2 真实 PPTX 产物（`examples/output/`）

| 项 | 值 |
|---|---|
| 文件 | `examples/output/demo.pptx`（6 页 / 41 元素 / 1 个原生图表 / 1 张实拍尺寸照片 / 320 KB） |
| `output_sha256` | `61eb6c18f3e6135972cf31198b97f5b3fea4c617677c023131b7708c56601675` |
| 结构硬门 | `guard: PASS`，0 error / 0 warning，字体度量未降级，`stats.images = 1` |
| 回读对账 | `fidelity: MATCH`，41/41 元素对齐（文字、字号、颜色、几何、图表数据、图片 sha256/尺寸/裁切） |
| 渲染对照 | `pixel: MATCH`，最差像素差 **0.0000** ≤ 阈值 0.004（6 页全取） |
| 可编辑性 | 文本为原生文本框；图表为原生 chart part（数据可改）；图片为原生 `p:pic`（可换图、可裁切）；`a:ea` 中文字体已绑定 |
| 确定性 | 同 deck 两次编译字节完全相同 |
| 未验证 | 目标 PowerPoint 打开/编辑/另存、投影可读性、独立盲评 |

`examples/output/report/` 内含 `manifest.json`、`preview-spec.png`、`preview-pptx.png` 与 12 张逐页图。

### 5.3 优化前后对比（只列被测量的维度）

**判断（32 条离散判断，同一批用例）**

| 维度 | 无判断基线 | 本技能包 | 说明 |
|---|---|---|---|
| 主张来源 | 一律当作作者主张 | 7/7 正确区分 author / derived / absent | 基线会把话题标签当结论 |
| 证据形状 | 一律 proposition | 15/15 正确（含分母陷阱、单位、时间、顺序） | 基线无法区分 |
| 页面角色 | 一律 explain | 3/3 | 封面元信息不误判 |
| 素材判断（J30–J32） | 0/3 | 3/3（出图/不出图、职责、来源） | 基线只有关键词，无素材叙事 |
| 合计 | 7/32 (21.9%) | 32/32 (100%) | 用例集非独立盲标，只作回归 |

**中文排版（12 句 × 5 宽 × 3 字号 = 180 组）**

| 指标 | 朴素逐字断行 | 本技能包 | 变化 |
|---|---|---|---|
| 数字与单位被拆到两行的比例 | 5.71% | **0%** | −5.71pt |
| 行首出现禁则字符的比例 | 12.01% | **0%** | −12.01pt |
| 无标点可依、纯逐字断开的行占比 | 82.58% | **67.42%** | −15.16pt |

**素材生产说明（8 条意图 × 可判定条款）**

| 条款 | 朴素提示词 | 本技能包 | 说明 |
|---|---|---|---|
| 说清「为什么拍」（职责句） | 0/6 | **6/6** | 基线只写画面里有什么 |
| 说清「给谁让位」（安静面 + 调子） | 0/5 | **5/5** | 且不含任何坐标 |
| 说清「交付多大」（画幅 + 最小像素） | 0/8 | **8/8** | 496×624 → `4:5` / 992×1248 |
| 反向词只排生产失效 | 0/2 | **2/2** | 作者禁项逐条保留 |
| 条款合计（可比项） | 16/19 | **45/45** | 只说明「写得全不全」，不评图 |

**边界与可靠性（本轮新增，均可复跑）**

| 检查 | 结果 |
|---|---|
| 策略数值叶子的消费者覆盖 | 96 个数值叶子 → 0 个死配置（迁移前 6 个） |
| 阻断清单 ↔ 实际错误码 | 27 码双向一致（迁移前存在名字漂移与 2 个死码） |
| 策略枚举 ↔ Schema 同名枚举 | role / evidence / composition / density / emphasis / focus / media_* 全部一致 |
| 判断字段是否进执行 | execution 必读字段全部被消费；6 个审计字段显式登记 |
| 词表命中是否改变几何 | 命中只多一条待确认问题，元素几何逐项相同 |
| 产物写盘 | 中断不产生半截文件；demo.pptx 与迁移前字节一致 |

**执行侧（开发中实测）**

| 现象 | 修复前 | 修复后 |
|---|---|---|
| 同一 deck 两次编译哈希 | 不一致（内嵌 xlsx 时间戳） | 一致 |
| 复合槽位坐标 | 越出画布（14 个 OUT_OF_CANVAS） | 0 |
| 示例稿 release 结论 | BLOCK（单位缺失，且是误报） | PASS，且真正缺单位仍会被拦 |
| 预览与 PPTX 一致性 | 无机制 | 逐页像素对账，最差 0.0000 |
| 判断依据可见性 | 无 | 每页 `source / why / 命中片段 / open_questions` |

---

## 6. 未完成 / 未验证（必须如实列出）

1. **目标 PowerPoint 未验证**：本包不调用 PowerPoint，也未使用 LibreOffice（环境无 root，安装失败）。
   字体回落、真实行距、中文实际断行、投影可读性均未验证。
2. **独立视觉盲评未执行**：`evals/blind_pack.py` 已跑通打包与计分链路（链路自检、
   量表外分数拒绝、空表 `NOT RUN` 均已验证），但没有真实评审者回收的评分表。
   因此**不对视觉质量做任何优劣宣称**（构图、层级、节奏只剩人工判断）。
3. **判断用例集不是独立盲标**：与本包同步编写，看过即回归。100% 不代表泛化能力。
4. **上游仓库审计到函数清单级别**：读取了两个仓库的目录、SKILL/README/contract 与全部脚本的
   函数/常量大纲，未逐行审阅 2.25 万行实现；第 4 节的「不采纳」判断基于结构与契约层面，
   不针对具体实现细节。
5. **提示词层没有真实出图验证**：`pvad.py media` 能产出可执行的说明并有 45 条条款对照，
   但本包不调用图像模型，**没有一张图是按这些说明真的生成/拍摄出来的**。
   说明写到什么程度算够、模型是否照做，都未被验证（详见 §6.6）。
6. **素材（图片）链路已实现，但审美部分未盲评**：判断（`decision.media`）、几何（cover 裁切）、
   编译（原生 `p:pic` + 真裁切）、渲染（同一 `visible_image()`）、硬门（8 个 ASSET 码，含
   文字压图的**像素实测**对比度）、对账（sha256/尺寸/裁切）全部落地并有回归用例；
   但仍未做：PowerPoint 真机下图片渲染、多图跨页节奏（`max_per_page=2` 只有单图端到端证据）、
   裁切构图的盲评（「这张图裁得好不好」仍只能人判）。详见 §6.5。
7. **只覆盖一种画布尺寸（1280×720）与一套中性配色**：换画布、换品牌色需同步验证
   （配色只做了 token 校验与强调色预算，没有做对比度自动计算）。
8. **`editorial_text` / `structure_map` 两个配方没有被示例稿覆盖**，只有单测级的构造验证。
9. **边界测试覆盖的是当前策略**：死配置、阻断码、枚举一致性都是「此刻成立」的事实；
   新增策略键或错误码时若忘了同步，测试会红——但它无法判断「这个键本身该不该存在」。
10. **中文词组边界**：断行允许在词中间断（如「直营门／店」）。这是中文排版的合法行为，
   但读感一般；未实现词组级禁断（需要词典，收益/成本未评估）。

---

### 6.5 素材（图片）链路：本轮补齐的能力与仍未做的部分

用户指出的缺口是「图片控制的内容是否缺少」。审计确认属实，且存在一条**假通过**路径：
判断层可以引用素材、`compose` 可以不产 image 元素、`render` 可以不画图——三者任意一处缺失，
像素对照都会因为「两边都没有图」而判 `MATCH`。本轮把整条链路补成「缺一处就报错」。

| 环节 | 实现 | 失败即暴露 |
|---|---|---|
| 判断 | `judgment._judge_media`：声明优先（`required`/`existing`），未声明只产 `ASSET_HINT` 问题 | 词表命中不会偷偷出图 |
| 职责 | `decision.media{decision,role,function,reason,source}` | `ASSET_DECLARATION` / `ASSET_REASON_MISSING` |
| 几何 | `compose._compose_media`：cover 裁切比例、`source_zone`、缺文件也写元素并标 `_missing` | `ASSET_MISSING`（缺图不静默） |
| 编译 | `compile._add_image`：原生 `add_picture` + 真实 `crop_left/top/right/bottom`；压图用真半透明矩形 | `fidelity` 报 `crop` 未应用 |
| 渲染 | `render.visible_image()`：渲染与度量共用同一贴图口径，缺图画红框叉 | 「两张不同图产生 >1% 像素差」用例 |
| 硬门 | `assets.py` 度量 + `guard` 8 个 ASSET 码 + `SOURCE_ZONE_COLLISION` + 路径越界告警 | 格式（按真实解码格式，非后缀）、分辨率、裁切、安全区、对比度、预算 |
| 对账 | `fidelity` 比对 `type/sha256/source_px/crop` | 包里换了一张图也能发现 |

**关键判断（写进 SKILL §5）**：素材是「判断」不是「装饰」。出图必须由作者声明，
职责必须能写成一句话（「拿掉它，什么信息会损失」），文字压图必须按观众看见的那块像素实测对比度，
裁切必须是真裁切而不是拉伸。**宁可不出图，也不出假图 / 装饰图 / 证明不了事的图。**

**仍未做（不在本轮宣称范围内）**：PowerPoint 真机下的图片渲染与投影可读性；
多图同页与跨页素材节奏的端到端证据（策略上限 2，示例只用 1 张）；
裁切构图质量（主体是否被切、裁完是否好看）的盲评；
素材来源授权与被摄者同意的流程化（目前只做到「路径越界告警」与「生成图不得当事实」）。

---

### 6.6 提示词层：素材「拍什么」的说明（本轮新增）

用户指出的缺口是「缺少控制图片质量的提示词」。上一轮补齐了「图在不在、够不够用」，
但**「这张图该拍什么」没有落点**：作者只写一个路径，判断层通过、QC 通过，生产端拿到的仍是空白。
本轮补的是这一层，并且刻意只补这一层——本包不调用图像模型，也不冒充它会。

| 环节 | 实现 | 边界 |
|---|---|---|
| 语言 | `policy/prompts.yaml`：职责句、安静面句、调子句、生产尾段、反向词、画幅吸附 | 只放**短语与顺序**，不含判断 |
| 翻译 | `scripts/prompt.py`：读已结算的 `decision.media` + 图位几何 + `safe_area` | 不改判 function/role/reason；不猜主体 |
| 触发 | `compose` 只在**文件缺失/解码失败**时写入 `brief`；已有素材不出提示词 | 提示词是生产说明，不是评审作文 |
| 导出 | `pvad.py media deck.json --out media_briefs.json` | 一条命令拿到 prompt/negative/画幅/目标像素/指纹 |
| 自审 | `validate_semantics` 检查短语表与枚举一一对应、顺序覆盖、策略自身不得含风格词 | 有枚举没短语 = 静默丢一句职责 |
| 证据 | `tests/test_prompt.py` 28 例 + `evals/prompt_eval.py`（45 条条款，朴素基线 16/19） | 只证明说明可执行，不证明图好看 |

**从上游 `references/assets.md` / `scripts/assets.py` 采纳的思想**（重写，非搬运）：
清单驱动（`asset_manifest` → 本包的 `media_briefs.json`）、提示词是决策后的**确定性翻译**、
「反向词只排除生产失效」、紧凑质量尾段、指纹（内容没变就不重出图）。

**刻意不采纳**：`medium / lighting / material / texture / fusion` 六层风格词叠加——
本包的图位是文稿的一部分，风格由作者的 `subject` 决定；策略层自动叠加风格正是「行业词表导致
视觉趋同」的来源。本包只用 `style_words_banned` 反过来审查自己有没有偷渡风格词。

**明确未验证**：这些提示词在真实图像模型上的出图效果、成图与说明的一致性、跨模型可移植性。
本环境没有调用任何图像模型（也不该调）：`prompt_eval` 只度量说明是否可执行，
图的优劣仍只有人工盲评（§7 P0-2）。要宣称「提示词让图更好」，必须先在真实模型上做对照盲评。

---

### 6.7 边界治理：这一轮到底改了什么（P1）

审计方法不是读代码凭感觉，而是三条机械检查——它们本身就是交付物的一部分
（`tests/test_boundaries.py`，19 例）：

1. **策略必须有人消费**：遍历 `strategy.yaml` 的每个数值叶子，在 `scripts/` 里找不到消费者的即失败；
2. **阻断清单必须与错误码一致**：扫 guard 源码里 `rep.error("…")` 的全部码，与
   `verification.blocking_codes` 双向比对；
3. **枚举只能有一份来源**：策略枚举与 `plan.schema.json` ／ `brief.schema.json` 的同名枚举逐项相等。

跑出来的真实问题（不是假想）：

| 层 | 问题 | 属于哪类越界 | 处理 |
|---|---|---|---|
| Python | `judgment.MUST_NOT`：9 个配方的禁令清单写在代码里 | 设计判断混进执行层 | 迁到 `composition.yaml` |
| Python | `judgment._ladder_for` / `compose._claim_ladder`：同一字阶映射写了两份 | 重复实现 | 只在配方里留一份 |
| Python | `judgment._derive_world`：兜底视觉世界的中文文案 | 设计语言在代码里 | 迁到 `judgment.default_world` |
| Python | `assets.MAX_PIXELS` / `_SAMPLE_MAX_SIDE` / `fidelity.SIZE_PT_TOL` | 阈值硬编码 | 全部进策略 |
| Python | `typography._BREAK_AFTER_PUNCT` + 55% 行宽 | 可调规则藏在常量里 | 进 `typography.line_break.*` |
| Python | `guard.BLOCKING` + 名字漂移（`OVERLAP` vs `ELEMENT_OVERLAP`） | 策略语义写在代码里的典型后果 | 进策略并加一致性测试 |
| 策略 | 6 个死配置（对比度、安全区宽度、平坦占比、usage 码、2 个重复键） | 写了没人执行 = 两套真相 | 4 个接线，3 个删除 |
| 判断 | `emphasis=lead` 只进 meta，不影响任何执行 | 判断未进入执行流程 | 接成字号档 + 装不下时诚实降级 |
| 执行 | PPTX 直接写目标文件 | 交付可靠性 | 临时文件 + 原子替换 |

**没做的事也写清楚**：没有为了「分层好看」拆模块（`judgment.py` 仍是 700+ 行，只在
追加用例暴露耦合时再拆）；没有引入 `medium/lighting/texture` 那类风格层（会造成同质化）；
没有把测试数量当成质量结论。

---

## 7. 已知风险与后续优先级

**P0（影响可信度，必须补）**
1. 在真实 PowerPoint（Windows/macOS）打开、编辑、另存、重开，记录字体回落与断行差异，
   把结论写进 manifest 的 `target_powerpoint`。
2. 组织 ≥2 名评审者的盲评：候选 vs 无判断基线（用 `blind_pack.py`），把分差写进报告；
   若分差 < 0.5，如实写「未观察到差异」。

**P1（影响覆盖面）**
3. 素材链路的剩余部分：PowerPoint 真机图片渲染、多图跨页节奏的端到端样例、
   裁切构图的人工/盲评（像素 QC 已实现并有 27 条用例，不再重复列入）。
4. 提示词层的外部验证：拿 `media_briefs.json` 在 1–2 个真实图像模型上出图，
   用 `blind_pack.py` 做「按说明出图 vs 一句主体」的对照盲评；结果未回收前不得宣称提示词提升出图质量。
5. 画布与品牌适配：第二套画布尺寸 + 品牌色下的对比度自动校验（`contrast` 阈值已在策略里，未接线）。
6. 把示例稿扩到覆盖全部 8 个配方，并补 3–5 个跨行业 brief（避免只有一套叙事节奏）。

**P2（工程质量）**
7. 判断用例继续扩充并冻结版本（每次扩充后 `--ablate` 必须无未覆盖项）。
8. `judgment.py`（773 行）可按「主张 / 证据 / 角色 / 构图」拆分，但**只在追加用例暴露耦合时**做，
   不为目录好看拆分。
9. Windows/macOS 上的字体探测路径实测（当前只验证了 Linux 的 Noto CJK）。

---

## 8. 验收标准自检

| 标准 | 状态 | 依据 |
|---|---|---|
| 设计判断更透明 | ✅ | 每条判断带 source/why/命中片段；提示词的每句话带 `slot`，能追到策略条目 |
| 执行链路更确定 | ✅ | 三步一趟；字节确定；失败给根因组；指纹拒绝混用 |
| 模块职责无明显越界 | ✅ | 阈值零硬编码 + 死配置为 0（`test_boundaries`）；设计禁令/字阶/焦点在配方里；阻断清单与错误码双向一致 |
| 设计判断能够影响实际执行 | ✅ | `emphasis=lead` → 字号档上移（有断言）；装不下时 `emphasis_notes` 记录；`must_not` 逐字来自策略 |
| 真实 PPTX 产物可验证 | ✅ | 原生文本/图表/图片 + 41/41 回读对账 + 像素 0.0000 + 哈希 |
| 质量提升有对照证据 | ⚠️ 部分 | 判断（+78.1pt）、排版（三项指标归零）、提示词说明（45/45 vs 16/19）有对照；**视觉质量与出图效果无对照，未做盲评** |
| 整体复杂度得到控制 | ✅ | 生产代码 5088 行完成 plan→build→check/media/briefs；本轮净增 183 行（其中 60 行是把设计表从 Python 搬到策略）；未新增分支树，未拆模块 |

**总评：可交付的 v1.0；视觉终验与盲评是交付后的第一优先级，不作为「已完成」宣称。**
