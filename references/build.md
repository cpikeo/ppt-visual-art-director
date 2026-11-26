# 机器契约：plan、编译、核验、定点编辑

字段形状以 `scripts/plan.schema.json` 为准（`scaffold.py fields <节点>` 查紧凑字段表），技术阈值以 `scripts/limits.json` 为准。本文只讲 schema 表达不了的语义、坑与命令。

## 必填与枚举速查（写 plan 之前先对一遍）

schema 校验是硬阻断（`E-SCHEMA`），下面这些**在文档正文里读不到、但一次写错就整份被拒**：

| 位置 | 必填 | 类型／枚举的坑 |
| --- | --- | --- |
| 根 | `deck_id` `title` `audience` `goal` `core_claim` `brief` `pages` | 画布固定 16:9（1280×720 px），无 `aspect_ratio` 字段；旧稿带该字段先 `scaffold.py migrate` |
| `brief` | `insight` `concept` `veto` | `veto` 与 `motif` 都是**数组**（`motif` 写成字符串会被拒）；`art` 要填就五个键 `light` `lens` `material` `gamut` `grain` 齐填 |
| `page` | `id` `role` `focus` `elements` | `role` 只有六个值：`open` `build` `breathe` `turn` `hold` `close`（叙事动作，不是自定义标签；页面职责写进 `composition_rationale`）；`evidence` 是**字符串数组**；`rhythm` 条数要等于页数 |
| `page_numbers.hide` | — | 是**页 id 的字符串数组**，不是页码数字 |
| `element` | `id` `type` | `additionalProperties: false`：拼错的键直接被拒；`x y w h` 除 `path`／`group` 外都要写 |
| `chart` | `kind` `data` | `table` 的数据是 `data: {headers, rows}`，不是 `series`；`scatter` 用 `series[].points = [[x,y],…]` |

`brief.editorial`（`amplify` `recede` `cut`，字符串或字符串数组均可）与 `brief.system`（`typography` `color` `space`）是可选的决策记录，schema 只校验形状；页 `depth`（`background` `midground` `foreground`，元素 id 数组）有引用与顺序校验。

**内容不静默丢失（`E-DROP`）**：`init` 对非 `free` 页逐字核对 `title` `sub` `kicker` `meta` `eyebrow` `body` `note` `aside` 是否进入版面，并拒绝配方不认识的页字段（拼错、写给另一个配方的键）。口径、示意图声明这类限定语被丢掉比版式难看更糟。

## 用 scaffold 起稿（默认路线）

`scaffold.py` 只做算术，不做判断：**色彩系统**（种子色 → CIELAB 推出环境/结构/文字/强调/数据色，正文 ≥ 7:1、注释 ≥ 4.8:1、强调色 ≥ 4.5:1，数据色两两灰度差 ≥ 9.5 L*，深浅两套）、**字号阶梯**（`hall`／`meeting`／`screen`）、**12 栏网格**（边距 x 0.8 / y 0.7，沟槽 0.25）、**框高**（`model.fit_height_in`）、**自适应大字号**（`fit_pt`：按 86% 宽度估算也不折行的最大字号，渲染实测大字比估算宽 10–15%）、**跨页母题**（默认不画；`design.motif: "bar"` 在每页 (0.8, 0.7) 处画 0.6×0.07 in 强调色短横线）。

```bash
python scripts/scaffold.py system --seed "#B4421E" --mood warm --scene meeting   # 看色彩系统
python scripts/scaffold.py init outline.json plan.json                           # 生成 plan 初稿
```

outline = plan 的根字段（`deck_id` `title` `audience` `goal` `core_claim` `brief` `assets`…原样带入）＋`design`＋`pages`：

```json
{"design": {"seed": "#B4421E", "mood": "warm", "scene": "meeting",
            "fonts": {"heading": "Noto Serif CJK SC", "body": "Noto Sans CJK SC"}},
 "pages": [{"recipe": "big_number", "title": "…", "number": "41", "unit": "个百分点", "unit_scale": 0.26,
            "body": "…", "note": "口径：…", "evidence": ["…"], "beat": "这页做什么、多重"}]}
```

| `recipe` | 默认 `role` | 必填 | 可选 | 讲法 |
| --- | --- | --- | --- | --- |
| `cover` | open | `title` | `sub` `kicker` `meta` `image` `dark` | 标题底对齐锚在 64% 线，字号随标题长度取不折行最大值，上方整块安静区；有图且没写 `dark` 时按文字区真实像素明暗取墨色（`init --assets`）；折行标题自动平衡行宽，不留一两字的孤行 |
| `statement` | hold | `title` | `sub` `dark` | 一句话占满视野，左侧一道竖线是唯一装饰 |
| `section` | breathe | `title` | `number` `dark` | 暗场 + 巨型序号，换气与分章 |
| `big_number` | build | `title` `number` `body` | `unit` `unit_scale` `note` `eyebrow` | 数字取放得进 7 栏的最大字号，正文被细线隔在右栏 |
| `chart` | build | `title` `chart` | `aside` `note` `highlight` `focus_series` `eyebrow` | 图承担读数；`highlight`=类别序号（单系列柱/条，其余退隐）；`focus_series`=系列序号（柱/线/面积，多系列）：被论证的系列用强调色，其余退为灰，写入 `color_map`（plan 里已有 `color_map` 则不覆盖） |
| `compare` | turn | `title` `left` `right`（各 `label` `number` `body`，可加 `value`） | `scale` `note` `eyebrow` | 两数同字号同基线；两侧都有 `value` 时再画自零起的比例条。`scale: "value"`（两侧 `value` 均 > 0）＝**尺度隐喻**：数字字高与 `value` 成正比、共用一条基线，较大者占 8 栏，较小者字号不低于 28pt（再小不再成比例）；自动写入 `temperament: dramatic`、`hero_ratio`、`focus_element`。落差是论点时用它，两数本应等重时不用 |
| `list` | build | `title` `items` | `list`（number/bullet/dash）`note` `eyebrow` | 标题与原生列表同一顶线，竖线钉住两栏 |
| `table` | build | `title` `table`（headers/rows） | `note` `eyebrow` | 原生表格 |
| `steps` | build | `title` `steps`（2–5 项，各 `label`，可加 `when` `body`） | `note` `eyebrow` | 只表达先后，不画箭头（箭头暗示因果与时长） |
| `split_image` | build | `title` `image` | `body` `side` `ratio` `note` `eyebrow` | 图占一侧（`side` left｜right；`ratio` 图宽占页宽 0.3–0.65，缺省 0.5）；标题压顶、正文压底，两者间留 ≥0.4in，放不下先收标题字号 |
| `close` | close | `title` | `items` `dark` | 暗场收束到决定与下一步 |
| `free` | build | `title`（页题记录）`els` | `grid` `dark` `bg` `focus_element` `depth` | **自由构图**：逐个写元素，系统只供网格、色值与字号档（见下） |

**`design` 的判断出口**（默认推导只是起点）：`ground.light/dark` 页底色（纸色、墨绿、近黑；明度不对会 `E-OUTLINE`）、`accent.light/dark` 唯一强调色、`data` / `data_dark` 数据色序列、`motif: "none"` 关闭每页短横线（单页用 `motif: false`）。覆盖后文字墨阶按新底色重新推导，仍保证 ≥ 7:1 / 4.8:1。

**`free` 的页级网格**：`grid` 缺省 12 栏 × 6 行（边距 x 0.8 / y 0.7，沟槽 0.25 / 0.2）；写整数 N 即 N 栏（如 5、8）；写对象 `{cols, rows, margin, margin_y, gutter, gutter_y}` 自定；写 `"none"` 则本页不用网格（再用 `cols`/`rows` 报 `E-OUTLINE`）。网格是选出来的构图工具，不同页可以不同。

**`free` 的元素速记**（都可选，写了才换算，其余键原样交给 schema）：`cols:[起列,跨列]`→`x w`（按本页网格）；`rows:[起行,跨行]`→`y h`；`right:x`→以该 x 为右边；`size`＝档位名（`title body number caption hero`）或 pt；`fit:[最大pt,最小pt]`＝取不折行的最大字号；`ink`/`color`/`fill` 可写调色名（`ink ink2 muted quiet accent rule bg bg_alt`，组内子元素同样生效）；`bottom:y` 以该 y 为文字底边；`after:元素id`+`gap` 紧接上一元素下方；文字缺 `h` 由 `fit_height_in` 估出。要出血或错网格，直接写 `x w`。页级 `bg` 可写调色名或 `#hex`（如整页强调色场）。

每页的 `eyebrow` 是眉题（小号、位于母题横线右侧）。outline 缺字段或 `recipe` 写错时，`scaffold.py init` 以 `E-OUTLINE` 逐条说明并退出 2，不写文件。

每页可覆盖：`role` `claim` `focus` `rationale` `composition` `design_state` `beat` `speaker_notes` `evidence` `open_questions` `dark` `focus_element` `depth`。`image` 写 `{id, role, info_duty, focus:{x,y}, reason, duotone}`，图须已在 `assets` 登记。生成的 plan 是起点：之后照常 compile → verify → 看图，再手改文案与几何；配方里没有的讲法，手写该页元素即可，网格函数 `col(start, span)` 与色彩系统照用。

## 设计意图字段与状态（schema_version 2）

| 字段 | 位置 | 含义 | 谁来校验 |
| --- | --- | --- | --- |
| `schema_version` | 根 | 缺省=1（旧稿）；当前 2。旧稿 `scaffold.py migrate` 升级，不就地改写 | schema 枚举 |
| `design_state` | 页 | `draft` 待判断／`judged` AI 已判断／`confirmed` 人已确认；缺省按 judged（旧稿） | `W-PENDING`；`--production` 下 `E-PENDING` |
| `composition.strategy / focus_mode / grid / axes / exception / source / recipe` | 页 | 构图语法、焦点策略（single dual multi none）、网格规格、对齐轴、破格理由、来源（recipe/author/user）。**语义记录**，机器不评价好坏 | 仅一致性：`W-FOCUS` |
| `composition.temperament / negative_space{role,where} / hero_ratio` | 页 | 空间气质（7 选 1）、空白的功能（5 选 1）、主角对最强配角的最低视觉重量倍数。`hero_ratio` 抬高 `W-DOMINANT` 阈值；`temperament` `negative_space` 由 `verify.py` 对真实渲染量测（`V-SPACE` `V-RHYTHM`）；`negative_space.role=focus` 无 `focus_element` → `W-FOCUS` | 编译：`W-DOMINANT` `W-FOCUS`；核验：`V-SPACE` `V-RHYTHM` |
| `concept` | 页 | 一句话视觉概念（这页的尺度/视角/裁切在论证什么），先于构图填写；机器不评价 | — |
| `composition.reading_path` | 页 | 阅读顺序（元素 id） | `E-REF` 引用必须存在 |

作者没写 `focus` 则为占位符 `【待判断】`、页标 `draft`；`focus` 与 `rationale` 都写了才标 `judged`。`focus_mode: none` 与 `focus_element` 同时出现 → `E-REF`；`single` 却没有 `focus_element` → `W-FOCUS`。`W-SUM`：以 % 表示的饼/环合计不是 100；合计为 0 → `E-CHART`。`W-DOMINANT` 等只是面积/字号的粗略线索，不是视觉质量判定。

元素坐标、文本、叠放形状、素材与生成图、图表、页码的字段语义见 `elements.md`（手写元素、图表、素材登记时读；只用配方起稿可不读）。
