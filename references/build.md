# 机器契约：plan、编译、核验、定点编辑

字段形状以 `scripts/plan.schema.json` 为准（`scaffold.py fields <节点>` 查紧凑字段表），技术阈值以 `scripts/limits.json` 为准。本文只讲 schema 表达不了的语义、坑与命令。

## 必填与枚举速查（写 plan 之前先对一遍）

schema 校验是硬阻断（`E-SCHEMA`），下面这些**在文档正文里读不到、但一次写错就整份被拒**：

| 位置 | 必填 | 类型／枚举的坑 |
| --- | --- | --- |
| 根 | `deck_id` `title` `audience` `goal` `core_claim` `brief` `pages` | `aspect_ratio` 只有 `16:9`／`4:3` |
| `brief` | `insight` `concept` `veto` | `veto` 与 `motif` 都是**数组**（`motif` 写成字符串会被拒）；`art` 要填就五个键 `light` `lens` `material` `gamut` `grain` 齐填 |
| `page` | `id` `role` `focus` `elements` | `role` 只有六个值：`open` `build` `breathe` `turn` `hold` `close`（叙事动作，不是自定义标签；页面职责写进 `composition_rationale`）；`evidence` 是**字符串数组**；`rhythm` 条数要等于页数 |
| `page_numbers.hide` | — | 是**页 id 的字符串数组**，不是页码数字 |
| `element` | `id` `type` | `additionalProperties: false`：拼错的键直接被拒；`x y w h` 除 `path`／`group` 外都要写 |
| `chart` | `kind` `data` | `table` 的数据是 `data: {headers, rows}`，不是 `series`；`scatter` 用 `series[].points = [[x,y],…]` |

`brief.editorial`（`amplify` `recede` `cut`）与 `brief.system`（`typography` `color` `space`）是可选的决策记录，schema 只校验形状；页 `depth`（`background` `midground` `foreground`，元素 id 数组）有引用与顺序校验。

六种 `role` 的叙事动作与重量见 `design.md` 三（叙事）；写不出对应叙事动作，说明页序还没定。

## 用 scaffold 起稿（默认路线）

`scripts/scaffold.py` 只做算术，不做判断：**色彩系统**（种子色 → CIELAB 推出环境/结构/文字/强调/数据色，保证正文 ≥ 7:1、注释 ≥ 4.8:1、强调色 ≥ 4.5:1，数据色两两灰度差 ≥ 9.5 L*，深浅两套）、**字号阶梯**（`hall`／`meeting`／`screen`）、**12 栏网格**（边距 x 0.8 / y 0.7，沟槽 0.25，左轴固定）、**框高**（直接调用 `model.fit_height_in`，折行按 98% 宽度估）、**自适应大字号**（`fit_pt`：取按 86% 宽度估算也不折行的最大字号——渲染实测大字真实宽度比估算宽 10–15%）、**跨页母题**（默认不画；`design.motif: "bar"` 恢复 15.1 的每页 (0.8, 0.7) 处 0.6×0.07 in 强调色短横线）。

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
| `chart` | build | `title` `chart` | `aside` `note` `highlight` `eyebrow` | 图承担读数；`highlight`=类别序号（单系列柱/条，其余退隐） |
| `compare` | turn | `title` `left` `right`（各 `label` `number` `body`，可加 `value`） | `note` `eyebrow` | 两数同字号同基线；两侧都有 `value` 时再画自零起的比例条 |
| `list` | build | `title` `items` | `list`（number/bullet/dash）`note` `eyebrow` | 标题与原生列表同一顶线，竖线钉住两栏 |
| `table` | build | `title` `table`（headers/rows） | `note` `eyebrow` | 原生表格 |
| `steps` | build | `title` `steps`（2–5 项，各 `label`，可加 `when` `body`） | `note` `eyebrow` | 只表达先后，不画箭头（箭头暗示因果与时长） |
| `split_image` | build | `title` `image` | `body` `side` `eyebrow` | 半幅图，标题压顶、正文压底 |
| `close` | close | `title` | `items` `dark` | 暗场收束到决定与下一步 |
| `free` | build | `title`（页题记录）`els` | `grid` `dark` `bg` `focus_element` `depth` | **自由构图**：逐个写元素，系统只供网格、色值与字号档（见下） |

**`design` 的判断出口**（默认推导只是起点）：`ground.light/dark` 页底色（纸色、墨绿、近黑；明度不对会 `E-OUTLINE`）、`accent.light/dark` 唯一强调色、`data` / `data_dark` 数据色序列、`motif: "none"` 关闭每页短横线（单页用 `motif: false`）。覆盖后文字墨阶按新底色重新推导，仍保证 ≥ 7:1 / 4.8:1。

**`free` 的页级网格**：`grid` 缺省 12 栏 × 6 行（边距 x 0.8 / y 0.7，沟槽 0.25 / 0.2）；写整数 N 即 N 栏（如 5、8）；写对象 `{cols, rows, margin, margin_y, gutter, gutter_y}` 自定；写 `"none"` 则本页不用网格（再用 `cols`/`rows` 报 `E-OUTLINE`）。网格是选出来的构图工具，不同页可以不同。

**`free` 的元素速记**（都可选，写了才换算，其余键原样交给 schema）：`cols:[起列,跨列]`→`x w`（按本页网格）；`rows:[起行,跨行]`→`y h`；`right:x`→以该 x 为右边；`size`＝档位名（`title body number caption hero`）或 pt；`fit:[最大pt,最小pt]`＝取不折行的最大字号；`ink`/`color`/`fill` 可写调色名（`ink ink2 muted quiet accent rule bg bg_alt`，组内子元素同样生效）；`bottom:y` 以该 y 为文字底边；`after:元素id`+`gap` 紧接上一元素下方；文字缺 `h` 由 `fit_height_in` 估出。要出血或错网格，直接写 `x w`。页级 `bg` 可写调色名或 `#hex`（如整页强调色场）。

每页的 `eyebrow` 是眉题（小号、位于母题横线右侧）。outline 缺字段或 `recipe` 写错时，`scaffold.py init` 以 `E-OUTLINE` 逐条说明并退出 2，不写文件。

每页可覆盖 `role` `claim` `focus` `rationale` `composition` `design_state` `beat` `speaker_notes` `evidence` `open_questions` `dark` `focus_element` `depth`。`image` 写 `{id, role, info_duty, focus:{x,y}, reason, duotone}`，图须已在 `assets` 登记。生成的 plan 是起点：之后照常 compile → verify → 看图，再手改文案与几何；配方里没有的讲法，手写该页元素即可，网格函数 `col(start, span)` 与色彩系统照用。

## 设计意图字段与状态（schema_version 2）

| 字段 | 位置 | 含义 | 谁来校验 |
| --- | --- | --- | --- |
| `schema_version` | 根 | 缺省=1（旧稿）；当前 2。旧稿 `scaffold.py migrate` 升级，不就地改写 | schema 枚举 |
| `design_state` | 页 | `draft` 待判断／`judged` AI 已判断／`confirmed` 人已确认；缺省按 judged（旧稿） | `W-PENDING`；`--production` 下 `E-PENDING` |
| `composition.strategy / focus_mode / grid / axes / exception / source / recipe` | 页 | 构图语法、焦点策略（single dual multi none）、网格规格、对齐轴、破格理由、来源（recipe/author/user）。**语义记录**，机器不评价好坏 | 仅一致性：`W-FOCUS` |
| `composition.reading_path` | 页 | 阅读顺序（元素 id） | `E-REF` 引用必须存在 |

scaffold 不再把配方说明写进 `composition_rationale`、不再用标题推断 `focus`：作者没写，`focus` 为占位符 `【待判断】`、页标 `draft`；`focus` 与 `rationale` 都写了才标 `judged`。`focus_mode: none` 与 `focus_element` 同时出现 → `E-REF`；`single` 却没有 `focus_element` → `W-FOCUS`。`W-SUM`：以 % 表示的饼/环合计不是 100。`free` 页元素 id 不再与母题保留 id `bar` 冲突；调色名在 `runs`/`stroke`/`gradient` 里同样生效。`W-DOMINANT` 等仍是面积/字号的粗略线索，不是视觉质量判定。

## 坐标与元素

- `x y w h` 为英寸，原点左上，y 向下；`font_size` 为 pt。16:9 = 13.333 × 7.5；4:3 = 10 × 7.5。
- 元素类型：`title` `body` `text` `number` `caption` `footer`（文字），`image` `chart` `legend` `shape` `divider` `path` `group`。
- **页标题**：页内第一个顶层 `title` 元素编译为版式（Title Only）的原生标题占位符，位置与样式显式写出，不继承母版；其余 `title`、组内 `title` 与其它文字是普通文本框。页内没有 `title` 时不留空占位符。要让大纲、读屏、幻灯片导航读到页标题，每页写一个 `title`。
- `path` 与 `group` 不写 `x y w h`（由点集或子元素推出）。元素 `id` 页内唯一；成品对象名为 `<页 id>.<元素 id>`，页码对象为 `_page_number`。
- 允许出血；文字与图表伸出画布会被裁（编译期提示）；完全在画布外的元素阻断。

## 文本

- `content` 用 `\n` 分段。同一行内换字重、字号、颜色或字族用 `runs`；`runs` 拼接必须与 `content` 逐字相等。不要把一个句子拆成两个框。
- 文本自带内边距（左右 0.05、上下 0.03 in）。要贴住网格轴线（大标题、数字）写 `pad_in: 0`。
- 默认：`title` 40pt 粗、`number` 72pt 粗、`body` `text` 20pt、`caption` `footer` 13pt；行距 1.12（标题、数字）、1.32（正文）、1.25（注释）。深底（亮度 < 0.18）未声明颜色时用浅墨。
- 字族三分：`fonts.heading`、`fonts.body`（正文、图表、页码）、`fonts.number`（缺省同 heading）。
- `letter_spacing_pt` 范围 −40…40：几百 pt 的巨型数字需要按字号比例收紧字距（约 −2%…−4% 字号），旧版 −4 的上限无法表达。
- `rotation` 对所有对象生效，按 360 取模。
- **原生列表**：`list: "bullet"|"dash"|"number"`，每个 `\n` 分段是一项，写成真正的项目符号/自动编号（`buChar`／`buAutoNum`）与 0.32 in 悬挂缩进，符号与第一段文字同色；不要手敲“1.”“•”。估宽按缩进后的版心算；`verify.py` 逐项比对渲染文字（渲染器会在每项前插入符号）。
- 页面 `speaker_notes` 写入备注页（页面 `notes` 是设计备注，不进成品）。元素 `alt` 写入替代文字：图片缺省取 `asset_reason`，图表缺省取图题，装饰元素不写。两者都由 `verify.py` 逐项核对。
- `W-FIT` 的行高按字体取值（`limits.json` 的 `type.line_pitch_em.by_font`）：Noto/思源 CJK 1.448、雅黑 1.32、宋体/黑体 1.15，未知字体 1.25。用别的字体且估不准，在 `by_font` 里补一行。
- 图表与表格内的字号由 `limits.json` 统一给出（表 18、图题 15、图内标注/轴/图例 14）；要改，改阈值，不逐图改。

## 叠放、形状、路径、组

- **叠放顺序就是声明顺序**，先写的在下。承载层（底图、整幅 `hero` 主视觉、蒙版、色块、做底的巨型数字或标题）要有东西压在上面；`hero` 作承载层时，文字必须落在简报声明的安静区，由 `W-CONTRAST` / `W-BUSY` 量测，`subject` / `illustration` / `icon` / `detail` 不能被压，在自己身上声明 `overlap_allowed: true`；`number` 与 `title` 可以做承载层，`body`、`caption`、`footer` 与图表不能自我豁免。实质重叠（面积 ≥ 0.5 in²、短边 ≥ 0.05 in、占较小者 ≥ 25%）且下层未声明即阻断（`E-OVERLAP`）。
- `shape_kind`：`rect` `round_rect` `oval` `triangle` `diamond` `chevron` `arrow_right` `arrow_up`。`corner_in` 仅用于圆角矩形（半径，英寸）。`fill_alpha` 低于 0.15 基本不可见。
- `gradient`：`{from, to, angle, from_alpha?, to_alpha?}`，`angle` 0 为左到右、90 为上到下；末端透明度应接近 0。`stroke`：`{color, width_pt, dash, cap}`。
- 路径顶点为坐标 `[x, y]` 或锚点 `{"el": "n1", "anchor": "right", "at": 0.5}`。两点为直线，三点以上为折线；`closed: true` 加 `color` 为填充。锚点在编译期解析，改形状要重编译；旋转组内的锚点会错位，连线应同组或改用坐标。
- 组内子元素用页面绝对坐标，组内叠放不查重叠。

## 素材与生成图

素材从素材根解析（`compile` `verify` `revise` 共用同一规则；`--assets` 优先，其次 plan 同级 `assets/`，再次 plan 所在目录）：先按 `image_id` 精确找，再补 `.png` `.jpg` `.jpeg` `.webp`，不做同名不同扩展名的兜底。**结果必须留在素材根内**：绝对路径、含 `..`、经符号链接指向根外的文件一律视为找不到（`E-IMG`）。图片按 `cover` 裁切，`crop_focus` 为裁窗中心。

**三个维度各归各的字段，互不替代**（详见 `imagery.md` 〇）：

- **素材来源**：`assets.<id>.source`、`generated`／`generation`、`rights`、`used_as_evidence`。
- **资产角色** `image_role`：`background` `texture`（承载层，可声明 `overlap_allowed`）、`hero`（主视觉）、`subject`（人物或产品）、`illustration`、`icon`、`detail`；旧写法 `evidence` 仍接受，等价于 `info_duty: evidence`。
- **信息职责** `info_duty`（元素级，仅 `image`）：`evidence` `structure` `identity` `atmosphere` `mood` `material` `none`。只有 `evidence` 带来源约束——不能指向生成图（`E-ASSET`）；不声明给 `W-ASSET`；`none`（纯装饰）也给 `W-ASSET`，须说明留下的理由。

每个素材都登记在 `assets`。**生成与否只由 `generation` 决定**：

```json
"assets": {
  "hero":  {"source": "自有素材库 m-01", "rights": "cleared", "reference": "编号 m-01"},
  "plate": {"source": "自建生成", "rights": "pending", "generated": true,
            "generation": {"tool": "图像生成工具", "prompt": "完整提示词",
                           "file": "plate.png", "status": "generated", "sha256": "64 位十六进制",
                           "brief": {"intent": "承担什么职责", "quiet_zone": "左侧 65% 连续低纹理，明度不低于 70%", "reject": ["可观察的淘汰条件"]}}}
}
```

- `source` 必填。`rights`：`cleared`／`pending`（提示）／`blocked`（阻断）。
- `generated` 只能是对 `generation` 是否存在的声明。`generated: false` 却有记录，或 `generated: true` 却无记录，都阻断（`E-ASSET`）。记录缺 `tool` `prompt` `file` `sha256` 任一即阻断；`status` 须为 `generated`；`file` 须与 `image_id` 解析到同一文件；`sha256` 须与文件实际哈希一致。
- `generation.brief`（`intent` `quiet_zone` `reject`）是出图前的艺术指导简报；缺失给 `W-ASSET`。`quiet_zone: "none"` 表示不叠字，若仍有文字压在图上给 `W-ASSET`。
- `used_as_evidence: true` 的素材不能是生成图。
- `duotone: {dark, light}` 把图压进色域，原生可编辑；PowerPoint 生效，LibreOffice 只去饱和。
- 图压着字的量测（`check_text_on_image`）：按裁切、`duotone`、`alpha` 还原文字框正下方的像素，算达标占比与纹理；图与文字之间隔着色块或渐变时不适用，退回 `I-CONTRAST`。它是静态近似，不替代渲染。
- 核验器比对 PPTX 内嵌字节与素材原件：像素尺寸不同阻断，字节不同（可能重编码）提示。

## 图表

**有效开关只有一处**：`model.chart_flags(chart)`。plan 里没写 `legend`／`data_labels` 时，饼/环 `data_labels` 缺省开、`legend` 缺省关；柱线面散 `data_labels` 缺省关，`legend` 在两个以上系列时开；表格恒关。compile、checks、verify 一律读它，显式写出的值（含 `false`）优先。读数闭环（每个扇区认得出是什么）由 `model.chart_readout_problems` 判定，checks 据此阻断，verify 另从成品 XML 复核（`V-READOUT`）。

```json
{"kind": "bar", "horizontal": true, "data_labels": true, "axis_labels": false,
 "title": "自然搜索撑住了留存", "unit": "%", "period": "2025 全年", "basis": "次月留存",
 "data": {"categories": ["自然搜索", "信息流"], "series": [{"name": "次月留存", "values": [52, 31]}]}}
```

| `kind` | 修饰 | 得到什么 |
| --- | --- | --- |
| `bar` | `horizontal`、`stacked`（`true`／`"percent"`）、`combo_line_series` | 竖柱、横条、堆叠、百分比堆叠、柱＋线组合 |
| `line` | — | 等距类别折线（不按真实日期定位） |
| `area` | `stacked` | 面积、堆叠面积 |
| `pie` `donut` | — | 单系列、非负值。缺省 `data_labels` 开（饼带类别名，环只有读数）、`legend` 关；读数闭环见下 |
| `scatter` | — | `series[].points = [[x, y], …]`（单点系列允许：用来把一个主角从背景点群里单独标出） |
| `table` | `column_weights`、`column_align` | 原生表格；数据写 `data: {"headers": [...], "rows": [[...], ...]}`，不是 series |

- **颜色映射只有一处**：`model.chart_color_map(plan, chart)` 按下表优先级解析出「类别/系列 → 颜色」，compile、checks、verify、自定义图例全部读它，没有第二份实现。

  | 优先级 | 字段 | 语义 |
  | --- | --- | --- |
  | 1 | `chart.color_map` | 按**名称**显式指派；名称对不上报 `E-CHART`（指派静默失效就是错色根因） |
  | 2 | `chart.palette` | 本图专用色序列，覆盖整稿序列 |
  | 3 | 根 `palette` | 整稿数据色序列，按类别/系列顺序取用 |
  | 4 | `limits.json` → `defaults.palette` | 兜底 |

  主题强调色（accent1）**不参与数据着色**：主题只服务使用者在 PowerPoint 里新建的对象。饼/环分到同色即 `E-CHART`（数据失真），非饼环循环取色给 `W-CHART`。
- **颜色有四种职责，各有唯一来源**（`model.py` 的“图表视觉契约”；compile / checks / verify / `chart_audit.py` 共读）：

  | 职责 | 来源 | 规则 |
  | --- | --- | --- |
  | 语义色（系列、类别、强调项） | `chart_color_map` | 上表优先级 |
  | 承载面 | `chart_surface(page, el)` | 页面底色 + 其下**整块覆盖**该图的色块（同 `resolve_bg`）。深色面板压在浅页上，图按面板取色，不按页面 |
  | 结构色（轴、网格、图例、图题） | `chart_style(surface)` | 取在承载面上对比更高的一套；中灰底两套都 < 4.5 时给 `W-CHART`，改深或改浅 |
  | 标注色（直标） | `chart_label_plan` | 位置与墨色一起定，见下表 |

  | 图 | 直标位置 | 墨色落在 | 墨色 |
  | --- | --- | --- | --- |
  | 簇状柱/条 | `outEnd` | 承载面 | 结构色 |
  | 堆叠/百分比堆叠柱 | `ctr` | 系列色块 | 深墨/白墨取高对比；两者都 < 4.5（如 `#C55A11`）改纯黑 |
  | 面积 | 不写（不允许） | 系列色块 | 同上；多系列非堆叠面积加直标 → `E-CHART`（互相遮挡） |
  | 折线、散点 | `t`（多线时最后一条 `b`） | 承载面 | 多系列沿用系列色（对比 ≥ 4.5），否则退回结构色 |
  | 组合折线 | `t` | 可能压在柱上 | 承载面色衬底 + 系列色/结构色 |
  | 饼 | `outEnd`（带类别名） | 承载面 | 结构色 |
  | 环 | 不写（不允许） | 环带 | 逐扇区深墨/白墨 |

  直标一律写在**系列**上（不依赖 plot 级继承）。`dLblPos` 对类型有限制：donut、area 不允许写，堆叠柱不允许 `outEnd`，XSD 认不出这条，但 PowerPoint 会提示“需要修复”——`chart_audit` 逐处检查。单位取图表级 `unit`，缺省取各系列声明的同一个单位。
  系列色对承载面 < 1.5:1 为 `E-CHART`（几乎不可见），低于 `limits.json` 的 `chart.min_contrast` 为 `W-CHART`。`highlight_category`（**0 起的整数索引**，不是类别名）的图不能开原生图例（图例是系列色，柱是退隐灰）。
- **自定义图例用 `legend` 元素，不要手绘色块**：`{"type":"legend","for_chart":"<本页 chart 的 id>","x":…,"y":…,"w":…,"h":…}`，可选 `direction`（`vertical` 缺省／`horizontal`）、`swatch_in`（缺省 0.16）、`show_value`（仅饼/环附读数）、`font_size`、`color`。色块与标签由被引用图表的**同一份映射**生成，编译成可编辑的组合对象；核验器逐块比对色值与标签，错色即 `V-CHART` 阻断。手绘色块会收到 `W-CHART`（疑似手绘图例）。
- 组合图共用一根数值轴（`combo_line_series` 指定系列改画折线，折线层显式不平滑）。**没有双数值轴**。
- **轴起点**：柱/条/面积在数据全为非负且未写 `value_axis.min` 时，编译器固定从 0 起（自动轴会无声截轴）；`verify.py` 会拦成品里被改掉起点的图（`V-CHART`）。含负值的非堆叠柱开直标时，下界取整齐刻度并多留 25%，避免负柱直标与类别标签相撞；折线与散点保持自动轴。
- **饼**带类别名直标时，绘图区内缩留出标签带；直标单位：中文单位紧贴数字（`12万`），西文单位留一个空格（`12 ms`）。
- 直标：`data_labels` 作用于柱、线、面积、饼、环与组合图的柱层；组合折线的直标需 `combo_line_labels: true`。
- 硬检查：系列名唯一、类别名唯一（它们是颜色映射与 `combo_line_series` 的键）、数值有限（非 NaN / Infinity）；同图单位一致；期间与口径不一致阻断；`axis_labels: false` 须同时 `data_labels: true`；百分比堆叠每列合计 100；截轴（`min > 0`）须在 `value_axis.note` 写明起点，且该起点要落在页面可见的 caption、轴名或图题上（否则 `W-NOTE`）；饼与环既无直标也无图例即阻断。
- **改稿 `chart_data`** 与新建图表同一条数据契约：数值须为有限实数（`model.non_finite_series`，NaN / Infinity / 布尔 / 字符串一律拒绝）、系列名不重复；只改数值不改类别与系列数；原图固定了数值轴 `min`/`max` 而新数值越界会被无声裁掉，直接拒绝；组合图与散点图不接受 `chart_data`，请整稿重建。
- 不支持气泡、雷达、树图、瀑布、桑基、甘特、趋势线、误差线。其余关系用形状与路径手工组合，不与数据绑定。

## 页码

编译器按页序生成，默认右下、12pt、从 1 起。`hide` 列出不显示的页（需在交付说明写理由）。文字与图表不得与页码保留框相交；要在右下放长期来源，整稿改 `corner: top_right`。图片压到角区须用 `page_number_tone` 声明明暗。
