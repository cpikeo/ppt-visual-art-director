# 机器契约：plan、编译、核验、定点编辑

字段形状以 `scripts/plan.schema.json` 为准，技术阈值以 `scripts/limits.json` 为准。本文只讲 schema 表达不了的语义、坑与命令。

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

六种 `role` 的重量分配见 `design.md` 五（节奏）；写不出对应叙事动作，说明页序还没定。

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
- `rotation` 对所有对象生效，按 360 取模。
- 页面 `speaker_notes` 写入备注页（页面 `notes` 是设计备注，不进成品）。元素 `alt` 写入替代文字：图片缺省取 `asset_reason`，图表缺省取图题，装饰元素不写。两者都由 `verify.py` 逐项核对。
- `W-FIT` 的行高按字体取值（`limits.json` 的 `type.line_pitch_em.by_font`）：Noto/思源 CJK 1.448、雅黑 1.32、宋体/黑体 1.15，未知字体 1.25。用别的字体且估不准，在 `by_font` 里补一行。
- 图表与表格内的字号由 `limits.json` 统一给出（表 16、图内标注/轴/图例 14）；要改，改阈值，不逐图改。

## 叠放、形状、路径、组

- **叠放顺序就是声明顺序**，先写的在下。承载层（底图、蒙版、色块、做底的巨型数字或标题）要有东西压在上面，在自己身上声明 `overlap_allowed: true`；`number` 与 `title` 可以做承载层，`body`、`caption`、`footer` 与图表不能自我豁免。实质重叠（面积 ≥ 0.5 in²、短边 ≥ 0.05 in、占较小者 ≥ 25%）且下层未声明即阻断（`E-OVERLAP`）。
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
  系列色对承载面 < 1.5:1 为 `E-CHART`（几乎不可见），低于 `limits.json` 的 `chart.min_contrast` 为 `W-CHART`。`highlight_category` 的图不能开原生图例（图例是系列色，柱是退隐灰）。
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

## 网格与装载估算（算术，不用脚本）

- 列宽 `col_w = (画布宽 − 2×边距 − 沟槽×(列数−1)) ÷ 列数`。默认 16:9、12 列、边距 x 0.7 / y 0.6、沟槽 0.25。第 i 列起点 `x = 边距 + (i−1)×(col_w+沟槽)`；跨 n 列宽 = n×col_w + (n−1)×沟槽。
- 比例切分：`free = 版心宽 − 沟槽×(栏数−1)`，第 i 栏宽 = `free × 比例ᵢ ÷ Σ比例`。
- 参考线：三分 `W/3, 2W/3`、`H/3, 2H/3`；黄金位 `0.382W, 0.618W`、`0.382H, 0.618H`。
- 装载估算：每行字数 ≈ `(w − 2×内边距) ÷ (字号÷72)`，中文一字一 em，西文数字约 0.55 em；高度 ≈ 行数 × 字号 × 1.2 × 行距 ÷ 72。估不准以渲染为准。

## plan 写作顺序

先写 `brief` 与每页的 `role`、`claim`、`focus`（此时无坐标）→ 排页序与节奏 → 逐页定元素与主角 → 最后填几何（先网格算术，再字号档位，再框高）→ 编译，逐条读诊断。宁可改文案，不要改字号。

## 命令与退出码

```bash
python scripts/compile.py plan.json out.pptx [--assets DIR] [--report r.json] [--force]
python scripts/verify.py plan.json out.pptx [--assets DIR] [--render-dir DIR] [--report v.json] [--force]
python scripts/revise.py inspect 原稿.pptx [--json i.json] [--force]
python scripts/revise.py 原稿.pptx edit.json out.pptx [--assets DIR] [--report r.json] [--force]
```

- 输出只能是 `.pptx`（`E-FORMAT` 阻断）；成品包不合格（内容类型、宏）退出 3。
- `out`、`--report`、`--json` 已存在即阻断（`E-EXIST`），`--force` 才替换；先写临时件再原子替换。路径守卫只有一处（`model.guard_paths`）：输出/报告不得与任何输入同文件（符号链接与硬链接按真实文件比），不得经符号链接写出，报告只能是 `.json`，渲染目录不得是输入所在目录。`revise.py inspect` 解析前先做包检查（宏、越界部件名、解压体积）。
- `--assets` 缺省取 plan 同级 `assets/`，否则取 plan 所在目录。`verify.py` 的渲染目录缺省为 `render/<pptx 名>/`。
- 退出码：`compile.py` 0 成功、2 硬错误未写文件、3 运行期失败（`.error.pptx` 不是交付物；它与交付物、plan、报告走同一套 `guard_paths`：符号链接、目录、已存在且无 `--force` 一律不写，写入同样先临时文件再原子替换）；`verify.py` 0 无阻断、1 有阻断或格式不合格；`revise.py` 0 写成功且保留核对通过、1 保留核对未通过（交付路径未写）、2 硬错误、3 运行期失败。

## 诊断码

行格式 `[HARD] E-CHART p2/bar: …`。HARD 用 `E-` 必须改；SOFT 用 `W-` 须逐条判断并写理由；INFO 用 `I-` 只提示看哪一页。`V-` 来自 verify.py。

| 码 | 含义 | 处置 |
| --- | --- | --- |
| `E-ASSET` `W-ASSET` | 素材来源、权利、生成记录、简报不一致；生成图作证据 | 补齐记录，或改用真实素材 |
| `E-OVERLAP` | 实质重叠，下层未声明承载 | 多半忘了 `overlap_allowed` |
| `E-CHART` `W-CHART` `W-NOTE` | 图表口径、单位、截轴说明、颜色映射（`color_map` 指派对不上类别、饼/环同色、手绘图例色块）、系列色沉进承载面、图例与强调柱对不上 | 改措辞、拆图、补颜色映射；错色没有放行开关 |
| `V-LABEL-INK` `V-LABEL-POS` | 成品图表里：直标/轴/图例墨色落在其真实表面上 < 4.5:1（`chart_audit.py` 只读 XML）；或 `dLblPos` 对该图表类型非法；或墨色与决策表不一致 | 重新编译；手改成品会再次被拦 |
| `V-MARKER` | 成品图表里：散点 marker 不可见（symbol=none 或无填充）、小于 5pt、没有关闭连线；折线 marker 色 ≠ 线色。散点 marker 色与调色板对不上记 `V-CHART`；marker 对承载面的对比、灰度区分由 `chart_audit` 经 marker 填充计算（`V-CONTRAST` `V-GRAY`） | 重新编译；手改成品会再次被拦 |
| `W-GRAY` | **任意两个**系列的颜色转灰度后明度差 ΔL* 低于 `chart.min_series_delta_lstar`（成品里同一判断为 `V-GRAY`） | 拉开明度，或直标系列名、换线型与标记 |
| `E-CHART`（读数） | 饼/环没有任何类别来源：饼需直标、原生图例或 `legend` 元素之一；环内直标只有数值，需原生图例或 `legend` 元素 | 开 `legend`／加 `legend` 元素；纯图例无数值直标给 `W-CHART` |
| `V-READOUT` | 成品饼/环里读不到类别（HARD）或数值（SOFT）：只读 XML 里的 `c:legend`、`showCatName`、`showVal` | 与 plan 无关，成品被改过或编译链有误 |
| `E-FORMAT` | 不是合格的无宏 `.pptx` | 输出改为 `.pptx`；原稿含宏则人工另存 |
| `E-PRESERVE` | 定点编辑保留核对不通过 | 看 `preservation` 字段定位越界 |
| `W-FIT` | 装载估算溢出（行数按字宽折算；`%` `×` `@` 按近全角宽度计；高度按“首行可见高度＋后续行距推进”估，按渲染实测校准） | 先改文案，再加框，不压字号 |
| `W-SPACE` | 中文与数字/西文之间手敲了**多于一个**的空格、不断行空格或全角空格 | 删掉，依赖字体原生字距 |
| `I-SPACE` | 中西文之间只有**一个**普通空格：通行的中文排法，LibreOffice 会叠加自动间距显松、PowerPoint 里正常 | 是风格取舍不是缺陷；全稿统一取一种并写进交付说明 |
| `V-BLEED` | 文字落在声明为出血的元素里，被页边裁掉 | 确认被裁部分不含信息；没声明出血的文字出页仍是 `V-CLIP` 阻断 |
| `W-CONTRAST` `V-CONTRAST` | 对比不足；文字压图时是对真实像素的量测（达标像素占比低于 `limits.json` 的 `image.zone.min_ok_share`） | 换墨色或落点、改 `crop_focus`、重做简报；不靠浓遮罩 |
| `W-BUSY` | 图上文字对比够，但下方纹理过强（灰度标准差超限） | 换落字区或重做简报 |
| `W-PPI` | 有效分辨率低 | 换素材或缩小显示 |
| `W-COVER` | 文字落在上层色块或图上 | 挪位置或调层级 |
| `W-RUNS` | 原有混排被统一 | 用 `runs` 写明分区样式 |
| `V-TEXT-RENDER` | 元素的字未落在自己框内渲染 | 查归属，不因别处有同样的字放行 |
| `V-FONT` | 声明字体未参与渲染 | 状态降为 PARTIAL，目标机复核 |
| `W-FONT-FACE` `I-FONT-FACE` | 家族匹配上了，**地区字面**没匹配：声明简体而 PDF 取到日文字面（Noto/思源 CJK 的 `.ttc` 由 LibreOffice 按集合首个字面导出） | 不改状态但必须写进交付说明；直/骨/门/者一类共享码位的字形按别的地区字面渲染，简体字形要人在 PowerPoint 里复核 |
| `V-NOTES` `V-ALT` | 成品里的备注／替代文字与 plan 不一致 | 重新编译；不要手改成品 |

## 定点编辑：改既有稿

**包结构保留（白名单）**：写出前比 (1) 每个部件**生效**的内容类型（`Override` 优先、其次扩展名 `Default`；写法互换不算变），已有部件的类型被换、新增 `ppt/media/` 部件类型不是 image/audio/video，均阻断；(2) 被点名页的非 image/chart/hyperlink/oleObject 关系（版式、备注、批注等）按类型计数须一致；(3) 未点名页的 slide 与 rels、其余部件按规范化指纹逐一一致；(4) 整包 OPC 完整性（`model.package_integrity`：CRC、重名、缺内容类型、Override 指向不存在的部件、关系悬空）。合法变化只有：被点名页与声明改数的图表部件，以及新增 `ppt/media/`、`ppt/embeddings/`。填充与描边的指纹包含主题色（主题名＋明度）与渐变内容摘要，读不出时记 `fill_unreadable` 而不是跳过。

只改已有对象时写 `edit.json`，在原文件上做声明的改动，输出新文件，原稿不动。

```json
{"deck_id": "edit-01", "intent": "标题改成论断，灰块改成整页深底",
 "ops": [
   {"op": "text",  "page": 1, "target": "Title 1", "content": "留存才是增长"},
   {"op": "box",   "page": 1, "target": "Title 1", "y": 4.62, "w": 9.6},
   {"op": "style", "page": 1, "target": "Title 1", "font_size": 60, "color": "#F3F0E8"},
   {"op": "fill",  "page": 1, "target": "Rectangle 1", "color": "#14181B"},
   {"op": "order", "page": 1, "target": "Rectangle 1", "to": "back"},
   {"op": "image", "page": 1, "target": "Picture 1", "image_id": "hero"}]}
```

| `op` | 允许的键（公共键：`page` `target`／`target_regex`／`target_index` `note`） |
| --- | --- |
| `text` | `content` `runs` `space_after_pt`，以及样式键（只改字，不得引入原稿没有的样式） |
| `box` | `x` `y` `w` `h` |
| `style` | `font` `font_ea` `font_size` `bold` `color` `letter_spacing_pt` `align` `line_spacing` `pad_in` `valign` |
| `fill` | `color` `alpha` |
| `line` | `color` `width_pt` `dash` `hide` |
| `order` | `to`（`back`／`front`） |
| `image` | `image_id` |
| `chart_data` | `data`（不改系列数与系列名；组合图不支持） |
| `delete` | — |
| `background` | `color`（按页） |

- 名称重复即阻断，改用 `target_index`。未知键直接拒绝。
- **富文本保真**：段落属性与 run 属性整份沿用，超链接保留；原文含超链接而 run 结构对不上，拒绝写出（`E-OPS`）。含域或软回车（`a:fld` `a:br`）的文本框不做 `text` 改写。
- **保留核对**（见 SKILL.md 第七节）：逐对象 XML（exclusive C14N）、标识与样式容器、超链接地址、页面非形状结构、关系、逐页特性、包内部件指纹。
- **输出**只能是 `.pptx`；原稿含宏即拒绝（`E-FORMAT`）。
- **边界**：不新增页、不复制页、不改页序、不改系列数、不动母版与版式、不加动画与切换。

## 报告字段

- `compile.py`：`blocked`、`output`、`pages`、`sha256`、`issues`。
- `verify.py`：`status`、`sha256`、`render`（PDF、PNG、`font_substituted`、`text_checks`）、`structure`、`layers`（`objects`／`render`／`fonts`／`human`）、`findings`。`render` 非 `PASS` 时 `status` 只能是 `PARTIAL`；`fonts` 为 `SUBSTITUTED` 时不能是 `READY`。
- `revise.py`：`status`、`output`、`sha256`（失败为 null）、`ops`、`preservation`（`parts.changed` 为越界，`parts.moved` 为声明过的改写；`rejected` 为证据路径）、`issues`。

`sha256` 用于确认“这份文件是不是我验的那一份”，不是复现凭据。

## 能力边界

- 判断记录字段（`brief.*`、`audience`、`goal`、`core_claim`、页 `focus` `composition_rationale` `open_questions`、`brief.motif`、`brief.art.*`）由 schema 校验必填与形状，脚本不对其内容做语义判断——它们是写给人与下一轮 AI 的决策记录，不是机器可证明的字段。

| 能力 | 级别 | 说明 |
| --- | --- | --- |
| 画布 16:9、4:3 | 完整 | 其他比例需同时改 `limits.json` 与 schema，未测 |
| 文本、形状、折线路径、组、图片裁切 | 完整 | 对象核对覆盖几何、旋转、逐 run 样式、字节 |
| 原生图表：柱、条、堆叠、百分比堆叠、组合、线、面积、饼、环、散点；原生表格 | 完整 | 无双数值轴、瀑布、桑基、甘特、气泡、雷达、趋势线 |
| 按数据点着色 | 部分 | 仅 `highlight_category`（单系列柱/条，其余退隐灰，不能开原生图例）与饼/环按类别取色；折线、面积、散点不支持，强调单条靠直标或旁注 |
| 颜色映射与图例一致性 | 完整 | 类别/系列→颜色唯一映射（`color_map`＞图级 `palette`＞整稿 `palette`）；`legend` 元素按同一映射生成并逐块核验 |
| 演讲者备注、替代文字、页码、页标题 | 完整 | 页标题是原生标题占位符；阅读顺序仍按声明顺序，未做逐页阅读顺序审查 |
| 既有稿定点编辑 | 完整（限白名单操作） | 不新增页、不改页序、不改系列数、不动母版与版式 |
| 渐变、`duotone` | 有限 | PowerPoint 生效；LibreOffice 只做去饱和或近似 |
| 路径连线 | 有限 | 直线与折线，锚点是编译期坐标，不是连接符绑定 |
| 图片裁切 | 有限 | 按 `cover`，不保证主体不被裁到 |
| 图片生成 | 外部依赖 | 工具无关：只登记工具、完整提示词、文件、SHA-256、权利状态；本包不调用任何模型，生成图不作证据 |
| 渲染核验 | 外部依赖 | LibreOffice + PyMuPDF（`requirements-render.txt`，可选）；缺失给 `PARTIAL`。Noto CJK 的 `.ttc` 在 PDF 里只能按家族匹配，分不出 SC 与 JP 字面 |
| PowerPoint 实机验证 | 外部依赖 | 本包从未做过；结论由人在真实 PowerPoint 中签收，未做即 `NOT VERIFIED` |
| 字体 | 外部依赖 | 只写字体名，不嵌入；目标机缺失会被静默替换。主题字体（拉丁与东亚）取 `fonts`，新建 run 带 `lang`（含中文为 zh-CN） |
| 母版与版式 | 有限 | 默认模板的 4:3 占位符几何按画布宽度横向缩放；只有一套母版，不做自定义版式库 |
| 动画、切换、嵌入字体、新建 SmartArt、宏 | 不支持 | `.pptm` 一律拒绝 |

## 模块与依赖

```
model.py    共用：诊断、plan 载入、度量、颜色映射(chart_color_map)、几何、成品包校验、alt_text（不依赖 python-pptx）
ooxml.py    低层写入：填充、渐变、描边、圆角、duotone、run 属性        ← model
charts.py   原生图表、表格与 legend 图例                               ← model, ooxml
chart_audit.py 成品图表的有效色审计（只读 chart XML，不读 plan 的颜色）     ← model
checks.py   plan 校验与整稿模式信号                                   ← model
compile.py  元素构建、build、CLI                                      ← model, checks, ooxml, charts
verify.py   plan ↔ PPTX ↔ 渲染；environment 事实                       ← model, checks, charts, chart_audit
revise.py    inspect 与定点编辑 + 保留核对                              ← model, ooxml
```

回归测试、作品样本与评测证据**不在分发包内**。新增检查必须故意破坏一次、确认它会响。

`limits.json` 是阈值唯一来源（含图上文字量测 `image.zone`、成品包大小上限 `package`），`plan.schema.json` 是数据契约唯一来源。

## 输入安全

plan、edit.json、素材与 PPTX 都视为不可信输入：素材路径限制在素材根内；解析外部 PPTX 的 XML 一律不展开实体、不联网（防 XXE 与实体膨胀）；成品包拒绝越界部件名与解压后超过 `package.max_unpacked_mb` 的包。脚本不联网、不执行包内内容。

## 机器边界

- 编译器与核验器共用同一套 plan 模型与阈值：共用部分有缺陷，两边会一起“一致”。所以对象核对不能替代看图。堆叠柱标签墨色就是一例：对象一致，但深色块上的深字读不出，只有看渲染才发现。
- 没跑过的检查与通过的检查长得一样。新增检查必须反向构造（故意破坏，确认它会响）。标题里的数字与结论不属于任何检查：必须自己用图表数据回算（评测里出现过标题写“少 1 分钟”而数据斜率是 2.3 的错误）。
- `W-NOTE` 恒为提醒：截轴起点必须由你写进页面可见的 caption、轴名或图题，机器不判断你写没写对。
