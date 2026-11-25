# 元素、文本、素材与图表字段语义

手写或改 plan 元素（`free` 页、图表、素材登记、页码）时读本文；配方、outline 与设计意图字段见 `build.md`。字段速查：`scaffold.py fields <element|chart|asset|…>`。

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
- **原生列表**：`list: "bullet"|"dash"|"number"`，每个 `\n` 分段是一项，写成真正的项目符号/自动编号（`buChar`／`buAutoNum`）与随字号变化的悬挂缩进（编号 1.5 em、符号 1.1 em，不小于 0.32 in；`model.list_indent_for` 是 compile 与估宽共用的唯一实现），符号与第一段文字同色；不要手敲“1.”“•”。估宽按缩进后的版心算；`verify.py` 逐项比对渲染文字（渲染器会在每项前插入符号）。
- 页面 `speaker_notes` 写入备注页（页面 `notes` 是设计备注，不进成品）。元素 `alt` 写入替代文字：图片缺省取 `asset_reason`，图表缺省取图题，装饰元素不写。两者都由 `verify.py` 逐项核对。
- `W-FIT` 行高按字体取 `limits.json` 的 `type.line_pitch_em.by_font`（Noto/思源 CJK 1.448、雅黑 1.32、宋体/黑体 1.15，未知 1.25），估不准就补一行。图表与表格内字号由 `limits.json` 统一给出，要改改阈值。

## 叠放、形状、路径、组

- **叠放顺序就是声明顺序**，先写的在下。承载层（底图、整幅 `hero` 主视觉、蒙版、色块、做底的巨型数字或标题）要有东西压在上面；`hero` 作承载层时，文字必须落在简报声明的安静区，由 `W-CONTRAST` / `W-BUSY` 量测，`subject` / `illustration` / `icon` / `detail` 不能被压，在自己身上声明 `overlap_allowed: true`；`number` 与 `title` 可以做承载层，`body`、`caption`、`footer` 与图表不能自我豁免。实质重叠（面积 ≥ 0.5 in²、短边 ≥ 0.05 in、占较小者 ≥ 25%）且下层未声明即阻断（`E-OVERLAP`）。
- `shape_kind`：`rect` `round_rect` `oval` `triangle` `diamond` `chevron` `arrow_right` `arrow_up`。`corner_in` 仅用于圆角矩形（半径，英寸）。`fill_alpha` 低于 0.15 基本不可见。
- `gradient`：`{from, to, angle, from_alpha?, to_alpha?}`，`angle` 0 为左到右、90 为上到下；末端透明度应接近 0。`stroke`：`{color, width_pt, dash, cap}`。
- 路径顶点为坐标 `[x, y]` 或锚点 `{"el": "n1", "anchor": "right", "at": 0.5}`。两点为直线，三点以上为折线；`closed: true` 加 `color` 为填充。锚点在编译期解析，改形状要重编译；旋转组内的锚点会错位，连线应同组或改用坐标。
- 组内子元素用页面绝对坐标，组内叠放不查重叠。

## 素材与生成图

素材从素材根解析（`compile` `verify` `revise` 共用同一规则；`--assets` 优先，其次 plan 同级 `assets/`，再次 plan 所在目录）：先按 `image_id` 精确找，再补 `.png` `.jpg` `.jpeg` `.webp`，不做同名不同扩展名的兜底。**结果必须留在素材根内**：绝对路径、含 `..`、经符号链接指向根外的文件一律视为找不到（`E-IMG`）。图片按 `cover` 裁切，`crop_focus` 为裁窗中心。

来源、角色（`image_role`）、职责（`info_duty`）三个维度各归各的字段，定义见 `imagery.md` 一。`background` `texture` 是承载层，可声明 `overlap_allowed`；只有 `info_duty: evidence` 带来源约束（不能指向生成图，`E-ASSET`）；不声明或 `none` 给 `W-ASSET`。

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
- `duotone: {dark, light}` 原生可编辑；PowerPoint 生效，LibreOffice 只去饱和。
- 图压着字的量测口径见 `imagery.md` 四。
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
- **颜色四种职责，各有唯一来源**（`model.py`“图表视觉契约”，compile / checks / verify / `chart_audit.py` 共读）：语义色 `chart_color_map`；承载面 `chart_surface`（页面底色＋其下整块覆盖该图的色块，深色面板压在浅页上时按面板取色）；结构色 `chart_style(surface)`（轴、网格、图例、图题，取对比更高的一套，中灰底两套都 < 4.5 给 `W-CHART`）；标注色 `chart_label_plan`（位置与墨色一起定，写在**系列**上）。直标位置：簇状柱/条 `outEnd`、堆叠柱与面积 `ctr`（压在系列色块上，深/白墨取高对比，都 < 4.5 改纯黑）、折线与散点 `t`、饼 `outEnd`（带类别名）、环不写（环带内逐扇区取墨）。`dLblPos` 有类型限制（donut、area 不允许写，堆叠柱不允许 `outEnd`；XSD 认不出，PowerPoint 会提示“需要修复”），`chart_audit` 逐处检查；多系列非堆叠面积加直标 → `E-CHART`。单位取图表级 `unit`。系列色对承载面 < 1.5:1 为 `E-CHART`，低于 `chart.min_contrast` 为 `W-CHART`。`highlight_category`（**0 起整数索引**，不是类别名）的图不能开原生图例。
- **自定义图例用 `legend` 元素，不要手绘色块**：`{"type":"legend","for_chart":"<本页 chart 的 id>","x":…,"y":…,"w":…,"h":…}`，可选 `direction`（`vertical` 缺省／`horizontal`）、`swatch_in`（缺省 0.16）、`show_value`（仅饼/环附读数）、`font_size`、`color`。色块与标签由被引用图表的**同一份映射**生成，编译成可编辑的组合对象；核验器逐块比对色值与标签，错色即 `V-CHART` 阻断。手绘色块会收到 `W-CHART`（疑似手绘图例）。
- 组合图共用一根数值轴（`combo_line_series` 指定系列改画折线，折线层显式不平滑）。**没有双数值轴**。
- **轴起点**：柱/条/面积在数据全为非负且未写 `value_axis.min` 时，编译器固定从 0 起（自动轴会无声截轴）；`verify.py` 会拦成品里被改掉起点的图（`V-CHART`）。含负值的非堆叠柱开直标时，下界取整齐刻度并多留 25%，避免负柱直标与类别标签相撞；折线与散点保持自动轴。
- `data_labels` 作用于柱、线、面积、饼、环与组合图的柱层；组合折线直标需 `combo_line_labels: true`。
- 硬检查：系列名唯一、类别名唯一（它们是颜色映射与 `combo_line_series` 的键）、数值有限（非 NaN / Infinity）；同图单位一致；期间与口径不一致阻断；`axis_labels: false` 须同时 `data_labels: true`；百分比堆叠每列合计 100；截轴（`min > 0`）须在 `value_axis.note` 写明起点，且该起点要落在页面可见的 caption、轴名或图题上（否则 `W-NOTE`）；饼与环既无直标也无图例即阻断。
- **改稿 `chart_data`** 与新建图表同一条数据契约：数值须为有限实数（`model.non_finite_series`，NaN / Infinity / 布尔 / 字符串一律拒绝）、系列名不重复；只改数值不改类别与系列数；原图固定了数值轴 `min`/`max` 而新数值越界会被无声裁掉，直接拒绝；组合图与散点图不接受 `chart_data`，请整稿重建。
- 不支持气泡、雷达、树图、瀑布、桑基、甘特、趋势线、误差线。其余关系用形状与路径手工组合，不与数据绑定。

## 页码

编译器按页序生成，默认右下、12pt、从 1 起。`hide` 列出不显示的页（需在交付说明写理由）。文字与图表不得与页码保留框相交；要在右下放长期来源，整稿改 `corner: top_right`。图片压到角区须用 `page_number_tone` 声明明暗。
