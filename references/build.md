# 机器契约：plan、编译、核验、定点编辑

字段形状以 `scripts/plan.schema.json` 为准，技术阈值以 `scripts/limits.json` 为准。本文只讲 schema 表达不了的语义、坑与命令。

## 坐标与元素

- `x y w h` 为英寸，原点左上，y 向下；`font_size` 为 pt。16:9 = 13.333 × 7.5；4:3 = 10 × 7.5。
- 元素类型：`title` `body` `text` `number` `caption` `footer`（文字），`image` `chart` `shape` `divider` `path` `group`。
- `path` 与 `group` 不写 `x y w h`（由点集或子元素推出）。元素 `id` 页内唯一；成品对象名为 `<页 id>.<元素 id>`，页码对象为 `_page_number`。
- 允许出血；文字与图表伸出画布会被裁（编译期提示）；完全在画布外的元素阻断。

## 文本

- `content` 用 `\n` 分段。同一行内换字重、字号、颜色或字族用 `runs`；`runs` 拼接必须与 `content` 逐字相等。不要把一个句子拆成两个框。
- 文本自带内边距（左右 0.05、上下 0.03 in）。要贴住网格轴线（大标题、数字）写 `pad_in: 0`。
- 默认：`title` 40pt 粗、`number` 72pt 粗、`body` `text` 20pt、`caption` `footer` 13pt；行距 1.12（标题、数字）、1.32（正文）、1.25（注释）。深底（亮度 < 0.18）未声明颜色时用浅墨。
- 字族三分：`fonts.heading`、`fonts.body`（正文、图表、页码）、`fonts.number`（缺省同 heading）。
- `rotation` 对所有对象生效，按 360 取模。

## 叠放、形状、路径、组

- **叠放顺序就是声明顺序**，先写的在下。承载层（底图、蒙版、色块）要有东西压在上面，在自己身上声明 `overlap_allowed: true`；不能写在文字或图表上。实质重叠（面积 ≥ 0.5 in²、短边 ≥ 0.05 in、占较小者 ≥ 25%）且下层未声明即阻断（`E-OVERLAP`）。
- `shape_kind`：`rect` `round_rect` `oval` `triangle` `diamond` `chevron` `arrow_right` `arrow_up`。`corner_in` 仅用于圆角矩形（半径，英寸）。`fill_alpha` 低于 0.15 基本不可见。
- `gradient`：`{from, to, angle, from_alpha?, to_alpha?}`，`angle` 0 为左到右、90 为上到下；末端透明度应接近 0。`stroke`：`{color, width_pt, dash, cap}`。
- 路径顶点为坐标 `[x, y]` 或锚点 `{"el": "n1", "anchor": "right", "at": 0.5}`。两点为直线，三点以上为折线；`closed: true` 加 `color` 为填充。锚点在编译期解析，改形状要重编译；旋转组内的锚点会错位，连线应同组或改用坐标。
- 组内子元素用页面绝对坐标，组内叠放不查重叠。

## 素材与生成图

素材从 `--assets` 解析：先按 `image_id` 精确找，再补 `.png` `.jpg` `.jpeg` `.webp`，不做同名不同扩展名的兜底。图片按 `cover` 裁切，`crop_focus` 为裁窗中心。

每个素材都登记在 `assets`。**生成与否只由 `generation` 决定**：

```json
"assets": {
  "hero":  {"source": "自有素材库 m-01", "rights": "cleared", "reference": "编号 m-01"},
  "plate": {"source": "自建生成", "rights": "pending", "generated": true,
            "generation": {"tool": "图像生成工具", "prompt": "完整提示词",
                           "file": "plate.png", "status": "generated", "sha256": "64 位十六进制"}}
}
```

- `source` 必填。`rights`：`cleared`／`pending`（提示）／`blocked`（阻断）。
- `generated` 只能是对 `generation` 是否存在的声明。`generated: false` 却有记录，或 `generated: true` 却无记录，都阻断（`E-ASSET`）。记录缺 `tool` `prompt` `file` `sha256` 任一即阻断；`status` 须为 `generated`；`file` 须与 `image_id` 解析到同一文件；`sha256` 须与文件实际哈希一致。
- `used_as_evidence: true` 的素材不能是生成图。
- `duotone: {dark, light}` 把图压进色域，原生可编辑；PowerPoint 生效，LibreOffice 只去饱和。
- 核验器比对 PPTX 内嵌字节与素材原件：像素尺寸不同阻断，字节不同（可能重编码）提示。

## 图表

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
| `pie` `donut` | — | 单系列、非负值 |
| `scatter` | — | `series[].points = [[x, y], …]` |
| `table` | `column_weights`、`column_align` | 原生表格 |

- 组合图共用一根数值轴（`combo_line_series` 指定系列改画折线）。**没有双数值轴**。
- 直标：`data_labels` 作用于柱、线、面积、饼、环与组合图的柱层；组合折线的直标需 `combo_line_labels: true`。
- 硬检查：同图单位一致；期间与口径不一致阻断；`axis_labels: false` 须同时 `data_labels: true`；百分比堆叠每列合计 100；截轴（`min > 0`）须在 `value_axis.note` 写明起点，且该起点要落在页面可见的 caption、轴名或图题上（否则 `W-NOTE`）；饼与环既无直标也无图例即阻断。
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
python scripts/verify.py plan.json out.pptx [--render-dir DIR] [--report v.json]
python scripts/revise.py inspect 原稿.pptx [--json i.json]
python scripts/revise.py 原稿.pptx edit.json out.pptx [--assets DIR] [--report r.json] [--force]
```

- 输出只能是 `.pptx`（`E-FORMAT` 阻断）；成品包不合格（内容类型、宏）退出 3。
- `out` 已存在即阻断（`E-EXIST`），`--force` 才替换；先写临时件再原子替换。输出与报告不得指向输入或成品。
- `--assets` 缺省取 plan 同级 `assets/`，否则取 plan 所在目录。`verify.py` 的渲染目录缺省为 `render/<pptx 名>/`。
- 退出码：`compile.py` 0 成功、2 硬错误未写文件、3 运行期失败（`.error.pptx` 不是交付物）；`verify.py` 0 无阻断、1 有阻断或格式不合格；`revise.py` 0 写成功且保留核对通过、1 保留核对未通过（交付路径未写）、2 硬错误、3 运行期失败。

## 诊断码

行格式 `[HARD] E-CHART p2/bar: …`。HARD 用 `E-` 必须改；SOFT 用 `W-` 须逐条判断并写理由；INFO 用 `I-` 只提示看哪一页。`V-` 来自 verify.py。

| 码 | 含义 | 处置 |
| --- | --- | --- |
| `E-ASSET` `W-ASSET` | 素材来源、权利、生成记录不一致 | 补齐记录，或改用真实素材 |
| `E-OVERLAP` | 实质重叠，下层未声明承载 | 多半忘了 `overlap_allowed` |
| `E-CHART` `W-CHART` `W-NOTE` | 图表口径、单位、截轴说明 | 改措辞或拆图，没有放行开关 |
| `E-FORMAT` | 不是合格的无宏 `.pptx` | 输出改为 `.pptx`；原稿含宏则人工另存 |
| `E-PRESERVE` | 定点编辑保留核对不通过 | 看 `preservation` 字段定位越界 |
| `W-FIT` | 装载估算溢出 | 先改文案，再加框，不压字号 |
| `W-CONTRAST` `V-CONTRAST` | 对比不足 | 换墨色或落点，不靠浓遮罩 |
| `W-PPI` | 有效分辨率低 | 换素材或缩小显示 |
| `W-COVER` | 文字落在上层色块或图上 | 挪位置或调层级 |
| `W-RUNS` | 原有混排被统一 | 用 `runs` 写明分区样式 |
| `V-TEXT-RENDER` | 元素的字未落在自己框内渲染 | 查归属，不因别处有同样的字放行 |
| `V-FONT` | 声明字体未参与渲染 | 状态降为 PARTIAL，目标机复核 |

## 定点编辑：改既有稿

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

## 机器边界

- 原生图表只有上表所列类型，没有双数值轴与其余高级图。其余关系用形状与路径手工组合，不与数据绑定。
- 路径只有直线与折线，锚点为编译期坐标，不是 PowerPoint 连接符绑定。
- 页标题为普通文本框，不是标题占位符；字体只写字体名。
- 图片按 `cover` 裁切，不保证主体不被裁到；`duotone` 只换色相。
- 编译器与核验器共用同一套 plan 模型与阈值：共用部分有缺陷，两边会一起“一致”。所以对象核对不能替代看图。
- 没跑过的检查与通过的检查长得一样。新增检查必须反向构造（故意破坏，确认它会响），见 `tests/test_regression.py`。


## Microsoft PowerPoint 真实环境兼容核对清单（Tier 3）

在 LibreOffice 渲染通过（`READY` 或 `PARTIAL`）后，交付到目标 Windows / macOS PowerPoint 放映或审阅时，需按本清单实机复核（未实机打开时必须标为 `NOT VERIFIED`）：

1. **字体与断行**：`fonts.heading` / `fonts.body` / `fonts.number` 在目标机是否已安装；无字体时 PowerPoint 会静默替换并改变行宽，需重点看长标题、表格数字列与多行正文末行。
2. **双色调与蒙版**：`duotone` 在 PowerPoint 中原生生效，而 LibreOffice 仅做去饱和，需实机确认亮暗映射后对比度是否仍满足文字可读。
3. **渐变与透明度**：`gradient` 的 `from_alpha` / `to_alpha` 与形状 `fill_alpha` 在 Windows PowerPoint 下是否平滑过渡，无硬边色带。
4. **裁切与出血**：`image` 的 `cover` 裁切与画布边缘出血对象在放映模式（Slide Show）下是否如期裁齐，无意外白边。
5. **原生图表与直标**：柱/条/折线图表的 `data_labels` 在 PowerPoint 中是否与图例、轴线或相邻柱体碰撞；`highlight_category` 突出色与退隐色层次是否清晰。
