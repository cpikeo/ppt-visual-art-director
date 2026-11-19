# 机器契约：plan、编译、核验

设计与机器之间的边界。**字段形状以 `scripts/plan.schema.json` 为准**（必填项、枚举、数值范围都在那里，本文不重复罗列）；技术数值以 `scripts/limits.json` 为准。本文只解释 schema 表达不了的语义、容易踩的坑与命令。

## 坐标与画布

- `x / y / w / h` 单位是**英寸**，`font_size` 是 **pt**；原点在页面**左上角**，`y` 向下为正。
- 16:9 画布 13.333 × 7.5；4:3 画布 10 × 7.5（`aspect_ratio`）。
- `path` 与 `group` 不写 `x/y/w/h`：前者由点集推出包围盒，后者由子元素并集推出。写了也被忽略。
- 元素 `id` 页内唯一（含 group 子元素）；成品对象名是 `<页 id>.<元素 id>`，页码对象固定为 `_page_number`。
- 出血允许：元素可以超出画布；文字与图表伸出会被裁掉（编译期提示），完全在画布外的元素直接阻断。

## 文本

- `content` 用 `\n` 分段；**同一行内**要换字重、字号、颜色或字体，用 `runs`。
- `runs` 也可以含 `\n`——它同时承担分段：拼接（连 `\n` 一起）必须与 `content` 逐字相等。两行标题第二行换色，就在一个文本框里写 `runs`，不要拆成两个框。
- 不要把一个句子拆成两个文本框来混排——基线会漂移，核验也会按两个对象算几何。
- 文本自带内边距（左右 0.05 in、上下 0.03 in），装载估算与真实排版都按它算；框高要留出这段余量。要让文字**贴住**网格轴线（大标题、数字、页边注释），写 `pad_in: 0`，否则光学边缘会偏 0.05 in。
- 装载估算按**逐段最大字号**与逐 run 折算的宽度算，也把 `space_after_pt` 计入：标题里放一个 96pt 数字加 24pt 单位，按元素级字号一刀切会少算好几行。估不准就以真实渲染为准。
- 没写样式时取默认：`title` 40pt 粗／`number` 72pt 粗／`body` `text` 20pt／`caption` `footer` 13pt；行距 1.12（标题、数字）、1.32（正文）、1.25（注释）。深底页（亮度 < 0.18）未声明 `color` 时自动用浅墨。默认值只保证"不会出错"，不保证好看。
- 字族三分：`fonts.heading`（标题）、`fonts.body`（正文、图表、页码）、`fonts.number`（关键数字，缺省同 heading）。数字单独换等宽或窄体，是编辑设计里最省力的一处区分。
- `rotation` 对文字、形状、图片、图表、路径、组合一律生效（编译期写在对象自己的 `xfrm rot` 上），核验按 0.2° 回读。角度写多少都行，落盘时按 360 取模：`-90` 与 `270` 是同一个竖排。

## 叠放与重叠

**叠放顺序就是声明顺序**：先写的在下。没有第二个开关。

背景图、蒙版、色块、闭合路径这类承载层，要有东西压在它上面时，在**它自己**上声明 `overlap_allowed: true`。压在上面的元素不能自我豁免——被压的那层才知道自己是不是承载层。

`overlap_allowed` 只能出现在 `shape` / `divider` / `path` / 背景或纹理图上；写在文字或图表上会阻断。实质性重叠（面积 ≥ 0.5 in²、短边 ≥ 0.05 in、占较小者 ≥ 25%）且下层未声明时阻断。

## 形状、渐变、描边、圆角

- `shape_kind`：`rect` / `round_rect` / `oval` / `triangle` / `diamond` / `chevron` / `arrow_right` / `arrow_up`。
- `corner_in` 只用于 `round_rect`：圆角**半径**（英寸），按短边换算成 `adj`，超过短边一半会被裁成胶囊。不给就用渲染器默认（约短边 16.7%，常显过圆）。
- `color` 是填充，`fill_alpha` 是不透明度（低于 0.15 基本读不出来）；只有描边的框不给 `color`，只给 `stroke`。
- `gradient` 是两停线性渐变 `{from, to, angle, from_alpha?, to_alpha?}`；`angle` 0 为从左到右、90 为从上到下。用于压图蒙版、光带、大面积过渡。**末端透明度要接近 0**，否则边界会显形。
- `stroke` 是 `{color, width_pt, dash, cap}`；`dash` 取 `solid` / `dash` / `dot` / `dash_dot` / `long_dash`。成品不带主题投影：形状的填充与描边全部来自 plan。

## 路径与锚点

顶点二选一：坐标 `[x, y]`，或别的元素的锚点 `{"el":"n1","anchor":"right","at":0.5}`。两个点 → 直线连接符；三点以上 → 自由折线；`closed: true` 加 `color` 就是填充多边形。`arrow` 取 `none` / `end` / `both`，闭合路径不画箭头。

**锚点在编译期解析成坐标**：这是确定性换来的代价。改完形状要重新编译，图不会自己跟着动。

**旋转组里的锚点会错位。** 组旋转作用于成品对象，而锚点按未旋转的几何解析——把线连到旋转组内的子元素，线会落在错的位置。要么把连线也放进同一个组，要么改用显式坐标。

核验会读回每个顶点的位置与闭合开关：顶点位移超过 0.02 in 即阻断。

## 组合

子元素用**页面绝对坐标**，包围盒由编译器推出；整组可在 PowerPoint 里再拆开单独编辑。组内成员之间的叠放是设计许可的，不查重叠；组与其他元素之间照常检查。文字压在组内色块上时，色块仍要声明 `overlap_allowed`。

## 图片与素材

素材从 `--assets` 目录解析：先按 `image_id` 精确找，再补常见扩展名（`.png` / `.jpg` / `.jpeg` / `.webp`）。**不做"同名不同扩展名"的兜底**——那会在 `hero.png` 与 `hero-2.jpg` 并存时悄悄取到另一张图，而核验器用同一个解析函数，两边一起错。图片按 `cover` 等比裁切到给定框，`crop_focus` 是裁窗中心，`alpha` 控制不透明度。

```json
"assets": {
  "hero":  {"source": "自有素材库 m-01", "rights": "cleared", "reference": "编号 m-01", "generated": false},
  "plate": {"source": "自建生成", "rights": "pending", "generated": true,
            "generation": {"tool": "图像生成工具", "prompt": "实际提交的完整提示词",
                           "file": "plate.png", "status": "generated", "sha256": "64 位十六进制"}}
}
```

- `source` 必填；`rights` 取 `cleared` / `pending` / `blocked`（版权与肖像、隐私合并判断）。`blocked` 阻断，`pending` 给提示——交付前确认，或写进交付说明。声明 `cleared` 却没写 `reference` 会提示。
- `used_as_evidence: true` 的素材必须是真实素材；生成图作证据直接阻断。
- 有 `generation` 就说明是生成图：`status` 必须是 `generated`，`file` 必须与页面 `image_id` 解析到**同一个物理文件**，`sha256` 必须等于该文件真实哈希；核验器还会拿它与 PPTX 内嵌字节比对。
- `duotone: {dark, light}` 把图压进整稿色域：暗部→`dark`、亮部→`light`。原生可编辑，随时可摘；PowerPoint 生效，LibreOffice 只做去饱和。

## 图表

```json
{"kind": "bar", "horizontal": true, "data_labels": true, "axis_labels": false,
 "title": "自然搜索与老客推荐撑住了留存",
 "unit": "%", "period": "2025 全年", "basis": "次月与第三月两次访问",
 "highlight_category": 0,
 "data": {"categories": ["自然搜索", "老客推荐", "信息流", "应用商店"],
          "series": [{"name": "次月留存", "values": [52, 47, 31, 24], "unit": "%"}]}}
```

| `kind` | 修饰 | 得到什么 |
| --- | --- | --- |
| `bar` | `horizontal`、`stacked`（`true` / `"percent"`）、`combo_line_series` | 竖柱、横条、堆叠、百分比堆叠、柱＋线组合 |
| `line` | — | 等距类别折线（不按真实日期定位） |
| `area` | `stacked` | 面积图、堆叠面积图 |
| `pie` `donut` | — | 单系列、非负值 |
| `scatter` | — | 纯标记散点，数据用 `series[].points=[[x,y],…]` |
| `table` | `column_weights`、`column_align` | 原生表格，细横线分组 |

**组合图**：`kind: "bar"` + `combo_line_series: ["毛利"]`，把指定系列改画折线，与柱共用同一根数值轴。只有共轴才算组合图；**没有双数值轴**。

**直标分三层，互不越界**：

- `data_labels` 作用于柱/线/面积/饼/环，以及组合图的**柱层**。
- 折线层的直标要单独写 `combo_line_labels: true`。折线上的点通常落在柱体里，统一墨色在上面读不出——默认不给。
- 饼图直标放在扇区外并带类别名，因此默认不出图例；环图直标留在环带内、只留数值，按带色自动选深墨或白墨，类别交给图例。`legend` 可显式开关这两者。

**编译期硬检查**（口径是最常见的失真来源）：

- 同一张图里**单位必须一致**；期间与口径同理，不一致就阻断——没有"确认一下就能过"的开关，改措辞或拆图。
- `axis_labels: false` 收起整根数值轴，必须同时 `data_labels: true`；直标会自动带单位（`52%`、`14.2 亿元`）。
- 百分比堆叠每列必须合计 100：开了直标就直接阻断，没开也会提示——否则图形归一化而直标仍是原值。
- 截轴必须在 `value_axis.note` 写明起点（`min > 0` 且无 note 即阻断）。
- `highlight_category` 只用于单系列、非堆叠柱／条图：该类别用调色板首色，其余退到按底色兜底的中性色（只改明度不改色相）。
- 饼／环既无直标也无图例即阻断；散点没有轴名时会给提示。
- `show_title: false` 连图题一起不画：结论已由页面标题承担时，图只留读数。
- 组合图折线层没开 `combo_line_labels` 会给提示：要读数就开，要形状就认下这条提示。
- 环图的直标是逐点节点：墨色与数字格式都写在每个点上，单位不会在渲染里丢掉（核对时看得到 `47%` 而不是 `47`）。
- `value_axis.note` 是写给**编译器与读者之外**的说明字段：它不会画进图里。截轴起点必须落到页面的 caption、轴名或图题上，否则 `W-NOTE` 会提示这一点——读者看不到的限定等于没有。
- 表格不画图题；数字列右对齐，`column_weights` 数量必须等于列数。

## 页码

整稿一个契约，编译器按页序生成，不在页面里手写。默认右下、12pt、从 1 起、纯数字。

- `hide` 列出不显示页码的页；确需隐藏要在交付说明里写理由。
- 文字与图表不得与页码保留框相交（编译期阻断，旋转对象按转后的外接框算）——右下要长期放来源或图注时，整稿改 `corner: top_right`，不要逐页切换。
- 色块压到角区是允许的（页码总在最上层）：编译器会按页码**实际落点**在深墨/浅墨两档里挑能读的那个，所以"底色其实由一张铺满的色块给出"这类页不会写出深底深字。
- 图片压到角区无法靠一个色值判读，要自己用 `page_number_tone` 声明角区明暗；不声明就按页面底色算，并在编译期给 `I-PAGE` 提示。

## 网格与装载估算

```bash
python scripts/grid.py --span 1:6 --span 8:5     # 默认 16:9、12 列、边距 0.7/0.6、沟槽 0.25
python scripts/grid.py --split 5,7               # 按比例切版心（非对称构图的骨架）
python scripts/grid.py --guides thirds           # 三分线；--guides golden 取黄金位
python scripts/grid.py --json                    # 含页码保留框，供脚本消费
```

先定 2–4 条主对齐轴（左边缘、内容区起点、右栏起点、底部来源线），元素贴轴。跨列要对应内容关系：论断＋证据常见 5＋空档＋6，真正等重的对照才等分。空列是资产——空档应连续并承担分组、停顿或回秤。

**比例先于列。** 等分只是三种起点之一：等分给秩序，三／四分给稳定，黄金位给紧张。`--split 3,5,4` 直接给出按比例切分的栏宽与起点（含沟槽），`--guides` 给出三分或黄金位坐标——先知道线在哪，再决定东西放不放上去。

**装载估算**（编译期用同一模型给提示）：每行满宽字数 ≈ `(w − 2×内边距) ÷ (字号 ÷ 72)`，中文一字一 em、西文数字约 0.55 em；需要的高度 ≈ 行数 × 字号 × 1.2 × 行距 ÷ 72。`pad_in: 0` 会放宽每行字数，估不准就以真实渲染为准。

## 最小骨架

能通过编译的最小单页，结构对，内容自己写。

```json
{
  "deck_id": "demo", "title": "示例",
  "audience": "谁看、已知什么、在哪看",
  "goal": "离场后要理解或决定什么",
  "core_claim": "最重要的一条主张",
  "brief": {
    "insight": "观众以为问题在获客，数据显示在留存",
    "concept": "用一条不断下移的基线，把流失画成位移而不是缺口",
    "veto": ["不用漏斗图——它暗示一个并不存在的线性过程"],
    "art": {"light": "侧逆光，中性偏冷", "lens": "平视中景", "material": "哑光纸",
            "gamut": "炭黑到纸白，一处赭石", "grain": "细颗粒"},
    "motif": ["左下角固定短基线"]
  },
  "pages": [{
    "id": "p1", "role": "open", "claim": "",
    "focus": "先看标题，再看副题",
    "focus_element": "t",
    "composition_rationale": "标题压在下三分之一，上方留连续暗场，重心靠底部横带稳定",
    "page_number_tone": "dark",
    "elements": [
      {"id": "bg", "type": "shape", "color": "#14181B", "x": 0, "y": 0, "w": 13.333, "h": 7.5, "overlap_allowed": true},
      {"id": "t", "type": "title", "content": "留存才是增长",
       "x": 0.9, "y": 4.62, "w": 9.6, "h": 1.28, "font_size": 64, "color": "#F3F0E8", "pad_in": 0},
      {"id": "sub", "type": "body", "content": "2026 年度复盘",
       "x": 0.9, "y": 4.0, "w": 6.0, "h": 0.44, "font_size": 20, "color": "#C9C3B6"},
      {"id": "rule", "type": "path", "points": [[0.92, 6.3], [2.4, 6.3]],
       "stroke": {"color": "#B8834A", "width_pt": 1.25}, "arrow": "end"}
    ]
  }]
}
```

## 怎么写一份 plan

顺序是：先写 `brief`（洞察、概念、否决、艺术指导），再写每页的 `role`、`claim`、`focus`（这时还没有坐标），再排页序与节奏，然后逐页决定"用几个元素、谁承担主角"，最后才填几何。

- 几何先跑 `grid.py`，取列坐标而不是估数字；同一页的元素贴 2–4 条轴。
- 文字先按内容需要定字号档位，再用装载估算反推框高。宁可改文案，不要改字号。
- 一个元素一个职责：把"论断＋证据＋来源"塞进一个框，最后一定会出问题。
- 图表先写 `unit` `period` `basis` 和 `title`（说结论），再填数据。
- 写完先编译：硬错误会指出具体位置；提示逐条判断，不理解的不要忽略。

## 命令与退出码

```bash
python scripts/compile.py plan.json out.pptx [--assets DIR] [--report compile.json] [--force]
python scripts/revise.py 原稿.pptx edit.json out.pptx [--assets DIR] [--report revise.json] [--force]
python scripts/verify.py plan.json out.pptx [--render-dir render] [--report verify.json]
python scripts/inspect_pptx.py 原稿.pptx [--json inspect.json]
python scripts/grid.py [--aspect 4:3] [--cols 12] [--span 1:6] [--split 5,7] [--guides thirds] [--corner top_right] [--json]
```

- **不覆盖已有文件**：`out.pptx` 已存在时编译直接阻断（`E-EXIST`），确认要替换再加 `--force`。先写临时文件再原子替换，失败不留半成品；原稿不会被就地改写。
- **源文件保护**：输出、报告都不得指向 plan 自己，报告也不得指向成品；输出必须是 `.pptx`。违反即阻断，不写文件。
- `--assets` 缺省取 plan 同级的 `assets/`，没有则取 plan 所在目录。
- `verify.py` 的 `--render-dir` 缺省为 pptx 同级的 `render/<pptx 名>/`，写 PDF、逐页 PNG 与 `contact_sheet.png`。
- 退出码：`compile.py` 0 成功（可能有提示）、2 硬错误未写文件、3 运行期失败（另写 `.error.pptx`，不是交付物）；`verify.py` 0 无阻断、1 有阻断、2 输入不可读；`revise.py` 0 写成功且保留核对通过、1 保留核对未通过（交付路径未写）、2 硬错误、3 运行期失败。
- **核验状态只有三档，别把它们读成质量**：`BLOCK` 有硬问题；`PARTIAL` 对象核对通过但**渲染没跑**（缺 LibreOffice / PyMuPDF，或转换失败）——此时视觉判断没有任何机器依据；`READY` 对象与渲染都一致。三档都不表示"设计通过"。报告里 `layers` 分别记 `objects` / `render` / `human`，`human` 永远是 `NOT REVIEWED`，只有看图的人能改它。

## 定点编辑：改既有稿

重建会丢母版、动画、切换、SmartArt、嵌入对象、备注。只改已有对象时，写 `edit.json` 走 `revise.py`：在原文件上做声明的改动，其余原样，输出新文件，原稿不动。

```json
{
  "deck_id": "edit-01",
  "intent": "第 1 页标题改成论断、压到下半部、把灰块改成整页深底并换图",
  "ops": [
    {"op": "text",  "page": 1, "target": "Title 1", "content": "留存才是增长"},
    {"op": "box",   "page": 1, "target": "Title 1", "y": 4.62, "w": 9.6},
    {"op": "style", "page": 1, "target": "Title 1", "font_size": 60, "color": "#F3F0E8", "pad_in": 0},
    {"op": "box",   "page": 1, "target": "Rectangle 1", "x": 0, "y": 0, "w": 13.333, "h": 7.5},
    {"op": "fill",  "page": 1, "target": "Rectangle 1", "color": "#14181B"},
    {"op": "order", "page": 1, "target": "Rectangle 1", "to": "back"},
    {"op": "image", "page": 1, "target": "Picture 1", "image_id": "hero"}
  ]
}
```

| `op` | 参数 | 说明 |
| --- | --- | --- |
| `text` | `content`（必需）、`runs`、可选的 `align` `line_spacing` `space_after_pt` | 只换字时**不动样式**：新段落 i 的样式基线取原稿第 i 段；该段新 run 数与原 run 数一致时按位置继承（深色主句＋灰色括注不会被压平）。原稿没声明的字级继续继承，不会被写成 18pt |
| `box` | `x` `y` `w` `h` 任意几个 | 只改给了的边 |
| `style` | `font` `font_ea` `font_size` `bold` `color` `align` `line_spacing` `letter_spacing_pt` `valign` `pad_in` | **只改给了的属性**，其余保持原样。声明 `font` 会连东亚字族一起写 |
| `fill` | `color`、`alpha` | |
| `line` | `color` `width_pt` `dash` `hide` | `hide: true` 去描边 |
| `order` | `to`: `back` / `front` | 层级也是版面 |
| `image` | `image_id` | 换图，保留框位与裁切 |
| `chart_data` | `data.categories`、`data.series` | 系列数与系列名必须与原图一致；组合图（多图层）拒绝，会打乱两层格式 |
| `delete` | — | 点名了才允许消失 |
| `background` | `color` | 整页底色 |

- `page` 是 1 起页序；`target` 是对象名（`p2.t` 或 `Title 1`），也可 `target_regex`；重名时用 `target_index`（0 起，z 序）。定位不到或命中多个即阻断，不猜。**搜索不到组内对象**——先取消组合。
- **保留核对四层，缺一层就等于没查**：
  1. **结构**：页数、母版数、版式数、主题字体、宏。
  2. **逐页特性**：动画、切换、SmartArt、嵌入对象、音视频、备注是否还在。
  3. **对象**：未被点名的对象，类型、几何、旋转、**全文**、填色描边、裁切、图片字节、图表载荷、表格逐格逐项一致；**出现原稿没有的对象、或对象消失而没有 `delete` 点名，都算漂移**。
  4. **包内部件**：除被点名页的 `slideN.xml` 与它的 rels、`[Content_Types].xml`、新增的 `media/`、`embeddings/`，以及**声明过要改的那张图自己的 `chartN.xml` 与它的内嵌工作簿**，其余部件必须逐字节一致。母版主题、备注页被意外重写时，前三层全都对得上。放行是按部件名的：改了 A 图，B 图的部件被重写照样报。
- **只改文字的闸**：`text` 操作若没有声明任何样式字段（`font` `font_ea` `font_size` `bold` `color` `letter_spacing_pt` `align` `line_spacing` `pad_in` `valign`），改动后每一段用到的样式必须是它**原来就有**的。允许"少用"（那句灰色括注被删掉了），不允许"新出现"（14pt 副题被压成 32pt，或继承的 44pt 被写死成 18pt）——这类漂移页数、母版、几何全都对得上，只比文字字段是发现不了的。签名含段落级的对齐与行距：顺手改掉对齐是同一类漂移。同一对象上另有 `style` 操作、并显式声明了某个字段的，那个字段不参与判定（否则"先换字、再改字号"会被自己的闸拦住）。原段是混排、新内容 run 数又不同，会记 `W-RUNS` 说明"已统一为该段首个 run 的样式"：这是披露，不是通过。
- **核对不通过就不落交付路径**：先写临时件，核对通过才原子替换到 `out`；不通过时 `out` 一个字节都不动，证据另存 `<out 名>.preserve-failed.pptx`（不是交付物）。
- **改完再查版面**：`W-COVER` 报"文字落在上层色块上"（按字的落点算，不按空框），`W-BOUND` 报出画布。这两条只防不可读，不判断好看。
- **边界**：不新增页、不复制页、不改页序；不改系列数（会破坏配色与格式）；不动母版与版式本身；不能加动画或切换。要这些就整稿重建，并先在交付说明里点名会失去什么。

## 读诊断

诊断行形如 `[HARD] E-CHART p2/bar: …`。级别决定处置，前缀也按级别走：**HARD 用 `E-`、SOFT 用 `W-`、INFO 用 `I-`**。

- **`[HARD]`**：必须改，文件没写出来（或核验判为 BLOCK）。按消息指到的页/元素改，不要绕过。
- **`[SOFT]`**：可能有问题，逐条判断。判为可接受就在交付说明里写接受的理由；不要当没看见。
- **`[INFO]`**：说明性事实（如"文字下方是图片，对比需看渲染"），只影响你要看哪一页。

几条经验：`W-FIT` 先改文案再加框；`W-CONTRAST` 先换墨色或换落点，不靠加浓遮罩；`W-PPI` 换素材或缩小显示尺寸；`E-OVERLAP` 九成是忘了给承载层声明 `overlap_allowed`；`W-MONOTONY` 先确认那三页是不是同轴比较；`W-EMPTY` 确认那一页是有意的极简还是漏了内容。

## 机器边界

- 原生图表只有柱（横／竖／堆叠／百分比）、线、面积（含堆叠）、饼、环、散点与表格，以及柱＋线组合。没有双数值轴、气泡、雷达、树图、瀑布、桑基、自动甘特、趋势线与误差线。其余关系用文字、形状与路径手工组合：这些对象可单独编辑，但点的位置不与数据绑定，改数据要重新排版。
- `path` 的锚点是编译期坐标，不是 PowerPoint 的连接符绑定；组旋转后锚点不复位。路径只有折线与直线，没有贝塞尔曲线。
- 不生成演讲者备注、动画、页面切换、SmartArt、主题母版。
- 页标题是普通文本框，不是标题占位符：PowerPoint 大纲视图与读屏软件识别不到。无障碍有要求时在交付说明里提示。
- 字体只写字体名，不嵌入字体文件。
- 图片按 `cover` 裁切，`crop_focus` 只给裁窗中心；不保证主体不被裁到——那是选图与框的事。`duotone` 只换色相，不改明暗结构。
- 核验器与编译器共用同一套 plan 模型与阈值。共用部分若有缺陷，两边会一起"一致"而发现不了；所以对象核对不能替代看图。图表系列色、数值与文字均从成品 XML 实际读取，不从 plan 回算。
- **没跑过的检查与通过的检查长得一样**。验证要先反向构造：把某条规则故意破坏，确认它真的会响。饼环直标核对曾因一个提前 `return` 长期没执行，正向测试全绿；文字旋转、图片旋转、路径顶点也都曾"声明了但没生效"。改完必须重跑全套，不要只跑刚改的那条。
- 这条边界还会以更隐蔽的方式出现：**两个数放在不同单位里比较，检查就永远不响**。渲染行的坐标曾是 PDF 点（1/72 in）而 plan 是英寸，于是"孤字""行首标点"两条长期沉默；孤字判定又曾依赖 PDF 的 block 结构，而 LibreOffice 把每行写成独立 block，同样永不触发。两条都是全绿用例，靠反向构造才暴露。凡是"把 plan 和渲染结果对起来"的判断，先确认两边单位与分组口径一致，再相信它的绿灯。

## 报告

`--report` 写出的 JSON 是证据材料，不是成绩单。三个脚本字段不同，别混着读：

| 脚本 | 顶层字段 |
| --- | --- |
| `compile.py` | `blocked`、`output`、`pages`、`sha256`、`issues` |
| `verify.py` | `status`、`pptx`、`sha256`、`bytes`、`slides`、`expected_slides`、`render`（PDF、PNG、**实际用到的字族**、`text_checks`）、`structure`（slide／原生图表／内嵌工作簿／media 数）、`layers`（`objects`／`render`／`human`）、`asset_root`、`images`、`findings` |
| `revise.py` | `status`、`source`、`output`、`sha256`、`ops`（每条是否 `style_free`、声明过哪些样式）、`preservation`（`deleted`／`touched`／`parts`＝`parts`·`changed`·`lost`·`added`·`moved`／`rejected`）、`issues` |

两个细节别读错：`verify.py` 的 `layers` 是本次判断的依据分层，`render` 为 `"SKIP"` 时 `status` 只能是 `PARTIAL`；`revise.py` 的保留核对不通过时 `output` 字段仍指向你请求的路径，但那个文件**没有被写入**，`sha256` 为 `null`，证据在 `preservation.rejected`。`preservation.parts.moved` 是"声明过、因此允许被重写"的部件——改数会重写那张图自己的 `chartN.xml` 与内嵌工作簿，它们出现在这里是对的；`changed` 才是越界。

`sha256` 用来区分"这份文件是不是我验的那一份"，不是复现凭据：同一份 plan 编译出的对象与几何一致，但图表内嵌的工作簿带时间戳，字节不会完全相同。
