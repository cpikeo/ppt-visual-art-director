# 工程契约：plan、编译、核验

设计与机器之间的边界。字段形状以 `scripts/plan.schema.json` 为准，技术数值以 `scripts/limits.json` 为准，本文件解释语义与做法。

## 坐标与画布

- `x / y / w / h` 单位是**英寸**，`font_size` 是 **pt**；原点在页面**左上角**，`y` 向下为正。
- 16:9 画布 13.333 × 7.5；4:3 画布 10 × 7.5（`aspect_ratio`）。
- 元素 `id` 页内唯一；成品对象名是 `<页 id>.<元素 id>`，页码对象固定为 `_page_number`。
- 出血允许：元素可以超出画布，但文字超出会被裁切（编译期给提示），完全在画布外的元素直接阻断。

## 字段

**整稿**

| 字段 | 用途 |
| --- | --- |
| `deck_id` `title` | 标识与文件标题 |
| `audience` `goal` `core_claim` | 受众与场景、离场目标、最重要的一条主张 |
| `visual_concept` | 由哪条洞察推出哪个概念、否决了什么 |
| `rhythm` | 整稿乐谱，一页一句 |
| `palette` | 图表消费色，按系列／类别顺序取用；缺省用 `limits.json` |
| `fonts` | `heading` / `body` 两个字族角色 |
| `aspect_ratio` | `16:9`（默认）或 `4:3` |
| `page_numbers` | `corner`、`start`、`hide`（隐藏页码的页 id 列表） |
| `assets` | 素材登记，见「素材」 |
| `pages` | 页面 |

**页面**

| 字段 | 用途 |
| --- | --- |
| `id` `role` | 标识；叙述职责 `open/build/breathe/turn/hold/close` |
| `claim` | 本页主张；封面与过渡可留空 |
| `focus` | 读者这一页先看到什么、按什么顺序读 |
| `focus_element` | 承担主角的元素 id；有意并列时不写 |
| `composition_rationale` | 重心怎么成立、为什么这样断、放弃了什么 |
| `evidence` `open_questions` | 支撑主张的事实与来源；还没定的问题 |
| `background` | 页面底色，缺省白 |
| `page_number_tone` | 页码角区的实际明暗场，仅当与底色不同时声明 |
| `notes` | 设计备注（不写入演讲者备注） |
| `elements` | 元素 |

**元素**：`id` `type` `x` `y` `w` `h` 必填，按类型加：

| type | 必填 | 可选 |
| --- | --- | --- |
| `title` `body` `text` `number` `caption` `footer` | `content` | `font_size` `font` `bold` `color` `align` `valign` `line_spacing` `letter_spacing_pt` `space_after_pt` `runs` |
| `image` | `image_id` `asset_reason` | `image_role` `crop_focus` |
| `chart` | `chart` | — |
| `shape` `divider` | `color` | `shape_kind`（`rect` / `round_rect` / `oval`）`fill_alpha` `overlap_allowed` |

`runs` 用于同一行内的混排（数字配小一号单位、局部换色换字重）：拼接必须与 `content` 逐字相等，且此时 `content` 不能含换行。`divider` 就是一条极薄的色块（`h` 取 0.02–0.05），不是线对象。

没写的样式按类型取默认：`title` 40pt 粗／`number` 72pt 粗／`body` `text` 20pt／`caption` `footer` 13pt，行距 1.12（标题、数字）与 1.32（正文）。深底页（亮度低于 0.18）未声明 `color` 时自动用浅墨。默认值只保证"不会出错"，不保证好看。

## 最小骨架

能通过编译的最小单页。结构对，内容你自己写。

```json
{
  "deck_id": "demo",
  "title": "示例",
  "audience": "谁看、已知什么、在哪看",
  "goal": "离场后要理解或决定什么",
  "core_claim": "最重要的一条主张",
  "visual_concept": "由哪条洞察推出什么视觉概念",
  "pages": [
    {
      "id": "p1",
      "role": "open",
      "claim": "",
      "focus": "先看标题，再看副题",
      "focus_element": "t",
      "composition_rationale": "标题压在下三分之一，上方留连续暗场，重心靠底部横带稳定",
      "page_number_tone": "dark",
      "elements": [
        {"id": "bg", "type": "shape", "color": "#14181B",
         "x": 0, "y": 0, "w": 13.333, "h": 7.5, "overlap_allowed": true},
        {"id": "t", "type": "title", "content": "留存才是增长",
         "x": 0.9, "y": 4.62, "w": 9.6, "h": 1.28, "font_size": 64, "color": "#F3F0E8"},
        {"id": "sub", "type": "body", "content": "2026 年度复盘",
         "x": 0.9, "y": 4.0, "w": 6.0, "h": 0.44, "font_size": 20, "color": "#C9C3B6"}
      ]
    }
  ]
}
```

## 怎么写一份 plan

顺序是：先把每页的 `role`、`claim`、`focus` 写出来（这时还没有坐标），再排页序与节奏，然后逐页决定"用几个元素、谁承担主角"，最后才填几何。

- 几何先跑 `grid.py`，取列坐标而不是估数字；同一页的元素贴 2–4 条轴。
- 文字先按内容需要定字号档位，再用装载估算反推框高（下文「网格」末尾）。宁可改文案，不要改字号。
- 一个元素一个职责：一段文字只承担一种阅读角色；把"论断＋证据＋来源"塞进一个框，最后一定会出问题。
- 图表先写 `unit` `period` `basis` 和 `title`（说结论），再填数据。
- 写完先编译：硬错误会指出具体位置；警告逐条判断，不理解的不要忽略。

## 叠放与重叠

**叠放顺序就是声明顺序**：先写的在下。没有第二个开关。

背景图、蒙版、色块这类承载层，要有东西压在它上面时，在**它自己**上声明 `overlap_allowed: true`。压在上面的元素不能自我豁免——被压的那层才知道自己是不是承载层。实质性重叠（面积 ≥ 0.5 in²、短边 ≥ 0.05 in、占较小者 ≥ 25%）且下层未声明时阻断。

## 图表与表格

```json
{"kind": "bar",
 "title": "各渠道三月留存率",
 "unit": "%", "period": "2025 全年", "basis": "次月与第三月两次访问",
 "data_labels": true, "axis_labels": false, "highlight_category": 3,
 "data": {"categories": ["A", "B", "C", "D"],
          "series": [{"name": "留存", "values": [9, 11, 24, 31], "unit": "%"}]}}
```

- `kind`：`bar`（竖向簇状柱）、`line`（等距类别轴，不按真实日期定位）、`pie`、`donut`、`table`。
- `unit` `period` `basis` 会写进图题，是读者判断可信度的依据；同一张图里系列单位或期间与图表级声明矛盾即阻断。多系列共轴前必须 `shared_basis: true`。
- `axis_labels: false` 收起整根数值轴，必须同时 `data_labels: true`，否则读者没有读数。
- `highlight_category` 只用于单系列柱图：该类别用调色板首色，其余退到按底色兜底的中性色（只改明度不改色相）。
- `table` 的 `data` 用 `headers` / `rows`，`column_weights` 给列宽权重；表格用细横线分组，没有底色块。
- 饼／环只允许单系列、非负值。

## 素材

```json
"assets": {
  "hero": {"source": "自有素材库 m-01", "rights": "cleared", "reference": "编号 m-01",
           "generated": false, "used_as_evidence": false},
  "plate": {"source": "自建生成", "rights": "cleared", "generated": true,
            "generation": {"tool": "图像生成工具", "prompt": "实际提交的完整提示词",
                           "file": "plate.png", "status": "generated",
                           "sha256": "64 位十六进制"}}
}
```

- `source` 必填；`rights` 取 `cleared` / `pending` / `blocked`（版权与肖像、隐私合并判断）。`blocked` 阻断，`pending` 给提示——交付前确认，或在交付说明里写明。
- `used_as_evidence: true` 的素材必须是真实素材；生成图作证据直接阻断。
- 有 `generation` 就说明是生成图：`status` 必须是 `generated`，`file` 必须与页面 `image_id` 解析到**同一个物理文件**（同名不同扩展名不算），`sha256` 必须等于该文件真实哈希。核验器还会拿它与 PPTX 内嵌字节比对。
- 素材从 `--assets` 目录解析：先按 `image_id` 精确找，再补常见扩展名，最后按同名不同扩展名找。

## 页码

整稿一个契约，编译器按页序生成，不在页面里手写。默认右下、12pt、从 1 起、纯数字。

- `hide` 列出不显示页码的页；确需隐藏要在交付说明里写理由。
- 文字与图表不得与页码保留框相交（编译期阻断）——右下要长期放来源或图注时，整稿改 `corner: top_right`，不要逐页切换。
- 图片或色块压到角区是允许的：页码总是画在最上层。但它会提示你声明页面 `page_number_tone`（角区的实际明暗），页码墨色随之翻转；色块压角区会给警告，让你确认页码在它上面读得出来。
- 声明了 `overlap_allowed` 的承载层（背景图、蒙版）压到角区不报——它本来就在页码下面。

## 网格

```bash
python scripts/grid.py --span 1:6 --span 8:5     # 默认 16:9、12 列、边距 0.7/0.6、沟槽 0.25
python scripts/grid.py --json                    # 含页码保留框，供脚本消费
```

先定 2–4 条主对齐轴（左边缘、内容区起点、右栏起点、底部来源线），元素贴轴。跨列要对应内容关系：论断＋证据常见 5+空档+6，真正等重的对照才等分。空列是资产——空档应连续并承担分组、停顿或回秤。品牌或媒介另有规范时改参数，别迁就默认值。

**装载估算**（编译期会用同一模型给提示）：每行满宽字数 ≈ `(w − 0.1) ÷ (字号 ÷ 72)`，中文一字一 em、西文数字约 0.55 em。需要的高度 ≈ 行数 × 字号 × 1.2 × 行距 ÷ 72。估不准就以真实渲染为准。

## 命令

```bash
python scripts/compile.py plan.json out.pptx [--assets DIR] [--report compile.json]
python scripts/verify.py plan.json out.pptx [--render-dir render] [--report verify.json]
python scripts/inspect_pptx.py 原稿.pptx [--json inspect.json]
python scripts/grid.py [--aspect 4:3] [--cols 12] [--span 1:6] [--corner top_right] [--json]
```

`--assets` 缺省取 plan 同级的 `assets/`，没有则取 plan 所在目录。`verify.py` 的 `--render-dir` 缺省为 pptx 同级的 `render/`，会写 PDF、逐页 PNG 与 `contact_sheet.png`。

退出码：`compile.py` 0 成功（可能有警告）、2 硬错误未写文件、3 运行期失败（另写 `.error.pptx`，不是交付物）；`verify.py` 0 无阻断、1 有阻断、2 输入不可读。

## 诊断码

编译期：

| 码 | 级别 | 含义 |
| --- | --- | --- |
| `E-SCHEMA` | 硬 | JSON 语法或字段形状不合契约 |
| `E-ID` | 硬 | 页 id 或元素 id 重复 |
| `E-REF` | 硬 | `focus_element`、`image_id`、`runs` 拼接等引用不成立 |
| `E-GEOM` | 硬／提示 | 元素完全在画布外（硬）；文本框伸出画布（提示） |
| `E-OVERLAP` | 硬 | 实质重叠且下层未声明 `overlap_allowed` |
| `E-CHART` | 硬／提示 | 载荷长度、口径声明、饼图约束、轴与直标条件、列宽权重不一致（硬）；调色板不够（提示） |
| `E-ASSET` | 硬／提示 | 权利 blocked、生成图作证据、生成登记与哈希不符（硬）；`pending`（提示） |
| `E-IMG` | 硬 | 素材文件找不到 |
| `E-PAGE` | 硬／提示 | 文字或图表与页码保留框相交（硬）；色块压到角区（提示） |
| `I-PAGE` | 说明 | 图片压到页码角区，需按实际明暗声明 `page_number_tone` |
| `E-RUNTIME` | 硬 | 运行期异常；不覆盖正式输出，只写 `.error.pptx` |
| `W-FIT` `W-PT` `W-CONTRAST` | 提示 | 装载估算偏紧、字号低于下限、声明层面对比不足 |
| `W-PPI` `W-CROP` | 提示 | 有效 ppi 不足、裁切丢掉过多原图（画幅或落位框选错） |
| `I-CONTRAST` | 说明 | 文字下方是图片，静态算不出对比，需看渲染 |

核验期：`V-PAGES` `V-OBJ` `V-GEOM` `V-TEXT` `V-STYLE` `V-ORDER` `V-ASSET` `V-CHART` `V-CONTRAST` `V-PAGENO` `V-CLIP`（硬），`V-OVERFLOW` `V-BLANK` `V-FONT`（提示），`V-RENDER`（说明，缺工具）。

提示不是必须改的错。它标出"这里可能有问题"，改不改由你判断——但不要说没看见。读提示的经验：`W-FIT` 先改文案再加框；`W-CONTRAST` 先换墨色或换落点，不靠加浓遮罩；`W-PPI` 换素材或缩小显示尺寸；`I-CONTRAST` 只能看渲染；`E-OVERLAP` 九成是忘了给承载层声明 `overlap_allowed`。

## 机器边界

- 原生图表只有柱、线、饼、环与表格。没有 XY 散点、桑基、自动甘特、决策树、绑定端点的连接器。其余关系用文字、基本形状、分隔条手工组合：这些对象可单独编辑，但点的位置和带宽不与数据绑定，改数据要重新算、重新连线、重新渲染。
- 不生成演讲者备注、动画、页面切换。页 `notes` 只是设计备注。
- 页标题是普通文本框，不是标题占位符：PowerPoint 大纲视图与读屏软件识别不到。无障碍有要求时在交付说明里提示。
- 字体只写字体名，不嵌入字体文件。
- 图片按 `cover` 等比裁切到给定框，`crop_focus` 指定裁窗中心；不保证主体不被裁到——那是选图与框的事。
- 核验器与编译器共用同一套 plan 模型、文本式样解析与阈值（`compile.py`）。共用部分若有缺陷，两边会一起"一致"而发现不了；所以对象核对不能替代看图，`verify.py` 的系列色是从成品 XML 实际读取的，不从 plan 回算。

## 报告

`--report` 写出的 JSON 含 `status`、`sha256`、`slides`、`render`（PDF、PNG、字体）、`findings`。把它当作证据材料附在交付里，不要当成绩单。

`sha256` 用来区分"这份文件是不是我验的那一份"，不是复现凭据：同一份 plan 编译出的对象与几何一致，但图表内嵌的工作簿带时间戳，字节不会完全相同。
