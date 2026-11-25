# 既有稿定点编辑

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
| `text` | `content` `runs` `space_after_pt`，以及样式键（只改字，不得引入原稿没有的样式）；目标必须是文本对象，表格/图表/图片/组合会被 `E-OPS` 拒绝 |
| `box` | `x` `y` `w` `h` |
| `style` | `font` `font_ea` `font_size` `bold` `color` `letter_spacing_pt` `align` `line_spacing` `pad_in` `valign` |
| `fill` | `color` `alpha` |
| `line` | `color` `width_pt` `dash` `hide` |
| `order` | `to`（`back`／`front`） |
| `image` | `image_id` |
| `chart_data` | `data`（不改系列数与系列名；组合图不支持） |
| `delete` | — |
| `background` | `color`（按页） |

- 名称重复即阻断，改用 `target_index`。目标在组合对象里时加 `in_group: "<组名>"`：被点名的子对象按操作授权改动，同组其余子对象逐个与原稿比对，任何变化即 `E-PRESERVE`；`order` / `delete` 对子对象同样走这条路径，但不改组本身。未知键直接拒绝。
- **富文本保真**：段落属性与 run 属性整份沿用，超链接保留；原文含超链接而 run 结构对不上，拒绝写出（`E-OPS`）。含域或软回车（`a:fld` `a:br`）的文本框不做 `text` 改写。
- **保留核对**：逐对象 XML（exclusive C14N）、标识与样式容器、超链接地址、页面非形状结构、关系、逐页特性、包内部件指纹。
- **输出**只能是 `.pptx`；原稿含宏即拒绝（`E-FORMAT`）。
- **边界**：不新增页、不复制页、不改页序、不改系列数、不动母版与版式、不加动画与切换。
