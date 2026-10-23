# 数据契约：slide plan 字段、元素类型与失败码

本文件是 `scripts/compile_pptx.py` 的唯一结构契约，仅在 BUILD 阶段需要编译 `plan.json` 时读取一次。它只定义可执行的数据结构与客观可检约束，**不评价审美、不规定固定版式、不替代 AI 设计判断**。

## 1. 顶层：slide plan

```json
{
  "deck_id": "string, 必填，演示唯一标识",
  "title": "string, 必填，演示标题",
  "audience": "string, 必填，受众",
  "goal": "string, 必填，演示目标",
  "core_claim": "string, 必填，整稿核心主张（唯一要被记住的一句话）",
  "evidence": "string[], 可选，支撑核心主张的关键证据清单",
  "visual_concept": "string, 可选，整体视觉概念（明度基调、色彩关系、字体气质、图像语言）",
  "palette": "string[], 可选，调色板（HEX `#RRGGBB`，图表系列按序优先消费）",
  "fonts": { "heading": "string", "body": "string" },
  "aspect_ratio": "string, 可选，'16:9' | '4:3'，默认 '16:9'",
  "pages": "Page[], 必填，按叙事顺序排列的非空数组",
  "rhythm": "string[], 可选，整稿各页节奏摘要",
  "asset_reason": "string[], 可选，整稿素材选择或混排依据"
}
```

`pages` 长度即实际页数；不设固定图位配额，不凭空编造信息。

## 2. 页面：Page

```json
{
  "id": "string, 必填，页面唯一标识（整稿内不重复）",
  "role": "string, 必填，页面节奏职责：open | build | breathe | turn | hold | close",
  "claim": "string, 必填字段（字符串）；有明确主张时写结论句；原标题仅为话题标签且无证据可立主张、或纯过渡/呼吸页时置空字符串 \"\"，并在 notes 记录 open question 或留空依据",
  "evidence": "string[], 可选，本页支撑证据",
  "layout": "string, 可选，构图倾向标记：large_white | editorial | data_field | image_story | linear | asymmetric",
  "density": "string, 可选，信息密度：sparse | moderate | dense",
  "emphasis": "string, 可选，视觉分量：low | medium | high",
  "background": "string, 可选，本页背景色 HEX `#RRGGBB`（默认白色）",
  "elements": "Element[], 必填，本页元素列表",
  "notes": "string, 可选，设计判断备注、open question、证据缺口或回退说明"
}
```

## 3. 元素：Element

```json
{
  "id": "string, 必填，元素唯一标识",
  "type": "string, 必填，title | body | text | number | image | chart | shape | divider | caption | footer",
  "content": "string, 文本类元素（title/body/text/number/caption/footer）必填",
  "x": "number, 必填，左上角 X（英寸，≥ 0）",
  "y": "number, 必填，左上角 Y（英寸，≥ 0）",
  "w": "number, 必填，宽度（英寸，> 0）",
  "h": "number, 必填，高度（英寸，> 0）",
  "font_size": "number, 可选，字号（磅）",
  "bold": "boolean, 可选，是否加粗（默认 title、number 为 true，其余为 false）",
  "color": "string, 可选，前景文字或形状填充 HEX `#RRGGBB`",
  "align": "string, 可选，left | center | right | justify",
  "shape_kind": "string, 可选，shape/divider 几何类型：rect | round_rect | oval，默认 rect",
  "image_id": "string, image 元素必填，相对于 --assets 目录的资源名（可带或不带扩展名）",
  "asset_reason": "string, 可选，该素材的表达职责说明",
  "chart": "ChartSpec, chart 元素必填，原生图表或原生表格数据规范",
  "z_index": "number, 可选，叠放次序（小值在下、大值在上；不自动免除重叠检查）",
  "overlap_allowed": "boolean, 可选，设计授权的图层融合（如底图、遮罩、容器承载前景文字），默认 false"
}
```

## 4. 图表与原生表格规范：ChartSpec

```json
{
  "kind": "string, 必填，bar | line | pie | donut | table",
  "data": "object, 必填；图表使用 {\"categories\": [...], \"series\": [{\"name\": \"...\", \"values\": [...]}]}；table 同时支持 {\"headers\": [...], \"rows\": [[...], ...]} 或上述 categories/series 结构",
  "title": "string, 可选，图表/表格标题",
  "unit": "string, 可选，数据单位（自动并入标题标注，防止无单位裸数）",
  "period": "string, 可选，统计期间（自动并入标题标注）"
}
```

多指标单位不同或口径不可共轴时，应选用 `"kind": "table"` 渲染原生可编辑表格（或拆分为独立数字模块），不得强画在同一柱线坐标轴上。

## 5. 失败码与级别

| 失败码 | 默认级别 | 触发条件 | 程序处置 |
| --- | --- | --- | --- |
| `E001` | `HARD` / `SOFT` | 缺少必填字段、类型非法、页面 ID 重复（`HARD`）；枚举值非法、`claim` 为空且未在 `notes` 说明原因（`SOFT`） | `HARD` 阻断编译；`SOFT` 告警（`--strict` 下阻断） |
| `E002` | `HARD` | 元素宽高 `≤ 0`、超出当前画幅边界（16:9 为 `13.333×7.5`，4:3 为 `10.0×7.5`），或非全屏背景元素在未声明 `overlap_allowed=true` 时发生实质重叠 | 阻断编译，返回精确元素与重叠面积 |
| `E003` | `HARD` | `image` 缺失 `image_id`，或在资产目录中找不到对应图片文件 | 绘制带标注占位框供替换，标记 `HARD` 并返回非零退出码 |
| `E004` | `HARD` | `chart` 元素缺失 `chart` 或 `chart.data` | 阻断编译，反馈 AI 补齐或改用其他表达 |
| `E005` | `SOFT` | 投影可读性风险：中文 `title < 28pt`，`body/text/number < 18pt`，`caption/footer < 12pt` | 警告不阻断（允许辅助注释合理降级于正文） |
| `E006` | `SOFT` | 单页元素数 `> 12` | 警告不阻断，提示核查是否存在碎片化堆砌 |
| `E007` | `HARD` | `chart.kind` 非法、序列为空/长度不匹配、数值不可解析或原生图表/表格创建失败 | 绘制失败占位框，标记 `HARD` 并阻断通过 |
| `E008` | `SOFT` | 页面 `claim` 中的数值无法由图表/表格原始值、比例、求和、差值或涨跌幅推导对应 | 警告不阻断，提示 AI 或人工复核口径一致性 |

## 6. 归一化与执行边界

- 长度单位统一为英寸，字号单位为磅（pt），颜色格式为 `#RRGGBB`。
- 图片置入目标框时采用 OOXML 非破坏性等比居中裁切（`srcRect`），保持目标版面边界且绝不拉伸变形原图。
- 文本统一同时写入西文 `<a:latin>` 与东亚 `<a:ea>` 字体定义，确保中文字体设定真实生效。
- 原生表格（`kind: "table"`）与图表默认遵循无外框、透明底色场域、列对齐与水平发丝分割线（自适应深浅页面背景），避免默认重色表头块、斑马底纹与竖向网格框干扰信息阅读。
