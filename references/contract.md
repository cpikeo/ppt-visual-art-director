# 数据契约：slide plan 字段、元素类型与失败码

本文件是 BUILD 阶段的唯一结构契约，按需读取一次。它只描述数据结构及能明确表达的约束，**不评价审美、不替代设计判断**。审美是否成立、叙事是否有力，由 AI 判断与独立评测负责。

## 1. 顶层：slide plan

```json
{
  "deck_id": "string, 必填，演示唯一标识",
  "title": "string, 必填，演示标题",
  "audience": "string, 必填，受众",
  "goal": "string, 必填，演示目标",
  "core_claim": "string, 必填，核心主张（唯一要被记住的一句话）",
  "evidence": "string[], 可选，支撑主张的证据清单",
  "visual_concept": "string, 可选，整体视觉概念（色彩关系、摄影语言、材质基调）",
  "palette": "string[], 可选，配色（HEX，2–5 色）",
  "fonts": { "heading": "string", "body": "string" },  // 可选，中文优先字体
  "aspect_ratio": "string, 可选，'16:9' | '4:3'，默认 16:9",
  "pages": "Page[], 必填，按叙事顺序排列",
  "rhythm": "string[], 可选，各页节奏标签 open|build|breathe|turn|hold|close",
  "asset_reason": "string[], 可选，跨风格混排或素材选择的表达理由"
}
```

`pages` 长度即页数；不设图位额度，不编造信息。

## 2. 页面：Page

```json
{
  "id": "string, 必填，页面唯一标识",
  "role": "string, 必填，页面职责：open | build | breathe | turn | hold | close",
  "claim": "string, 必填，本页唯一主张（话题标签时留空并记 open question）",
  "evidence": "string[], 可选，本页证据",
  "layout": "string, 可选，构图方向：large_white | editorial | data_field | image_story | linear | asymmetric",
  "density": "string, 可选，密度目标：sparse | moderate | dense",
  "emphasis": "string, 可选，视觉分量：low | medium | high",
  "elements": "Element[], 必填，本页元素",
  "notes": "string, 可选，给 AI/作者的判断备注与回退说明"
}
```

## 3. 元素：Element

```json
{
  "id": "string, 必填，元素唯一标识",
  "type": "string, 必填，元素类型：title | body | text | number | image | chart | shape | divider | caption | footer",
  "content": "string, 可选，文本内容（文本类元素必填）",
  "x": "number, 必填，左上角 X（英寸，相对幻灯片）",
  "y": "number, 必填，左上角 Y（英寸）",
  "w": "number, 必填，宽度（英寸）",
  "h": "number, 必填，高度（英寸）",
  "font_size": "number, 可选，字号（磅，文本类元素）",
  "color": "string, 可选，HEX 颜色",
  "align": "string, 可选，left | center | right | justify",
  "image_id": "string, 可选，图片类元素引用的资产 ID",
  "asset_reason": "string, 可选，该素材的表达理由",
  "chart": "ChartSpec, 可选，图表类元素的数据规范",
  "z_index": "number, 可选，叠放顺序，默认按元素顺序",
  "overlap_allowed": "boolean, 可选，是否允许与其他元素实质重叠（设计授权的图层融合，如文字叠图、遮罩）。默认 false；仅在设计意图明确时设为 true"
}
```

## 4. 图表规范：ChartSpec（单一规范）

```json
{
  "kind": "string, 必填，bar | line | pie | donut | table",
  "data": "object, 必填，图表数据（数字、单位、期间、口径同源）",
  "title": "string, 可选，图表标题",
  "unit": "string, 可选，单位",
  "period": "string, 可选，期间"
}
```

图表形状来自证据不来自装饰；图例、网格、配色不抢数据。

## 5. 失败码（硬错误，程序阻断）

| 失败码 | 含义 | 处置 |
| --- | --- | --- |
| `E001` | 缺少必填字段（deck_id/title/audience/goal/core_claim/pages） | 阻断编译 |
| `E002` | 页面元素超出幻灯片边界或互相重叠 | 阻断受影响页面 |
| `E003` | 引用的 image_id/资产不存在 | 阻断该元素，提示作者 |
| `E004` | 图表数据缺失或结构非法（无法自动验证语义一致性） | 阻断，反馈 AI；语义一致性需人工或 AI 复核 |
| `E005` | 字号过低导致不可读（中文正文 < 18pt，标题 < 28pt） | 警告，不阻断 |
| `E006` | 页面元素数量超过阈值（> 12） | 警告，建议删减 |
| `E007` | 图表渲染失败（数据或类型非法） | 阻断该图表元素，提示作者 |
| `E008` | 图表数值与页面主张可能存在语义不一致 | 警告，不阻断，建议 AI 或人工复核 |

非法输入按 `E001–E002` 处理；`E003–E007` 为客观检查，按实际风险判定是否阻断；`E008` 为语义一致性软提示，仅建议复核，不阻断编译。

## 6. 归一化规则

- 所有长度单位统一为英寸；角度统一为度；颜色统一为 HEX `#RRGGBB`。
- 缺失可选字段使用契约默认值，不自动补全必填字段。
- 程序按契约执行，不因实现方便而重新解释设计意图。
