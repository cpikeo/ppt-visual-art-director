# 执行契约（写 deck.json 之前查一次）

判据在 `../SKILL.md`；阈值在 `../policy/strategy.yaml`；结构在 `../schema/deck.schema.json`。
本文只讲「字段怎么读、什么会被当场拦住」。

## 1. 三份文件的关系

| 文件 | 谁写 | 作用 | 指纹 |
|---|---|---|---|
| `plan.json` | `pvad.py plan` | 判断 + 依据 + 未决问题 | `policy_sha256` + `brief_sha256` |
| `deck.json` | `pvad.py build` 或作者 | 几何 / 字号 / 数据 | `deck_digest`（manifest 里） |
| `manifest.json` | `pvad.py check` | 证据与结论 | `output_sha256` |

`build` 会拒绝对不上指纹的 plan：brief 或策略改过，就必须重跑 plan。

## 2. deck 顶层

```jsonc
{
  "schema": "pvad-deck-v1",
  "canvas": {"width": 1280, "height": 720, "ppt_dpi": 96},
  "theme":  {"colors": {"background": "#F7F5F1", "ink": "#14161A", ...},
             "fonts": {"cn": "Noto Sans CJK SC", "latin": "DejaVu Sans"}},
  "meta":   {"metric_units": {"单店月产出": ["万元"]}},   // build 生成：跨页同指标单位登记
  "slides": [{ "id": "s01", "composition": "two_end", "layout_source": "plan",
               "meta": {"source": "...", "unit": "万元", "period": "...", "basis": "...",
                        "has_numeric_evidence": true, "claim_absent": false,
                        "metrics": [{"label": "...", "value": "4.6", "unit": "万元"}]},
               "elements": [ ... ] }]
}
```

- `theme.colors` 必须齐 policy 里的全部 token；少一个是 BLOCK（缺失会静默回落）。
- `theme.fonts` 最多两个家族（中文一个、拉丁一个）。
- `slide.meta.emphasis_applied` 记录「强调档有没有真的落到字号」；装不下时就把它写在
  `slide.meta.emphasis_notes` 里——判断成立但物理不允许时，如实降级而不是硬塞；
- `slide.meta` 是证据位：`release` 模式要求含数字的页把 `source / unit / period / basis` 配齐；
  且每个指标自己要有单位（不靠 deck 级单位兜底——把「万元」套到「门店数量」上就是口径事故）。

## 3. 元素信封与字段

```jsonc
{ "id": "s03.left_value", "type": "text", "role": "left_value",
  "x": 72, "y": 304, "width": 456, "height": 106,
  "text": "4.6\u00a0万元", "size": 88, "line_height": 1.06, "tracking": -1.4,
  "color": "primary", "bold": true, "align": "left", "anchor": "bottom",
  "max_lines": 3, "wrap": true, "layout_source": "plan" }
```

| type | 专属字段 |
|---|---|
| `text` | `text` `size` `color` `bold` `italic` `align` `anchor` `line_height` `tracking` `max_lines` `wrap` `font` |
| `shape` | `shape`(rect/rounded_rect/ellipse/triangle/diamond/line/arrow) `fill` `stroke` `stroke_width` `text` `text_size` `text_color` |
| `chart` | `chart_kind`(column/bar/line/donut/progress) `categories` `series[].name/.values` `show_values` `value_size` `label_size` `unit_suffix` |
| `image` | `source`（相对 brief 目录或绝对路径）`fit`(cover/contain) `function`(hero/context/proof/emotion/frame) `reason` `safe_area` `text_color`(dark/light) `overlay`(0–1) `crop`/`crop_loss`（build 自动算）`asset_sha256` `source_px` `format` |

- 坐标是**画布 px**（96dpi 基准），pt = px × 72 / 96。
- 线元素允许 `width` 或 `height` 为 0；其他元素退化即 BLOCK。
- 颜色只接受 `#RRGGBB` 或 theme token；写错名字会被拦——因为它不会报错，只会静默回落。
- **字段名写错＝静默丢弃**：`font_size`、`maxlines`、`textColor` 这类近似名会被认出并拦下（`FIELD_UNKNOWN`）。

## 4. 构图与数据的一致性

- `chart_kind` 与数据形状必须匹配：环图只接受单序列、非负、总和 > 0；
  序列长度必须等于类别数；非有限值一律拦。
- 环图合计不是 100 时告警 `DONUT_NOT_PERCENT`：量不该画环。
- 没有时间读数却声明 `series_chart`，`build` 会降级成 `metric_row` 并在
  `composition_downgraded` 写清理由——**不画空轴**。

## 5. 失败码（BLOCK）

| code | 含义 | 怎么修 |
|---|---|---|
| `SCHEMA_FAIL` | 结构不符 / 必填缺失 | 按 `schema/deck.schema.json` 改 |
| `GEOMETRY_INVALID` | 非有限值或退化几何 | 线元素才允许 0 尺寸 |
| `OUT_OF_CANVAS` | 越出画布 | 改构图，不缩放画布 |
| `ELEMENT_ID_DUPLICATE` | 元素/页面 id 重复 | 对账与取证都依赖唯一 id |
| `TEXT_OVERFLOW` | 行数或墨迹超出框 | 删内容 / 加高框，不缩字号 |
| `TEXT_TOO_WIDE` | 单行宽于框宽 | 改短文本或加宽 |
| `TEXT_SIZE_FLOOR` | 低于语义档下限 | 删内容 |
| `ELEMENT_OVERLAP` | 墨迹相交 | 重排，不靠缩间距 |
| `TOKEN_UNKNOWN` | 颜色名不认识 / token 缺失 | 只写 token 或 #HEX |
| `FIELD_UNKNOWN` | 疑似写错字段名 | 按第 3 节字段表 |
| `CHART_INVALID` | 载荷非法 | 检查类别/序列/数值 |
| `PROVENANCE_MISSING` | release 缺来源/单位/期间/口径 | 补齐，或删掉该证据 |
| `UNIT_INCONSISTENT` | 同指标跨页单位不一致 | 统一口径或拆指标 |
| `ACCENT_OVER_BUDGET` | 强调色超预算 | 只给语义落点 |
| `FONT_FAMILIES` | 字体家族 > 2 | 统一字体角色 |
| `EMPTY_DECK` / `EMPTY_SLIDE` | 空稿 / 空页 | 删页或补内容 |
| `LAYOUT_DRIFT` | 回读与 spec 不一致 | 这是编译器缺陷：修编译，不要改稿迁就 |
| `ASSET_MISSING` | 素材文件不存在 / 判断要图但页面没图 | 补文件或删图位；缺图不能静默通过 |
| `ASSET_FORMAT` | 后缀或真实格式不在白名单 | 转 png/jpg/webp；改后缀不改格式无效 |
| `ASSET_RESOLUTION` | 源图长边 < 放置尺寸 × oversample | 换大图或缩小图位，不放大糊图 |
| `ASSET_CROP` | cover 裁掉过多且没写 reason | 换图/换图位，或写明裁切理由 |
| `ASSET_SAFE_AREA` | 文字压图但没声明 safe_area | 先定文字落点 |
| `ASSET_CONTRAST` | 压图文字实测对比度/亮度差不足 | 改文字色、加遮罩、换图位 |
| `ASSET_BUDGET` | 单页图片超上限 | 图片不靠数量建立氛围 |
| `ASSET_DECLARATION` | 图片与页面判断的职责不一致 / text_color 非法 | 补职责或删图 |
| `CHART_MIXED_UNITS` | 同一张图混了两种单位（家 / %） | 拆图或改成两个指标块：不可换算的量不能共用一根轴 |

告警（不拦，但必须读）：`FOCUS_COMPETITION`、`DUPLICATE_TEXT`、`SAFE_AREA`、
`DONUT_NOT_PERCENT`、`CHART_SINGLE_CATEGORY`、`ORPHAN_LINE`、`FORBIDDEN_LINE_START/END`、
`DENSITY_OVER_TARGET`、`MEASUREMENT_DEGRADED`、`ASSET_TEXTURE`、`ASSET_OUTSIDE_PROJECT`。

## 5.0 素材入参（brief.asset）

| 字段 | 必填 | 作用 | 缺了会怎样 |
|---|---|---|---|
| `path` | ✅ | 文件路径（相对 brief 目录）；还没生产的也照写，作为交付目标 | `ASSET_MISSING` |
| `function` | — | `hero/context/proof/emotion/frame`，决定职责句 | 默认 `context` |
| `reason` | — | 拿掉这张图会损失什么 | 装饰性素材报 `ASSET_DECLARATION` |
| `subject` | — | **画面里出现什么**（生产说明的主体句） | `PROMPT_SUBJECT_MISSING`，不猜关键词 |
| `avoid` | — | 作者禁项（品牌标识、可辨识人脸…），追加进反向词 | 只留策略默认反向词 |
| `fit` | — | `cover`（默认）/`contain` | 按 function 推导 |
| `safe_area` | — | 文字压图的安全区（比例坐标；数组或对象都行） | 文字压图时 `ASSET_SAFE_AREA` |
| `text_color` | — | `dark`/`light`：压在图上的字用深还是浅 | 生产说明不发调子句；交付时 `ASSET_CONTRAST` |

## 5.1 素材契约（图片）

```jsonc
{ "id": "s01.media", "type": "image", "role": "media",
  "x": 712, "y": 56, "width": 496, "height": 624,
  "source": "assets/warehouse.jpg",     // 相对 brief 所在目录
  "fit": "cover", "function": "hero",
  "reason": "让观众看见渠道现场的真实密度……",   // 拿掉它，理解会损失什么
  "safe_area": {"x": 0.0, "y": 0.10, "width": 0.52, "height": 0.80},
  "text_color": "dark",
  "crop": [0, 0.12, 0, 0.12], "crop_loss": 1.36,   // build 自动计算
  "asset_sha256": "…", "source_px": {"width": 1536, "height": 1024} }
```

- `text_color` 是**语义提示**（图上文字用深色还是浅色），不是颜色 token；
- `crop` 是 cover 的真实裁切比例，编译时写入 PPTX（不裁就是拉伸变形，`fidelity` 会报）;
- `asset_sha256` 参与回读对账：包里的图必须与 spec 记录的素材是同一个文件；
- `brief`（仅当文件缺失/解码失败时出现）是生产说明，见下。

```jsonc
"brief": {                              // 只在「还缺这张图」时写入
  "schema": "pvad-prompt-v1",
  "prompt": "…", "negative": "no text, …",
  "clauses": [{"slot": "role", "text": "…"}, …],   // 每句出自哪条策略，可审计
  "ratio": "4:5", "target_px": {"width": 992, "height": 1248},
  "quiet_side": "left", "text_color": "dark",
  "fingerprint": "prompt-…", "issues": [] }
```

`pvad.py media deck.json --out media_briefs.json` 把全部图位导出成一份清单：

| 字段 | 含义 |
|---|---|
| `status` | `to-produce`（缺文件）/ `broken`（解码失败）/ `ready`（文件已在） |
| `prompt` / `negative` | 交给摄影师或图像模型的说明；`ready` 的条目不需要 |
| `ratio` / `target_px` | 画幅与最小源像素（图位 × oversample，且不低于长边下限） |
| `fingerprint` | 说明内容的哈希：内容没变就不用重出图 |

## 6. manifest 键

```
status                     PASS / BLOCKED（先写 BLOCKED，全链完成才改）
mode                       spec / draft / release
deck_sha256 / spec_digest  稿与规格的指纹
policy                     {policy_sha256, canvas, key_thresholds}
guard                      {status, errors[], warnings[], fix_groups[], stats, degraded}
fidelity                   {status: MATCH/MISMATCH, matched_elements, summary, mismatches[]}
pixel                      {pages[], worst_ratio, limit, status}
previews                   {spec_sheet, pptx_sheet, pages[]}
output_pptx / output_sha256
verification.target_powerpoint           恒为 NOT VERIFIED
verification.target_renderer             恒为 NOT VERIFIED
verification.independent_visual_review   恒为 NOT RUN（除非盲评真跑过）
```

交付时同时核对：**退出码 + manifest.status + output_sha256**。三者任一不符，不得宣称成功。

## 7. 文本测量边界

`guard` 的字号、折行、字距与段距是**真实字体度量**（PIL 字形前进宽度），但：
- 字体文件缺失时退化为估算，并把 `MEASUREMENT_DEGRADED` 写进报告——估算不得当作测量结论；
- 中文断行模型与目标 PowerPoint 的实际换行可能不同（尤其字距、标点宽度）；
- 因此「墨迹不相撞」是必要条件，不是视觉验收。
