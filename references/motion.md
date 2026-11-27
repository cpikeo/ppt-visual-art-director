# 动态艺术指导：Motion Grammar

动态服务叙事。每个运动都要能回答：**为什么在这里、对谁、先后依据是什么**；答不出就不动。运行时取值见 `web/deck.js`，这里给判断。

## 一、Motion Grammar（每条运动写满七项，缺项即未设计）

| 项 | 回答 | 例 |
| --- | --- | --- |
| 意图 | 建立焦点｜呈现数量｜揭示关系｜切换场景｜纵深与呼吸 | 呈现数量 |
| 对象 | 谁在动（一个元素或一组） | 巨型数字 |
| 状态 | 起止状态（终态必须是可读的完整版面） | 0 → 41 |
| 触发 | 进入页自动｜点击构建（`data-step`）｜交互 | 进入页 +0.4s |
| 时序 | 延迟、持续、先后（`data-at` `data-dur` `data-stagger`） | 1.1s |
| 缓动 | 进入用 ease-out（expo），转场用 in-out，匀速只给 drift | cubic-out |
| 转场关系 | 与前后页的运动是否同向、是否升维 | 上页 push → 本页 zoom |

## 二、预设与意图（`data-m`）

| 意图 | 预设 | 默认 | 判断 |
| --- | --- | --- | --- |
| 建立焦点 | `rise` `mask`（行揭示）`scale` | 720 ms expo-out | 只给主论断，同页最多 2 个 |
| 呈现数量 | `count` `grow` `growx` | 1100 / 720 ms | 动画终值必须等于文本终值；`count` 读 `textContent` 的首个数字 |
| 揭示关系 | `draw`（SVG 描边）`wipe`（`data-dir` l/r/u/d） | 720 ms | 方向 = 阅读方向或因果方向，别反 |
| 纵深与呼吸 | `drift`（`data-drift="x,y,scale"`） | 16 s linear | 背景层，位移 ≤ 5%，不抢前景 |
| 次要进入 | `fade` `slide` `sink` | 720 ms | 注释、标签；先主后次 |

页间转场（`<section data-transition>`，缺省取 `<html data-transition>`）：`push` 推进（同一叙事继续）｜`zoom` 升维/聚焦（进入大数字、局部）｜`wipe` 转折（`turn` 页）｜`fade` 呼吸（`breathe` / `open`）｜`lift` 收束（`close`）｜`cut` 无转场。760 ms，in-out；`prefers-reduced-motion` 一律即时切换。

## 三、节奏与连续性

1. **整稿一个主节拍**：转场时长、缓动、方向固定，换的是“种类”而不是“速度”；最重的一页允许一次例外（更慢或更静）。
2. **进入编排 ≤ 3.2 s**（`web_export.mjs audit` 检查）：观众不应等动画；页内先主后次，间隔 0.08–0.15 s。
3. **构建步骤服务演讲**：`data-step=n` 点击出现，退回即隐藏；一页最多 5 步；不放装饰性构建。
4. **同一对象跨页同一种动**：数字都 `count`，分隔线都 `wipe`，否则像各页独立炫技。
5. **静止也是动态设计**：`breathe` / `hold` 页只动一个元素，或不动。
6. **只用 transform / opacity / clip-path / stroke-dashoffset**，不触发布局；同屏同时运动的元素 ≤ 8。

## 四、降级与可访问性（硬约束）

- 内容默认可见；动效由 Web Animations API 叠加，JS 失败或 `Deck.errors>0` 回落终态（`html.dk-safe`），**不得用 CSS 预先隐藏内容**。
- `prefers-reduced-motion`、`?still`、`?export`：全部终态 + 即时切换。
- 切后台、快速连按：转场被完成而不是堆叠；`settle()` 保证任何时候可回到终态。
- 动画不承载唯一信息（数字终值、结论必须在静态版中完整）。
- 导出：动画与转场不进入 PPTX/PDF，导出的是终态；报告里写明。
