# 动画与切换（V23）

读本文：要给整稿加页面切换、对象动画、图表/文本分步构建，或改既有稿的动画。**默认无动画**；只有 plan 显式声明才编译。动画与静态构图由同一套判断驱动：先问“静态构图能不能说清楚”，能就不加。

## 一、该不该动（先于任何字段）

| 动画只用于 | 例 | 不该用 |
| --- | --- | --- |
| 揭示（分层出现的论证） | 三层论据逐层出现 | 一个元素一次点击 |
| 比较（并置前后） | 图表按系列出现，先基线后新值 | 全页统一飞入 |
| 解释（路径叙事） | 一个点沿路径走完流程 | 循环、弹跳、闪烁、无意义旋转 |
| 聚焦（强调一个对象） | 关键数字放大后复位 | 同时强调多个 |
| 转场（章节语法） | 章节间同一种 fade/push | 页页换切换 |

每个 `animation` 必填 `purpose`（reveal｜compare｜explain｜focus｜transition）——写不出“这次点击揭示什么”就删掉。每页建议 ≤ 3 个单击阶段、≤ 8 个动画、累计 ≤ 6 s（`limits.json` → `motion`，均为 SOFT）。整稿切换种类建议 ≤ 2。**减少动态**：`compile.py --motion reduced`（只留 appear/fade/wipe，≤ 400 ms，去掉位移/旋转/缩放/路径/重复，切换只留 cut/fade/dissolve）；`--motion static` 去除全部动画与切换，页面保持静态可读。`verify.py --motion` 须与编译时一致。

## 二、契约（`plan.schema.json` → `$defs.transition` / `$defs.animation`；字段表用 `scaffold.py fields page`）

页级：`transition`（对象）与 `animations`（有序数组，顺序 = 动画窗格顺序）。

`transition`：`type` cut｜fade｜dissolve｜push｜wipe｜split｜cover｜uncover｜circle｜diamond｜plus｜wedge｜wheel｜zoom｜checker｜blinds｜comb｜random_bar｜newsflash｜random；`direction`（取 ECMA-376 原值：push/wipe `l r u d`，cover/uncover 另有 `lu ru ld rd`，split/zoom `in out`，checker/blinds/comb/random_bar `horz vert`）；`speed` slow｜med｜fast；`duration_ms` 100–10000（写成 `mc:AlternateContent` + `p14:dur`，Fallback 保留 `spd`）；`advance_on_click`、`advance_after_ms`（自动前进；禁用点击必须给它）；`spokes`（仅 wheel）；`through_black`（仅 cut/fade）。切换与对象动画是两套节点（`p:transition` ≠ `p:timing`），互不替代。

`animation`：`id` `target`（**顶层元素 id**；group 整体可作目标，group 子元素不可）`effect` `purpose` 必填。

| class（缺省按 effect 推断，仅 `exit` 须显式） | effect | 实现（OOXML） | 边界 |
| --- | --- | --- | --- |
| entrance | appear / fade / fly / wipe / zoom | `p:set` visibility ＋ `animEffect` / `anim ppt_x,ppt_y` / `anim ppt_w,ppt_h` | fly/wipe 的 `direction` left｜right｜top｜bottom（缺省 bottom / left） |
| exit | appear / fade / fly / wipe | 反向行为 ＋ 结束时 `set visibility=hidden` | 退出前必须没有被“退出”过；先退出后进入 = `E-MOTION-ORDER` |
| emphasis | grow_shrink（`scale` %、`autoreverse`）/ spin（`degrees`）/ color（`color`） | `animScale` / `animRot` / `set fill.*` ＋ `animClr` | 可 `repeat` 1–10；不提供 pulse 预设名 |
| path | motion_path | `animMotion origin=layout pathEditMode=relative` | `path.kind` line（`dx` `dy`）｜polyline｜curve（`points`，英寸偏移，曲线为经过各点的三次贝塞尔） |

通用：`trigger` on_click｜with_previous｜after_previous（缺省 on_click）、`duration_ms`、`delay_ms`、`repeat`、`group`（连续同组，仅作可读性标注）。**分步构建** `build`：`paragraph`（文本，每段一个效果）、`series`／`category`（原生图表；图表仍是原生可编辑对象，`bldChart animBg=0`）；展开后第 1 步用 `trigger`，其余用 `step_trigger`（缺省 on_click；改 `after_previous` 可把逐段变自动播放）。构建只支持 entrance；表格、legend 不支持。

## 三、编译与身份（不得绕开）

`spec.py` 统一解析（`motion.normalize`，把元素 id 解析为成品对象名）→ `post.py` 在 group 落成、`cNvPr@id` 确定之后调用 `motion.compile_slide`（**唯一一处生成时间树**）。目标身份 = `cNvPr@name`→实际 `cNvPr@id`，不依赖渲染顺序或 shapeIndex。时间树：click 组（`delay=indefinite`）→ 并行段（after_previous 分段）→ 效果（with_previous 同段）；cTn id 在树内顺序唯一；首个动画非点击时写 `onBegin` 自动开始；`p:sld` 子节点顺序 `cSld, clrMapOvr, transition, timing, extLst`。

## 四、既有稿（`package.py`，不经 python-pptx 重存）

`package.py inspect DECK [--text]` 读出每页对象 id/名称、动画（动画窗格顺序）、切换、媒体、备注。`package.py motion DECK OUT --slide N --set motion.json [--replace]`：`motion.json` 为 `{"transition":…, "animations":[…]}`，`target` 可写对象名或 `#id`；页上已有动画时**不加 `--replace` 一律阻断**；含交互序列（`interactiveSeq`）或媒体命令（`p:cmd`）的页**一律 BLOCK**，不静默丢失。`remove-shape` 删对象时同步移除其动画与 bld 登记、顺延 click 锚点；`dup` / `delete` / `move` 保持动画有效（对象 id 随页内复制不变；含分节的稿子 `move` 阻断）。详见 `ops.md`。

## 五、能力级别（PASS / BLOCK / NOT VERIFIED）

- **PASS（结构层，QA 已执行）**：上述切换与效果的 XSD 合法、目标 id 解析、触发语义、plan↔成品对账、既有稿定点写入与清除、复制/重排/删除后引用完整。
- **BLOCK（未实现，契约拒绝）**：切换音效、按词/按字构建、音视频原生插入与播放、SmartArt 构建动画。不用 GIF、视频或帧序列冒充逐元素动画。
- **NOT VERIFIED（恒定）**：Microsoft PowerPoint 实机——动画是否按预期播放、动画窗格是否显示预期效果名、是否出现“修复”提示、切换 `direction` 与界面“效果选项”文案的对应；Keynote / WPS / LibreOffice 各自独立，均未验证。`presetID` 只影响动画窗格的效果名，行为由 `cBhvr` 子节点决定。LibreOffice 渲染不能证明 PowerPoint 动画。
