# 命令、退出码、诊断码与能力边界

读诊断码、退出码、报告字段、能力边界、输入安全时读本文；plan 字段与元素语义见 `build.md` `elements.md`。

## 命令与退出码

- 输出只能是 `.pptx`（`E-FORMAT` 阻断）；成品包不合格（内容类型、宏）退出 3。
- `out`、`--report`、`--json` 已存在即阻断（`E-EXIST`），`--force` 才替换；先写临时件再原子替换。路径守卫只有一处（`model.guard_paths`）：输出/报告不得与任何输入同文件（符号链接与硬链接按真实文件比），不得经符号链接写出，报告只能是 `.json`，渲染目录不得是输入所在目录。`revise.py inspect` 解析前先做包检查（宏、越界部件名、解压体积、压缩比）。
- `--assets` 缺省取 plan 同级 `assets/`，否则取 plan 所在目录。`verify.py` 的渲染目录缺省为 `render/<pptx 名>/`。
- 退出码：`compile.py` 0 成功、2 硬错误未写文件、3 运行期失败（`.error.pptx` 不是交付物；它与交付物、plan、报告走同一套 `guard_paths`：符号链接、目录、已存在且无 `--force` 一律不写，写入同样先临时文件再原子替换）；`officecheck.py` 0 PASS｜1 BLOCK｜2 输入错误；`package.py` 0 成功｜1 校验或操作被阻断（交付路径未写）｜2 路径/契约错误；`verify.py` 与状态一一对应：`READY` 0、`BLOCK` 1（含包格式不合格）、`PARTIAL` 4、输入或契约错误 2（映射只在 `verify.EXIT_BY_STATUS` 一处，报告里的 `exit_code` 同值）；`revise.py` 0 写成功且保留核对通过、1 保留核对未通过（交付路径未写）、2 硬错误、3 运行期失败。

## 诊断码

行格式 `[HARD] E-CHART p2/bar: …`。HARD（`E-`）必须改；SOFT（`W-`）须逐条判断并写理由；INFO（`I-`）只提示看哪一页；`V-` 来自 verify.py。每条输出自带含义与处置，下表只补**易误判**的几组：

| 码 | 含义与处置 |
| --- | --- |
| `E-OUTLINE` | outline 缺字段、recipe 不存在、steps 项数不对；按提示补，未写文件 |
| `E-ASSET` `W-ASSET` `E-IMG` | 来源、权利、生成记录、简报不一致，或生成图作证据；补记录或改真实素材 |
| `E-LAYER` `E-REF` | `depth` 与叠放顺序矛盾，或引用不存在的元素；背景元素写在前景之前 |
| `E-OVERLAP` | 实质重叠且下层未声明承载；多半忘了 `overlap_allowed` |
| `W-DOMINANT` | `focus_element` 的粗估视觉重量（文字＝字号²×字数，图/图表/色块＝面积）没有明显超过配角；只比尺度，不判断好看 |
| `E-CHART` `W-CHART` `W-NOTE` `W-SUM` | 口径、单位、截轴说明、颜色映射、系列色沉进承载面、饼/环合计非 100（为 0 即阻断）、饼/环无类别来源；改措辞、拆图、补映射，错色没有放行开关 |
| `W-GRAY` `V-GRAY` | 任意两系列灰度 ΔL* 低于 `chart.min_series_delta_lstar`（limits.json，现为 9.5，scaffold 色板名义 10 留损耗）；拉开明度，或直标、换线型与标记 |
| `V-LABEL-INK` `V-LABEL-POS` `V-MARKER` `V-READOUT` | 成品图表被改过或编译链有误（墨色对比 < 4.5、`dLblPos` 非法、散点 marker 不可见、饼/环读不到类别）；重新编译，不手改成品 |
| `E-FORMAT` | 不是合格的无宏 `.pptx`（含 DTD/实体、越界部件名、解压体积或单部件压缩比异常）；输出改 `.pptx`，宏稿人工另存 |
| `E-SCHEMA` | schema 不符，或 JSON 本身无效：`NaN`/`Infinity`、对象内重复键（标准解析器会静默取最后一个）、嵌套过深；plan、edit.json、outline 同一规则（`model.strict_json_loads`） |
| `E-RUNTIME` | 校验或编译期未预料的异常；`compile.py` 退出 3，不写交付件、不抛 Traceback |
| `E-PRESERVE` | 定点编辑保留核对不通过；看 `preservation` 定位越界 |
| `W-FIT` | 装载估算溢出；先改文案，再加框，不压字号 |
| `W-PUNCT` `W-SPACE` `I-SPACE` | 汉字旁半角标点；中西文间多于一个空格（错）；仅一个空格（通行排法，I 级，全稿统一并写进交付说明，LibreOffice 会显松） |
| `V-BLEED` `V-CLIP` | 文字落在出血元素里被裁（确认不含信息）；没声明出血的文字出页即阻断 |
| `W-FUSION` `I-FUSION` | 同页多图明度/冷暖/色度差超限（只量可计算部分，不评美）；换图或改裁切，有意并置则写 `composition.exception` |
| `W-CONTRAST` `V-CONTRAST` `W-BUSY` `W-PPI` | 对比不足（文字压图为真实像素量测）、纹理过强、分辨率低；换墨色或落点、改 `crop_focus`、重做简报，不靠浓遮罩 |
| `W-COVER` `W-RUNS` | 文字落在上层色块或图上；原有混排被统一，用 `runs` 写明 |
| `V-TEXT-RENDER` `V-FONT` | 元素的字未落在自己框内渲染；声明字体未参与渲染（降为 PARTIAL） |
| `W-FONT-FACE` `I-FONT-FACE` | 家族匹配、地区字面没匹配（如简体取到日文字面）或无法判定；降为 PARTIAL，直/骨/门/者等共码位字形须人在 PowerPoint 复核 |
| `V-ORPHAN` | 文本元素折行后末行只剩一个字（`limits.json` → `render_check.orphan_min_chars`，现为 4：短说明里的孤字最显眼）；图表轴与表格单元格不判。改文案或框宽；配方内的步骤、对照、大数字正文已在编译期做估算防护（`Ctx.widow_w`），渲染兜底仍是这条 |
| `V-SPACE` `V-RHYTHM` | 渲染量测（`scripts/spatial.py`，阈值 `limits.json` → `space`）：声明的 `temperament` 与实际占用/空白不符、未声明功能的整页过疏或大块空白、`negative_space.role=focus` 的空白离主角过远、连续 3 页版面分区几乎相同。SOFT，只指出位置；接受就写理由，不靠调阈值 |
| `V-NOTES` `V-ALT` | 成品备注/替代文字与 plan 不一致；重新编译 |
| `E-GEOM` `E-ID` `E-PAGE` `E-SHAPE` `E-TEXT` `E-SRC` | 结构类硬错误：元素缺坐标或 group 无子元素；页 id 重复；文字/图表压到页码保留框；shape 无填充无描边；文本为空；改稿找不到原稿。按提示改 plan 或路径 |
| `W-PT` `W-BOUND` `W-PAGE` `W-RHYTHM` `W-ANCHOR` `W-REF` `W-SHAPE` | 字号低于 `type.min_pt`；改动后出画布；`page_numbers.hide` 里不是页 id；`rhythm` 条数≠页数；锚点相关；锚点目标宽或高为 0；闭合路径写了箭头（忽略） |
| `I-ART` `I-GRID` `I-GROUP` `I-PAGE` `I-LABEL` `I-AXIS` | 提示级：有图却没写 `brief.art`；`grid=none` 没写破格理由；group 几何由子元素推出；图片压到页码角区；直标色非显式 RGB 未核对对比；读不到数值轴显示状态 |
| `E-MOTION-TARGET` `E-MOTION-ID` `E-MOTION-FIELD` `E-MOTION-ORDER` `E-MOTION-BUILD` `E-MOTION-PATH` `E-MOTION-GROUP` `E-MOTION-ADVANCE` | 动画/切换契约：目标不是本页顶层元素（或是 group 子元素）、动画 id 重复、字段与 effect 不匹配、先退出后进入、构建对象不合法、路径缺点或偏移过大、group 不连续、禁用点击又无自动前进；按提示改 plan |
| `W-MOTION-CLICKS` `W-MOTION-DENSITY` `W-MOTION-LENGTH` `W-MOTION-SLOW` `W-MOTION-VARIETY` `W-MOTION-BLANKET` | 动画克制度（SOFT）：单击阶段 > 3、一页动画 > 8、累计 > 6 s、单个 > 1.5 s、整稿切换种类 > 2、每页都有动画；写理由或删减，不靠调阈值 |
| `E-MOTION`（编译期）`E-MOTION-STRANDED` `E-MOTION-CTN` `E-MOTION-CLICK` `E-MOTION-RANGE` | 成品时间树不合格：段落/系列数与目标不符、悬空目标 id、cTn id 重复、click 锚点丢失、构建索引越界；`compile.py` 退出 3 或 `officecheck.py` BLOCK |
| `E-OPC` `E-REL` `E-XML` `E-XSD` `E-ID` `E-SHARED` `E-PRES` `E-CHART` | `officecheck.py` 结构层：Content Type/关系缺失、XML 不可解析、XSD 不合法、同页 id 重复、多页共享图表/工作簿/备注（复制页未深拷贝）、sldId 不合法 |
| `V-MOTION` `V-OFFICE` | verify 的动画核对：成品与计划的切换/动画序列/触发不一致；结构层问题 |
| `V-OBJ` `V-ORDER` `V-PAGES` `V-PAGENO` `V-ASSET` `V-ROT` `V-RENDER` | verify 的对象层与环境：成品缺对象；叠放顺序≠声明；页数≠plan；声明隐藏的页码仍在；找不到素材原件无法比对内嵌字节；旋转未生效；**`V-RENDER`＝没有 LibreOffice，渲染核对未做（状态 PARTIAL，不是通过）** |

## 报告字段

- `compile.py`：`blocked`、`output`、`pages`、`sha256`、`issues`。
- `verify.py`：`status`、`exit_code`、`sha256`、`render`（PDF、PNG、`font_substituted`、`text_checks`）、`structure`、`layers`（`objects`／`render`／`fonts`／`human`）、`findings`。`render` 非 `PASS` 时 `status` 只能是 `PARTIAL`；`fonts` 为 `SUBSTITUTED`、`FACE MISMATCH`、`FACE UNVERIFIED` 时不能是 `READY`（`render.font_face`：`VERIFIED` `MISMATCH` `UNVERIFIED` `NOT_APPLICABLE`）；`layers.powerpoint` 恒为 `NOT VERIFIED`。判定只在 `verify.decide_status` 一处。
- `revise.py`：`status`、`output`、`sha256`（失败为 null）、`ops`、`preservation`（`parts.changed` 为越界，`parts.moved` 为声明过的改写；`rejected` 为证据路径）、`issues`。

`sha256` 用于确认“这份文件是不是我验的那一份”，不是复现凭据。

## 能力边界

- 判断记录字段（`brief.*`、`audience` `goal` `core_claim`、页 `focus` `composition_rationale` `open_questions`）只校验必填与形状，脚本不做语义判断：它们是写给人与下一轮 AI 的决策记录。

| 能力 | 级别 | 说明 |
| --- | --- | --- |
| 画布 16:9（1280×720 px） | 完整 | 唯一画布。其他比例（含 4:3）不支持：`revise.py` 对非 16:9 稿 `E-CANVAS` 拒收，`verify.py` 页面尺寸不等于 12192000×6858000 EMU 报 `V-CANVAS` |
| 文本、形状、折线路径、组、图片裁切 | 完整 | 对象核对覆盖几何、旋转、逐 run 样式、字节 |
| 原生图表：柱、条、堆叠、百分比堆叠、组合、线、面积、饼、环、散点；原生表格 | 完整 | 无双数值轴、瀑布、桑基、甘特、气泡、雷达、趋势线 |
| 按数据点着色 | 部分 | 仅 `highlight_category`（单系列柱/条，其余退隐灰，不能开原生图例）与饼/环按类别取色；折线、面积、散点不支持，强调单条靠直标或旁注 |
| 颜色映射与图例一致性 | 完整 | 类别/系列→颜色唯一映射（`color_map`＞图级 `palette`＞整稿 `palette`）；`legend` 元素按同一映射生成并逐块核验 |
| 演讲者备注、替代文字、页码、页标题 | 完整 | 页标题是原生标题占位符；阅读顺序仍按声明顺序，未做逐页阅读顺序审查 |
| 既有稿定点编辑（`revise.py`） | 完整（限白名单操作） | 不新增页、不改页序、不改系列数、不动母版与版式；页的增删/复制/重排与动画编辑见 `package.py` |
| 渐变、`duotone` | 有限 | PowerPoint 生效；LibreOffice 只做去饱和或近似 |
| 路径连线 | 有限 | 直线与折线，锚点是编译期坐标，不是连接符绑定 |
| 图片裁切 | 有限 | 按 `cover`，不保证主体不被裁到 |
| 图片生成 | 外部依赖 | 工具无关：只登记工具、完整提示词、文件、SHA-256、权利状态；本包不调用任何模型，生成图不作证据 |
| 渲染核验 | 外部依赖 | LibreOffice + PyMuPDF（`requirements-render.txt`，可选）；缺失给 `PARTIAL`。Noto CJK 的 `.ttc` 在 PDF 里只能按家族匹配，分不出 SC 与 JP 字面 |
| PowerPoint 实机验证 | 外部依赖 | 本包从未做过；结论由人在真实 PowerPoint 中签收，未做即 `NOT VERIFIED` |
| 字体 | 外部依赖 | 只写字体名，不嵌入；目标机缺失会被静默替换。主题字体（拉丁与东亚）取 `fonts`，新建 run 带 `lang`（含中文为 zh-CN） |
| 母版与版式 | 有限 | python-pptx 模板遗留的 10 in 母版占位符几何按 16:9 画布宽度横向缩放；**主题色（dk2/lt2/accent1–6、超链接色）取 plan 的数据色与墨色，母版底色取多数页底色，母版标题左对齐、与默认标题同字号**，新增页与图形不会回到 Office 默认外观；`docProps/app.xml` 按成品重写（页数、画布、不声称由 Office 生成）。仍只有一套母版，不做自定义版式库；未使用的 11 个默认版式保留 |
| 页面切换（20 种标准类型）、对象动画（进入/退出/强调/路径）、On Click / With / After Previous、文本逐段、图表逐系列/类别 | 完整（结构层）| 编译与既有稿写入均走 `motion.py`；详见 `motion.md`。PowerPoint 实机播放 `NOT VERIFIED` |
| 既有稿结构操作（复制/删除/重排页、删对象、清理孤立部件、模板 POTX 分析、动画定点写入） | 完整（限 `package.py` 子命令）| 不经 python-pptx 重存；输出前过 `officecheck.py`（含原稿基线）；含分节的稿子不移动页；含 `interactiveSeq`/`p:cmd` 的页不改写动画 |
| 切换音效、按词/按字构建、音视频原生插入与播放、SmartArt 构建动画、嵌入字体、新建 SmartArt、宏 | 不支持 | 契约拒绝（`E-SCHEMA`）或明确 BLOCK；`.pptm` 一律拒绝 |

## 输入安全

plan、edit.json、素材与 PPTX 都视为不可信输入：素材路径限制在素材根内；外部 PPTX 的 XML 不展开实体、不联网，含 DTD/实体声明（含 UTF-16 编码）一律 `E-FORMAT`；包内越界部件名（含反斜杠写法）、解压后超过 `package.max_unpacked_mb` 的包、单个大部件压缩比超过 `package.max_part_ratio` 的包。已知限制：ZIP 头里的声明体积可被伪造，守卫以头部为准，不是流式计量。脚本不联网、不执行包内内容。

## 生成引擎能力矩阵（V22，PptxGenJS 4.0.1）

DIRECT = PptxGenJS 直接生成；OOXML = 由 `post.py` 按 spec 受控补齐；BLOCK = 不能可靠实现，`compile.py` 退出码 2 并报 `E-ENGINE`，不静默降级。

| 能力 | 类别 | 说明 |
| --- | --- | --- |
| 标题占位符、文本框、runs、段落、行距、字距、内边距、垂直对齐 | DIRECT + OOXML | 占位符几何/锚点、`a:ea`/`a:cs`、`lang`（zh-CN/en-US）、`noAutofit`、重复 `a:pPr` 由 post 校正 |
| 原生列表（符号/破折号/自动编号） | DIRECT + OOXML | 符号色 `a:buClr` 由 post 写入 |
| shape（8 种）、divider、填充透明度、描边、虚线、旋转 | DIRECT | |
| 渐变（含两端透明度）、圆角半径 `adj`、线帽、箭头 | OOXML | |
| path：多点折线/闭合多边形 | DIRECT（`custGeom`） | 点可引用其他元素的锚点（由 spec.py 解析） |
| path：两点 | OOXML | 改写为原生直线连接符 `p:cxnSp` |
| image：放置、透明度、旋转、替代文字 | DIRECT | |
| image：`crop_focus`/`crop_zoom`、预设与多边形蒙版、双色调、软边 | OOXML | `srcRect`/`prstGeom`/`custGeom`/`duotone`/`softEdge` |
| group、legend（色块+标签分组） | OOXML | PptxGenJS 无分组；`p:grpSp` 由 post 包裹，支持嵌套与分组旋转 |
| chart：bar/line/area/pie/donut/scatter，堆叠/百分比，横向，组合图（折线与柱共轴），嵌入工作簿 | DIRECT | |
| chart：逐系列直标（墨色/衬底/位置）、东亚字体、标题左对齐、饼图内缩、散点仅标记、组合折线标记、系列色 | OOXML | |
| chart：**散点各系列 X 值不一致** | **BLOCK** | PptxGenJS 散点共用一列 X；改统一 X |
| table | DIRECT + OOXML | 无样式表 ID 由 post 写入 |
| speaker_notes、页码、背景色、主题色/主题字体 | DIRECT / OOXML | 主题色板与东亚字体由 post 写入 theme1.xml |

`post.py` 同时修正 PptxGenJS 的已知结构问题：图表 `c:ser` 子元素次序（`dLbls` 误排在 `dPt` 之前）、2D 绘图层多出的 `axId`、非柱图混入 `invertIfNegative`、`[Content_Types].xml` 指向不存在的母版、重复媒体。图表 XML 自检（轴引用、series 序号、次序、`dLblPos`、数值有效性、线型）不通过即 `E-FORMAT` 阻断。

**验证状态矩阵（不打分）：** 对象核对 PASS（`verify.py`，本仓库夹具）｜XML 不变量 PASS｜可复现（同输入同字节）PASS｜LibreOffice 渲染 UNVERIFIED（运行环境无 soffice）｜Microsoft PowerPoint 兼容 UNVERIFIED（从未做过实机验证）。

