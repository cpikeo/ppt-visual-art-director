# 命令、退出码、诊断码与能力边界

读诊断码、退出码、报告字段、能力边界、输入安全时读本文；plan 字段与元素语义见 `build.md`。

## 命令与退出码

- 输出只能是 `.pptx`（`E-FORMAT` 阻断）；成品包不合格（内容类型、宏）退出 3。
- `out`、`--report`、`--json` 已存在即阻断（`E-EXIST`），`--force` 才替换；先写临时件再原子替换。路径守卫只有一处（`model.guard_paths`）：输出/报告不得与任何输入同文件（符号链接与硬链接按真实文件比），不得经符号链接写出，报告只能是 `.json`，渲染目录不得是输入所在目录。`revise.py inspect` 解析前先做包检查（宏、越界部件名、解压体积）。
- `--assets` 缺省取 plan 同级 `assets/`，否则取 plan 所在目录。`verify.py` 的渲染目录缺省为 `render/<pptx 名>/`。
- 退出码：`compile.py` 0 成功、2 硬错误未写文件、3 运行期失败（`.error.pptx` 不是交付物；它与交付物、plan、报告走同一套 `guard_paths`：符号链接、目录、已存在且无 `--force` 一律不写，写入同样先临时文件再原子替换）；`verify.py` 与状态一一对应：`READY` 0、`BLOCK` 1（含包格式不合格）、`PARTIAL` 4、输入或契约错误 2（映射只在 `verify.EXIT_BY_STATUS` 一处，报告里的 `exit_code` 同值）；`revise.py` 0 写成功且保留核对通过、1 保留核对未通过（交付路径未写）、2 硬错误、3 运行期失败。

## 诊断码

行格式 `[HARD] E-CHART p2/bar: …`。HARD 用 `E-` 必须改；SOFT 用 `W-` 须逐条判断并写理由；INFO 用 `I-` 只提示看哪一页。`V-` 来自 verify.py。

| 码 | 含义 | 处置 |
| --- | --- | --- |
| `E-OUTLINE` | `scaffold.py init` 的 outline 缺字段、recipe 不存在、steps 项数不对 | 按提示逐条补，未写出文件 |
| `E-ASSET` `W-ASSET` | 素材来源、权利、生成记录、简报不一致；生成图作证据 | 补齐记录，或改用真实素材 |
| `E-LAYER` `E-REF`（depth） | `depth` 声明的前/中/背景与叠放（声明）顺序矛盾，或引用了不存在的元素 | 把背景元素写在前景之前，或改声明 |
| `W-DOMINANT` | `focus_element` 的粗估视觉重量（文字＝字号²×字数，图/图表/色块＝面积；承载层不计）没有明显超过同页配角 | 放大主角、缩小或灰化配角，或删掉 focus_element。只比尺度，不判断好不好看 |
| `E-OVERLAP` | 实质重叠，下层未声明承载 | 多半忘了 `overlap_allowed` |
| `E-CHART` `W-CHART` `W-NOTE` | 图表口径、单位、截轴说明、颜色映射（`color_map` 指派对不上类别、饼/环同色、手绘图例色块）、系列色沉进承载面、图例与强调柱对不上 | 改措辞、拆图、补颜色映射；错色没有放行开关 |
| `V-LABEL-INK` `V-LABEL-POS` | 成品图表里：直标/轴/图例墨色落在其真实表面上 < 4.5:1（`chart_audit.py` 只读 XML）；或 `dLblPos` 对该图表类型非法；或墨色与决策表不一致 | 重新编译；手改成品会再次被拦 |
| `V-MARKER` | 成品图表里：散点 marker 不可见（symbol=none 或无填充）、小于 5pt、没有关闭连线；折线 marker 色 ≠ 线色。散点 marker 色与调色板对不上记 `V-CHART`；marker 对承载面的对比、灰度区分由 `chart_audit` 经 marker 填充计算（`V-CONTRAST` `V-GRAY`） | 重新编译；手改成品会再次被拦 |
| `W-GRAY` | **任意两个**系列的颜色转灰度后明度差 ΔL* 低于 `chart.min_series_delta_lstar`（成品里同一判断为 `V-GRAY`） | 拉开明度，或直标系列名、换线型与标记 |
| `E-CHART`（读数） | 饼/环没有任何类别来源：饼需直标、原生图例或 `legend` 元素之一；环内直标只有数值，需原生图例或 `legend` 元素 | 开 `legend`／加 `legend` 元素；纯图例无数值直标给 `W-CHART` |
| `V-READOUT` | 成品饼/环里读不到类别（HARD）或数值（SOFT）：只读 XML 里的 `c:legend`、`showCatName`、`showVal` | 与 plan 无关，成品被改过或编译链有误 |
| `E-FORMAT` | 不是合格的无宏 `.pptx` | 输出改为 `.pptx`；原稿含宏则人工另存 |
| `E-PRESERVE` | 定点编辑保留核对不通过 | 看 `preservation` 字段定位越界 |
| `W-FIT` | 装载估算溢出（行数按字宽折算；`%` `×` `@` 按近全角宽度计；高度按“首行可见高度＋后续行距推进”估，按渲染实测校准） | 先改文案，再加框，不压字号 |
| `W-SPACE` | 中文与数字/西文之间手敲了**多于一个**的空格、不断行空格或全角空格 | 删掉，依赖字体原生字距 |
| `I-SPACE` | （整稿汇总为一条，列出位置）中西文之间只有**一个**普通空格：通行的中文排法，LibreOffice 会叠加自动间距显松、PowerPoint 里正常 | 是风格取舍不是缺陷；全稿统一取一种并写进交付说明 |
| `V-BLEED` | 文字落在声明为出血的元素里，被页边裁掉 | 确认被裁部分不含信息；没声明出血的文字出页仍是 `V-CLIP` 阻断 |
| `W-CONTRAST` `V-CONTRAST` | 对比不足；文字压图时是对真实像素的量测（达标像素占比低于 `limits.json` 的 `image.zone.min_ok_share`） | 换墨色或落点、改 `crop_focus`、重做简报；不靠浓遮罩 |
| `W-BUSY` | 图上文字对比够，但下方纹理过强（灰度标准差超限） | 换落字区或重做简报 |
| `W-PPI` | 有效分辨率低 | 换素材或缩小显示 |
| `W-COVER` | 文字落在上层色块或图上 | 挪位置或调层级 |
| `W-RUNS` | 原有混排被统一 | 用 `runs` 写明分区样式 |
| `V-TEXT-RENDER` | 元素的字未落在自己框内渲染 | 查归属，不因别处有同样的字放行 |
| `V-FONT` | 声明字体未参与渲染 | 状态降为 PARTIAL，目标机复核 |
| `W-FONT-FACE` `I-FONT-FACE` | 家族匹配上了，**地区字面**没匹配（声明简体而 PDF 取到日文字面），或页面含中文但字面无法判定 | **状态降为 PARTIAL**（`font_face` = `MISMATCH` / `UNVERIFIED`）；直/骨/门/者一类共享码位的字形按别的地区字面渲染，简体字形要人在 PowerPoint 里复核 |
| `V-NOTES` `V-ALT` | 成品里的备注／替代文字与 plan 不一致 | 重新编译；不要手改成品 |

## 报告字段

- `compile.py`：`blocked`、`output`、`pages`、`sha256`、`issues`。
- `verify.py`：`status`、`exit_code`、`sha256`、`render`（PDF、PNG、`font_substituted`、`text_checks`）、`structure`、`layers`（`objects`／`render`／`fonts`／`human`）、`findings`。`render` 非 `PASS` 时 `status` 只能是 `PARTIAL`；`fonts` 为 `SUBSTITUTED`、`FACE MISMATCH`、`FACE UNVERIFIED` 时不能是 `READY`（`render.font_face`：`VERIFIED` `MISMATCH` `UNVERIFIED` `NOT_APPLICABLE`）；`layers.powerpoint` 恒为 `NOT VERIFIED`。判定只在 `verify.decide_status` 一处。
- `revise.py`：`status`、`output`、`sha256`（失败为 null）、`ops`、`preservation`（`parts.changed` 为越界，`parts.moved` 为声明过的改写；`rejected` 为证据路径）、`issues`。

`sha256` 用于确认“这份文件是不是我验的那一份”，不是复现凭据。

## 能力边界

- 判断记录字段（`brief.*`、`audience`、`goal`、`core_claim`、页 `focus` `composition_rationale` `open_questions`、`brief.motif`、`brief.art.*`）由 schema 校验必填与形状，脚本不对其内容做语义判断——它们是写给人与下一轮 AI 的决策记录，不是机器可证明的字段。

| 能力 | 级别 | 说明 |
| --- | --- | --- |
| 画布 16:9、4:3 | 完整 | 其他比例需同时改 `limits.json` 与 schema，未测 |
| 文本、形状、折线路径、组、图片裁切 | 完整 | 对象核对覆盖几何、旋转、逐 run 样式、字节 |
| 原生图表：柱、条、堆叠、百分比堆叠、组合、线、面积、饼、环、散点；原生表格 | 完整 | 无双数值轴、瀑布、桑基、甘特、气泡、雷达、趋势线 |
| 按数据点着色 | 部分 | 仅 `highlight_category`（单系列柱/条，其余退隐灰，不能开原生图例）与饼/环按类别取色；折线、面积、散点不支持，强调单条靠直标或旁注 |
| 颜色映射与图例一致性 | 完整 | 类别/系列→颜色唯一映射（`color_map`＞图级 `palette`＞整稿 `palette`）；`legend` 元素按同一映射生成并逐块核验 |
| 演讲者备注、替代文字、页码、页标题 | 完整 | 页标题是原生标题占位符；阅读顺序仍按声明顺序，未做逐页阅读顺序审查 |
| 既有稿定点编辑 | 完整（限白名单操作） | 不新增页、不改页序、不改系列数、不动母版与版式 |
| 渐变、`duotone` | 有限 | PowerPoint 生效；LibreOffice 只做去饱和或近似 |
| 路径连线 | 有限 | 直线与折线，锚点是编译期坐标，不是连接符绑定 |
| 图片裁切 | 有限 | 按 `cover`，不保证主体不被裁到 |
| 图片生成 | 外部依赖 | 工具无关：只登记工具、完整提示词、文件、SHA-256、权利状态；本包不调用任何模型，生成图不作证据 |
| 渲染核验 | 外部依赖 | LibreOffice + PyMuPDF（`requirements-render.txt`，可选）；缺失给 `PARTIAL`。Noto CJK 的 `.ttc` 在 PDF 里只能按家族匹配，分不出 SC 与 JP 字面 |
| PowerPoint 实机验证 | 外部依赖 | 本包从未做过；结论由人在真实 PowerPoint 中签收，未做即 `NOT VERIFIED` |
| 字体 | 外部依赖 | 只写字体名，不嵌入；目标机缺失会被静默替换。主题字体（拉丁与东亚）取 `fonts`，新建 run 带 `lang`（含中文为 zh-CN） |
| 母版与版式 | 有限 | 默认模板的 4:3 占位符几何按画布宽度横向缩放；**主题色（dk2/lt2/accent1–6、超链接色）取 plan 的数据色与墨色，母版底色取多数页底色，母版标题左对齐、与默认标题同字号**，新增页与图形不会回到 Office 默认外观；`docProps/app.xml` 按成品重写（页数、画布、不声称由 Office 生成）。仍只有一套母版，不做自定义版式库；未使用的 11 个默认版式保留 |
| 动画、切换、嵌入字体、新建 SmartArt、宏 | 不支持 | `.pptm` 一律拒绝 |

## 输入安全

plan、edit.json、素材与 PPTX 都视为不可信输入：素材路径限制在素材根内；解析外部 PPTX 的 XML 一律不展开实体、不联网（防 XXE 与实体膨胀）；成品包拒绝越界部件名与解压后超过 `package.max_unpacked_mb` 的包。脚本不联网、不执行包内内容。
