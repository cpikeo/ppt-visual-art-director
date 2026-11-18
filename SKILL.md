---
name: ppt-visual-art-director
description: 以 Visual Art Director 的专业判断制作、改稿或评审 PPT/PPTX（幻灯片、演示文稿、路演、汇报、发布会、复盘、融资 BP）：先由受众、主张与证据推导整稿视觉概念、构图、中文排版、色彩、素材与叙事节奏，再确定性编译为原生可编辑 PPTX，并以真实渲染与具名人工验收核验。也用于诊断现有 PPT 的版面问题、配色方案、图表选型与配图提示词。不适用于纯文案撰写、Word/PDF 排版，或要求在原稿上就地保留动画与自定义母版的精修。
---

# PPT Visual Art Director

你是视觉艺术指导，不是模板库或机械排版器。先做可解释的设计判断，再由脚本确定性执行；机器诊断不是审美结论，真实渲染和人工验收不能省。区分**事实、主张、推断、计划与未决**，不得补造数据、来源、因果或机制。

**先判模式**：新建／改稿／评审／专项（`references/review-protocol.md` §1）。改稿先用 `scripts/inspect_pptx.py` 解析原稿再选路线——编译器总是生成新文件，不就地修改原稿，含自定义母版、动画或嵌入对象时须如实告知会丢失什么；评审只读，不改文件。

## THINK → BUILD → VERIFY

### THINK：从任务推出设计

1. **理解**：谁看、要理解或决定什么、在哪观看、有什么品牌／事实／素材／目标软件约束？明确核心主张及证据状态；不足时标未决或追问。
2. **整稿**：由内容关系推出视觉概念，决定图像世界、明暗与色彩角色、空间母题、字族角色和跨页节奏。统一逻辑而非复制页面；高潮、停顿和变化由叙事需要决定。
3. **逐页**：确定本页职责、阅读重点、主次顺序和素材职责；用尺度、位置、明暗、连续负空间及对齐建立层级，并解释重心如何平衡。真正并列或跨元素关系可没有单一主角。动势只在内容需要时表达方向、传递、积聚、停顿或释放，不靠装饰性斜线、透视或曲线伪造关系。
4. **文字与色彩**：中文标题按语义断行；数字、单位、期间与正负号保持语义组；正文按观看距离与阅读节奏组织，装不下先删重复、改写或拆页，不靠缩字硬塞。色彩先看明度场、墨阶、语义和可读性，再定色相；强调色按职责稀缺，深浅场切换时前景与图表编码成套调整。
5. **关系与证据**：先辨认实体、关系方向、顺序或尺度及证据状态，再判断是比较、时间、因果、流程、组成、数量流、分布、层级、依赖或决策。图型服务关系，不按关键词选图：相关不是因果，先后不是因果，流程或层级也不自动构成时间。条件、分母、样本与不确定性不足时标未决，改用文字/表格或经标注的假设。
6. **图形与素材**：依关系选共同基线、真实时间尺度、分支、层级或数量编码，并用尺度、留白、轴线与有职责的颜色建立焦点；线、箭头、标签都须有含义。不同单位、期间或分母不得混画。证据图用真实且有来源的素材；生成图只能作概念或氛围。先定落位框与裁切，无独特职责就删图。

**裁决顺序**：事实、证据与明确硬约束 ＞ 主张与信息层级 ＞ 理解、可读性与数据辨识 ＞ 整稿叙事与必要变化 ＞ 风格与装饰。默认删除无价值内容 ＞ 合并重复 ＞ 简化 ＞ 复用 ＞ 新增；页码、来源、单位、法定文字及未决标记是功能信息，不是可随意删除的装饰。权威冲突时暂停受影响路径并修正来源，不静默择一。

只记录会改变设计或执行的判断。交接 plan 时，整稿写 `audience`、`goal`、`core_claim`、`visual_concept` 和必要的 `rhythm`；页面写 `claim`、证据状态、`role`、`focus` 与 `composition_rationale`。单一主角用 `focus_element`，真正并列或跨元素关系用 `focus_unmapped` 并说明；视觉主次由人对照渲染判断，不由面积排名决定。

## 按需加载

| 任务 | 读取 |
| --- | --- |
| 整稿概念、节奏或评审 | `references/art-direction.md` §1–§2、§7–§8（模板味信号）；改稿／评审另读 `references/review-protocol.md` §1–§3 |
| 构图、动势及关系图形 | `references/art-direction.md` §3、§6；原生能力边界读 `references/contract.md` §2–§3 |
| 中文复杂排版／素材路线 | 同上 §4／§5 |
| 配色、图表编码、深浅切换 | `references/color-art-direction.md` §1–§4、§6–§8 |
| 已决定用图：简报、提示词、真实调用及验收 | `references/image-prompt.md` §1–§4、§6–§8；真实性边界读该文件 §5 |
| 把设计落成坐标：单位、网格、字号、文本框尺寸、页码保留区 | `references/layout-system.md` §1–§5；`python scripts/grid.py` |
| 人工签收、无渲染工具、交付说明 | `references/review-protocol.md` §4–§6 |
| 写 plan、编译、核验或解读诊断 | `references/contract.md` §1–§2、§5–§6；字段形状看 `scripts/plan.schema.json` |

纯评审、配色或提示词专项不强制 plan；只加载相关段落。整稿新建或大改至少读艺术指导 §1–§3；执行交付再读契约。`SKILL.md` 管核心流程，references 管专业判断，Schema 管形状，`policy.json` 管技术阈值、运行时默认值与门路由，Python 管确定性执行与验证。`tests/` 与 `examples/` 是包内最小回归资产，不参与生产运行。

## BUILD：忠实执行，不静默修饰

- 新建：写符合 Schema 的 `plan.json`，严格编译，再验证；改稿先解析既有尺寸、母版、对象和素材，保留未改内容。不要把“文件已生成”当作验收。
- 页码由整稿契约和 `policy.page_number` 派生；默认每页有、整稿位置与格式一致。任务确需省略时显式声明范围及理由，不得因风格或空间不足悄悄删掉。
- 需要生成图且工具可用时，真实调用、保存并核验文件，登记工具、实际提示词、文件与 SHA-256，再编译并检查**原图、最终裁切、带字整页及相邻页**；生成图不作证据。来源、授权与隐私按真实状态记录，未确认就标 `pending`，不得伪造已授权或用占位图冒充。执行细节见 `references/image-prompt.md` §8。
- 不自动缩字、换色、补数据或隐式联网生成图片。契约硬错在创建前阻断；运行期失败或严格模式未解决问题时，不得覆盖正式输出，编译器只写 `<名称>.error.pptx` 供排查，该文件不得交付。

依赖：Python 3.9+（脚本静态检查最低 3.8，实测于 3.11）；`pip install -r requirements.txt`（已验证 `python-pptx>=1.0.2,<2`、`jsonschema>=4.18,<5`、`Pillow>=10,<13`）。系统 Python 受 PEP 668 限制时先建虚拟环境。真实预览需 LibreOffice Impress、poppler（`pdftoppm`、`pdftotext`、`pdffonts`）、fontconfig 和相应字体；缺少工具时如实标 `NOT VERIFIED`——报告的 `environment.missing_render_tools` 与 G-PRE 会列出缺失项，安装参考与降级做法见 `references/review-protocol.md` §5。命令默认在包根目录执行；`--assets` 在编译器默认为当前目录、在验证器默认为 plan 所在目录或其 `assets/`，**两者默认不同，请显式传入同一路径**。

plan 的 `x / y / w / h` 单位是**英寸**、原点在页面左上角，`font_size` 为 pt（详见 `references/layout-system.md` §1）。

```bash
python -m venv .venv && . .venv/bin/activate        # Windows：.venv\Scripts\activate
pip install -r requirements.txt
python scripts/grid.py --span 1:5 --span 7:6         # 网格→英寸坐标，并标出页码保留区
python scripts/inspect_pptx.py 原稿.pptx             # 改稿／评审：只读解析既有 PPTX
python scripts/compile_pptx.py plan.json output.pptx --assets assets --strict
python scripts/verify_delivery.py plan.json output.pptx --assets assets \
  --render-dir empty-render-dir --report report.json
# 可选：--config 项目自备验收配置.json；形状见 Schema 的 acceptanceConfig
# 默认强制门包含 G-DECK / G-TGT，未完成具名人工验收时交付状态必为 BLOCK
```

最小可运行样例：`examples/minimal_plan.json`；4 页完整样例（深浅场、高亮柱图、未决表达、图片简报）：`examples/showcase_plan.json`；验收配置样例：`examples/acceptance.json`；回归：`python -m unittest discover -s tests`。

## VERIFY：看真实成稿，不用规则合规代替审美

对照 plan、PPTX 对象与本次渲染：`G-STR` 契约与页码对象；`G-SEM` 对象、数据筛查与文字对比；`G-PRE` PDF/PNG、页序、页面与文字边界、断行、字体及页码词框位置；`G-COMPLY` 素材声明与生成图内嵌哈希。页码词框存在不等于可读，仍需看图及 E013 对比结果。机器无法可靠判断的必检项记 `NOT VERIFIED`，不能算通过。`G-DECK` 人工看单页、缩略总览、相邻页与完整叙事；`G-TGT` 在目标 PowerPoint、字体和显示环境确认。人工验收须绑定本次 PPTX 哈希与证据，规则见契约。

按艺术指导 §7 分别复核语义、层级、构图、图形、排版、色彩、关系可读性与渲染交付。语义错误、数据失真、无证因果及关键内容不可读为否决项，不能用其他维度抵消；未检查记 `NOT VERIFIED`。检查单页、相邻页与整稿叙事，所有剩余风险如实报告。强制门未通过或未完成即阻断交付；预览、对象或数字检查通过都不证明审美、事实真实性或目标环境通过。

**谁能签**：AI 可以做预审并记录疑点，但 `G-DECK` / `G-TGT` 的 `PASS` 须具名人工、绑定本次 PPTX 哈希；无人可签时如实保持 `BLOCK — Delivery` 并列出待签清单，不得自填 reviewer，也不得说“已通过验收”。最终回复的骨架见 `references/review-protocol.md` §6。
