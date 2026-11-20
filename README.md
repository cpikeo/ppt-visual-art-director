# ppt-visual-art-director

以视觉艺术总监的判断制作、改稿或评审 PPT。**AI 负责设计判断，脚本负责确定性执行与事实核验，真实作品负责证明质量，人负责最终签收。**

交付原生可编辑的 PPTX（只输出 `.pptx`），并用对象核对与真实渲染给出证据。机器检查不等于好看，预览不等于目标环境通过。

## 仓库结构

```
技能包（交付给 AI 的部分，只含运行必需文件）
  SKILL.md               判断原则、判断顺序、brief、任务路线、验收、交付
  references/design.md   概念、艺术指导、叙事、空间、信息设计、排版、色彩、诊断、复核
  references/imagery.md  图片与视觉资产：资产职责、简报、提示词、验收
  references/build.md    plan 语义、几何、素材与生成图、图表、命令、诊断码、定点编辑、输入安全、能力边界
  scripts/               model · ooxml · charts · checks · compile · verify · revise
                         plan.schema.json（数据契约，唯一）  limits.json（技术阈值，唯一）
  requirements.txt       必需依赖
  requirements-render.txt  可选：渲染核验（含 AGPL 的 PyMuPDF）

开发材料（不属于技能包，打包给 AI 时排除整个 dev/）
  dev/tests/             回归测试（素材边界、简报约束、图上文字量测、XML/包安全）
  dev/evals/             评测样本与 FINDINGS.md（当前：art-direction-demo）
  dev/compat/COMPAT.md   兼容性与未验证清单
```

## 安装

- Python 3.9+（只在 3.11 上测试过）。
- 必需：`pip install -r requirements.txt`。可以编译、定点改稿、做对象核对。
- 可选渲染核验：`pip install -r requirements-render.txt`，并安装 LibreOffice（`soffice` 在 PATH，或设环境变量 `SOFFICE`）。不装时 `verify.py` 给出 `PARTIAL`，并说明哪些判断没有依据。
- 许可提醒：PyMuPDF 是 AGPL-3.0（或商业许可），因此放在可选文件里；本仓库不分发它。
- 字体：只写字体名、不嵌入。样本使用 Noto Sans/Serif CJK SC（OFL）；换成微软雅黑等字体后的断行需在目标机复核。
- 脚本不联网。

## 常用命令

```bash
python scripts/compile.py plan.json out.pptx [--assets DIR]   # 新建：校验 + 编译
python scripts/verify.py plan.json out.pptx [--assets DIR]    # 核验，写 render/<名称>/（素材目录须与编译一致）
python scripts/revise.py inspect 原稿.pptx                     # 改稿／评审：只读解析
python scripts/revise.py 原稿.pptx edit.json out.pptx          # 改稿：定点编辑（保留核对，原稿不动）

python -m unittest discover -s dev/tests -v                   # 回归（开发）
```

## 图片艺术指导（本次升级重点）

- `references/imagery.md`：按资产职责（背景、主视觉、插图、图标、人物与产品）给出不同的视觉语言、五句话简报、提示词顺序与三级验收。
- plan：`image_role` 增加 `hero` `illustration` `icon`；生成图的 `generation.brief`（职责、落字区、淘汰条件）随 plan 登记，缺失给 `W-ASSET`；`evidence` 角色指向生成图直接阻断。
- 编译期量测：文字压在图上时，按裁切、`duotone`、`alpha` 还原文字框正下方的像素，给出 `W-CONTRAST`（落点对比不足）与 `W-BUSY`（纹理过碎），不再只提示“去看渲染”。这是静态近似，不替代渲染。

## 工程安全

素材只从素材根内读取（拒绝绝对路径、`..`、指向根外的符号链接）；`compile` `verify` `revise` 共用同一素材根规则；解析外部 PPTX 的 XML 不展开实体、不联网；成品包拒绝越界部件名与超大解压体积。

## 评测证据在哪

`dev/evals/art-direction-demo/`：5 页演示稿的 plan、3 张生成图（含生成记录与简报）、编译产物、对象核验报告；`dev/evals/FINDINGS.md` 写出做对了什么、没做好什么，以及哪些**没有验证**。

限制：样本内容是作者自写的示例；评价是作者自评，不是第三方评审，也不是 PowerPoint 实机；编写本次演示稿的环境没有 LibreOffice，**没有渲染**，状态为 `PARTIAL`。

## 能力边界

| 级别 | 内容 |
| --- | --- |
| 完整支持 | 16:9 / 4:3；原生文本、形状、折线路径、组、图片裁切；原生图表（柱、线、面积、饼、环、散点）与表格；演讲者备注；替代文字；页码；既有稿定点编辑 |
| 有限支持 | 页标题是普通文本框（读屏识别不到）；无双数值轴、桑基、甘特；表格无行高亮；渐变与 duotone 在 LibreOffice 只是近似；连线是坐标而非连接符；4:3 未做样本评测 |
| 外部依赖 | 图片生成（工具无关：只登记工具、提示词、简报、文件、SHA-256、权利状态，不调用任何模型）；LibreOffice；PowerPoint 实机验证；目标机字体 |
| 不支持 | 动画、切换、嵌入字体、新建 SmartArt、宏（.pptm 拒绝） |

## 验证等级与状态

`verify.py` 的状态：`BLOCK` 有硬问题；`PARTIAL` 渲染没跑或声明字体被替换；`READY` 对象一致、渲染完成、字体都真实参与。三档都不是设计结论，`human` 永远是 `NOT REVIEWED`。

| 等级 | 内容 | 状态 |
| --- | --- | --- |
| L0 对象 | plan ↔ PPTX 结构与对象核对 | 已做（含本次演示稿） |
| L1 LibreOffice | 无头渲染、断行、字体、越界 | 作者历史记录为 LibreOffice 7.4.7.2 / Debian 12；**本次修订未复跑** |
| L2 PowerPoint 实机 | 真实 Microsoft PowerPoint 打开与导出 | **NOT VERIFIED**（从未做过） |
| L3 人工签收 | 目标环境下人确认 | **NOT REVIEWED** |

## 许可证

**本仓库目前没有 LICENSE 文件，没有声明任何开源许可证。** 在作者明确选择并添加许可证之前，默认保留所有权利，使用、再分发与二次开发需要先取得作者许可。选择哪种许可证是作者的决定，不由本文档代为指定；若希望被广泛使用，常见选择是 MIT 或 Apache-2.0（依赖均为宽松许可，兼容）。

第三方依赖各自许可：python-pptx（MIT）、jsonschema（MIT）、Pillow（HPND）、lxml（BSD）；可选的 PyMuPDF 为 AGPL-3.0 / 商业双许可。演示稿里的 3 张生成图（`dev/evals/art-direction-demo/assets/`）权利状态为 `pending`，只作为评测素材。
