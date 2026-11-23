# ppt-visual-art-director

以视觉艺术总监的判断制作、改稿或评审 PPT。**AI 负责设计判断，脚本负责确定性执行与事实核验，真实渲染负责证明表现，人负责最终签收。** 交付原生可编辑的 `.pptx`。机器检查不等于好看，预览不等于目标环境通过。

当前版本 **12.0.0**（见 `CHANGELOG.md`、`VERSION`）。

```bash
# 在技能包根目录执行（即本 README 所在目录）
pip install -r requirements.txt
python scripts/compile.py plan.json out.pptx [--assets DIR]     # 校验 + 编译
python scripts/verify.py  plan.json out.pptx [--assets DIR]     # 对象核对 + 渲染核对
python scripts/revise.py  inspect 原稿.pptx                     # 改稿：只读解析
python scripts/revise.py  原稿.pptx edit.json out.pptx          # 改稿：定点编辑
```

如果你在技能包的**上一级目录**（比如刚 clone 下来），把 `scripts/…` 换成 `ppt-visual-art-director/scripts/…`，或先 `cd ppt-visual-art-director`。两处文档以前各写一套路径，12.0.0 起统一为"在技能包根执行"。

## 仓库结构

```
ppt-visual-art-director/
  SKILL.md                    判断原则、判断顺序、任务路线、验收、交付
  references/design.md        概念、艺术指导、叙事、空间、信息设计、排印、色彩、诊断、复核
  references/imagery.md       图片与视觉资产
  references/build.md         必填与枚举速查、plan 语义、素材、图表、命令、诊断码、定点改稿、能力边界
  scripts/                    model · ooxml · charts · checks · compile · verify · revise
                              plan.schema.json（数据契约）  limits.json（技术阈值）
  requirements.txt            必需依赖（python-pptx · jsonschema · Pillow · lxml）
  requirements-render.txt     可选：渲染核验（PyMuPDF，AGPL-3.0；另需 LibreOffice）

```

打包分发只取 `SKILL.md`、`references/`、`scripts/`、`requirements*.txt`、`LICENSE`。

## 12.0.0 修订要点

* **判断依据修正**：`W-FIT` 装载模型按渲染实测校准（首行可见高度＋后续行距推进），同一份评测稿 **5 条误报 → 0**，真溢出仍被覆盖；`W-SPACE` 分级，单个中西文空格降为 `I-SPACE`（风格取舍），多空格/不断行空格/全角空格仍为 `W-SPACE`。
* **新增字面核验** `W-FONT-FACE`：声明简体、PDF 实际取到日文字面时报出来——家族匹配 ≠ 简体字形已核验。
* **首次写 plan 的坑写进 `build.md` 开头**：`page.role` 六值枚举、`motif`/`evidence` 必须是数组、`page_numbers.hide` 是页 id、`element` 拒绝未知键。
* 能力边界补记：柱/条/折线不支持按数据点着色，只有饼/环按类别取色。

11.0.0 的修订要点（母版按画布铺开、主题字体取 plan 字族、`W-FIT` 行高按字体取值等）已并入本版。

## 验收与验证等级

L0 对象 → L1 LibreOffice 渲染 → L2 PowerPoint 实机 → L3 人工签收；`verify.py` 只到 L1。
状态 `READY` / `PARTIAL` / `BLOCK` 都不是质量结论：`READY` 只说明对象与渲染一致、声明字族参与渲染。
未在真实 PowerPoint 中跑过的一律 `NOT VERIFIED`。

## 许可

本仓库代码与文档按 MIT 授权（`LICENSE`）。可选依赖 PyMuPDF 为 AGPL-3.0（或 Artifex 商业许可），本仓库不分发它，是否安装由使用者的合规要求决定，见 `requirements-render.txt`。仓库不含任何字体文件。
