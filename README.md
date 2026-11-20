# ppt-visual-art-director

以视觉艺术总监的判断制作、改稿或评审 PPT。**AI 负责设计判断，脚本负责确定性执行与事实核验，真实渲染负责证明表现，人负责最终签收。** 交付原生可编辑的 `.pptx`。机器检查不等于好看，预览不等于目标环境通过。

## 仓库结构

```
ppt-visual-art-director/      ← 正式技能包（只打包这个目录）
  SKILL.md                    判断原则、判断顺序、任务路线、验收、交付
  references/design.md        概念、艺术指导、叙事、空间、信息设计、排印、色彩、诊断、复核
  references/imagery.md       图片与视觉资产
  references/build.md         plan 语义、素材、图表、命令、诊断码、定点改稿、能力边界
  scripts/                    model · ooxml · charts · checks · compile · verify · revise
                              plan.schema.json（数据契约）  limits.json（技术阈值）
  requirements.txt            必需依赖（python-pptx · jsonschema · Pillow · lxml）
  requirements-render.txt     可选：渲染核验（PyMuPDF，AGPL-3.0；另需 LibreOffice）
dev/                          ← 不属于技能包
  tests/                      回归测试（14 项）
  evals/                      三套不同主题的真实评测稿（plan、pptx、LibreOffice 渲染图、报告）、FINDINGS.md
  compat/COMPAT.md            已验证 / 依赖环境 / 未验证清单
```

## 命令

```bash
pip install -r ppt-visual-art-director/requirements.txt
python ppt-visual-art-director/scripts/compile.py plan.json out.pptx [--assets DIR]
python ppt-visual-art-director/scripts/verify.py  plan.json out.pptx [--assets DIR]
python ppt-visual-art-director/scripts/revise.py inspect 原稿.pptx
python ppt-visual-art-director/scripts/revise.py 原稿.pptx edit.json out.pptx
python -m unittest discover -s dev/tests -v        # 回归
bash dev/evals/run_all.sh                          # 重建评测稿并渲染核验（需 LibreOffice）
```

## 11.0.0 修订要点

* **交付质量**：母版与版式按画布铺开（原来 16:9 成品里仍是 4:3 占位符）；主题字体改为 plan 字族（原来是 Calibri＋空东亚字体）；新建文字带 `lang`。
* **判断依据修正**：`W-FIT` 行高按字体取值（Noto CJK 1.45，原来一律 1.2，真实渲染里大字与两行标题溢出）；`W-SPACE` 理由改为实测依据；巨型数字/标题可作承载层（字体即图像）；声明出血的文字不再被 `V-CLIP` 阻断（改 `V-BLEED` 提示）；环/饼图例不再误报缺系列名。
* **结构**：技能包与开发材料分开；README 不再引用不存在的文件；`design.md` 十补入 4 条来自真实渲染的判断。
* **证据**：三套评测稿（深色戏剧场、暖纸编辑场、钴蓝纯排版）实际编译与渲染，见 `dev/evals/FINDINGS.md`。

## 许可证

本仓库没有 LICENSE 文件，选择哪种由作者决定。依赖：python-pptx（MIT）、jsonschema（MIT）、Pillow（HPND）、lxml（BSD）；可选 PyMuPDF 为 AGPL-3.0 / 商业双许可。评测稿中的图片为 AI 生成，权利状态 `pending`，只作评测素材。
