# ppt-visual-art-director

以视觉艺术总监的判断制作、改稿或评审 PPT。**AI 负责设计判断，脚本负责确定性执行与事实核验，真实渲染负责证明表现，人负责最终签收。** 交付原生可编辑的 `.pptx`。机器检查不等于好看，预览不等于目标环境通过。

当前版本 **13.0.0**（见 `CHANGELOG.md`）。

```bash
pip install -r requirements.txt            # 渲染核验另装 requirements-render.txt 与 LibreOffice
python scripts/compile.py plan.json out.pptx [--assets DIR] [--report r.json] [--force]
python scripts/verify.py  plan.json out.pptx [--assets DIR] [--render-dir DIR] [--report v.json] [--force]
python scripts/revise.py  inspect 原稿.pptx [--json i.json] [--force]
python scripts/revise.py  原稿.pptx edit.json out.pptx [--assets DIR] [--report r.json] [--force]
```

## 正式包（打包分发只取这些）

```
SKILL.md                 判断原则、任务路线、验收、交付
references/              design.md · imagery.md · build.md（按需读取）
scripts/                 model · ooxml · charts · checks · compile · verify · revise
                         plan.schema.json（数据契约）· limits.json（技术阈值）
requirements.txt · requirements-render.txt · LICENSE
```

`python dev/build_package.py` 生成 zip，并断言没有 dev/ 内容混入。

## 开发与证据（不属于正式包）

```
dev/works/       五套完整作品：plan 生成脚本、plan.json、PPTX、LibreOffice 渲染 PNG、verify 报告、修正前截图
dev/evals/       12.0.0 基线（同一批 plan）、同页前后对照、定点改稿演示、整稿总览、环境记录
dev/tests/       33 项回归测试（正常样本 + 故意破坏的负向样本）
dev/README.md    作品集设计检查、问题清单、验证分层与结论
```

运行：`pip install -r requirements-render.txt pytest && python dev/works/run.py && python -m pytest dev/tests -q`

## 验证等级

L0 对象 → L1 LibreOffice 渲染 → L2 PowerPoint 实机 → L3 人工签收。`verify.py` 只到 L1。
`READY` / `PARTIAL` / `BLOCK` 都不是质量结论。**未在真实 PowerPoint 中跑过的一律 NOT VERIFIED。**

## 许可

MIT（`LICENSE`）。可选依赖 PyMuPDF 为 AGPL-3.0（或 Artifex 商业许可），本仓库不分发它。仓库不含字体文件。
