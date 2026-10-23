# 基准渲染产物（去标识）

本目录保存**固定案例**在各版本的渲染结果，用于版本对照与人工盲评。文件名只写页面角色，
**不含版本号、模型名或作者信息**，以保证盲评时评审者无法从文件名推断来源。

## 内容

| 文件 | 来源 | 说明 |
| --- | --- | --- |
| `render/fixture-p1-cover.png` | `dev/tests/fixtures/render_deck.json` 第 1 页 | 深色场封面：衬线标题、发丝线、页脚层级 |
| `render/fixture-p2-chart.png` | 同上第 2 页 | 原生柱状图：六季度序列、无重复图例、浅发丝网格 |
| `render/fixture-p3-table.png` | 同上第 3 页 | 原生三线表：无竖线、无斑马纹、数值列右对齐 |
| `render/fixture-p4-image.png` | 同上第 4 页 | 16:9 源图 → 竖框等比居中裁切 |
| `render/pie-donut-contract.png` | 历史图表契约参考页 | 仅作历史参考；缺少独立固定 plan/PPTX SHA，不纳入本轮像素一致结论 |

## 复现命令

```bash
python -m pytest dev/tests/test_render.py        # 生成产物到 pytest 临时目录并逐项断言
```

固定环境：LibreOffice Impress `25.2.3.2` + poppler `pdftoppm` 110 dpi + `Noto Sans/Serif CJK SC`。该命令复现固定四页及当前测试临时产物，不能凭空复现缺少固定输入的历史饼/环参考图。

## 使用边界

- 这些 PNG 只证明**在该渲染链下**的几何、字体、裁切与标注事实，不证明设计合格。
- tracked baseline 是不可覆盖的历史证据。新增版本须按同一案例、同一 dpi、同一字体生成到独立运行目录，并记录 Git SHA、PPTX SHA 与环境；不得覆盖旧文件后仍称其为优化前基线。
- 盲评时由非评审者复制到去标识目录、随机化顺序并保管版本映射；文件名和可见标签不得泄露版本身份。评审者事先不应知道哪一份来自改动前或改动后。
