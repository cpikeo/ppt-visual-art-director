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
| `render/pie-donut-contract.png` | 图表契约验证页（饼/环逐扇区着色 + 仅百分比标注） | 验证"类别不靠单一颜色编码 + 标注可读" |

## 复现命令

```bash
python -m pytest dev/tests/test_render.py        # 生成产物到 pytest 临时目录并逐项断言
```

固定环境：LibreOffice Impress `25.2.3.2` + poppler `pdftoppm` 110 dpi + `Noto Sans/Serif CJK SC`。

## 使用边界

- 这些 PNG 只证明**在该渲染链下**的几何、字体、裁切与标注事实，不证明设计合格。
- 新增版本时请按同一案例、同一 dpi、同一字体重新生成，**覆盖或并列存放均可**，但不得改名成带版本号的名称。
- 盲评时由第二位评审者按随机顺序打开，评审者事先不应知道哪一份来自改动前、哪一份来自改动后。
