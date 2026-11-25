# 16.0.0（相对 15.3.0）

修复
- scaffold：未开母题横线时，眉题仍右移 0.85 in，与页题错轴（孤儿标签）。现与页题同轴。
- limits.json：`Noto Sans/Serif SC|TC|JP|KR` 家族名没有行高条目，落到默认 1.25，巨型数字框偏矮；补 8 个键（实测 1.448）。
- verify：越界判断用 span.bbox（字体 ascent+descent 行盒）→ 大字号数字误报 V-OVERFLOW；改按墨迹范围（基线上 0.95em/下 0.2em）。
- scaffold list：条目固定 +6pt 缩在角落；改为取最长条目不折行的最大字号（上限 body+12）。
- scaffold close：标题折行/大字号档把条目推出页底（W-GEOM）；整块上移，不压字号。
成本
- 诊断同级别/同代码/同文案多页合并一行（`model.show`）；JSON 报告仍逐条。
- 新增 `scaffold.py fields <节点>`：约 1KB 紧凑字段表，代替通读 40KB schema。
- SKILL.md 去掉重复的“按需加载”表与重复叙述；build.md 移出模块依赖、机器边界、网格算术、写作顺序；design.md 移出 3 条工具自身历史教训（见 dev/HISTORY.md）。
- 清理未使用 import。
隔离
- 测试与案例在 dev/（不进 Skill 包）：`python -m unittest discover -s dev/tests -v`。
