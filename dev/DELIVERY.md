# 交付验收与证据协议

本文只说明如何配置和执行验收。字段、错误严重度、策略数值和门定义分别以 `scripts/plan.schema.json`、`scripts/policy.json` 为准；`dev/acceptance.json` 是本仓库的案例配置，不是其他项目的默认值。

## 1. 交付前六问

1. 每页主张与证据是否一致？未知事实是否标为未决？
2. 实际 PPTX 是否保留类别、系列、数值、单位、期间和编辑能力？
3. 强制验证是否执行？非强制未验证项与风险是否有独立记录？
4. 原始素材与交付裁切是否同时可追溯，授权、隐私和生成属性是否成立？
5. 本次 PPTX 的真实预览是否存在中文断行、溢出、异常字体或渲染不一致？
6. 指定 PowerPoint、人工整稿与独立评审是否实际完成？没有证据就写 `NOT VERIFIED`，不能写通过。

## 2. 最小配置

```json
{
  "required_gates": ["G-STR", "G-SEM", "G-PRE", "G-DECK", "G-TGT", "G-COMPLY"],
  "key_pages": ["cover", "chart", "closing"],
  "high_risk_pages": ["chart"],
  "historical_issue_pages": [],
  "target_powerpoint": {
    "configured": false,
    "version": null,
    "operating_system": null,
    "fonts": [],
    "status": "PENDING_CONFIGURATION"
  },
  "sampling": { "low_risk_rate": 0.2, "seed": 42 },
  "reviews": {
    "G-DECK": { "status": "NOT_VERIFIED", "evidence": null, "reviewer": null },
    "G-TGT": { "status": "NOT_VERIFIED", "evidence": null, "reviewer": null }
  },
  "risk_approvals": []
}
```

缺省配置的强制门由 `scripts/policy.json` 的发布策略决定；正式项目应显式声明 `required_gates`，不能利用缺省降低要求。配置不可读、引用不存在页面或包含未知门 ID 均为冲突。关键页、高风险页和历史缺陷页始终进入样本，其余低风险页按固定种子抽样；复杂图表、素材页及发现系统性问题的同类页必须扩检，不能只靠抽样比例替代。

## 3. 六个稳定门 ID

| 门 | 验证内容 | 未执行或未完成时的含义 |
| --- | --- | --- |
| G-STR | Schema、policy、非法值、重复 ID、引用、有限几何与可编译结构 | 强制时 `BLOCK — Delivery` |
| G-SEM | plan 与实际 PPTX 的可见文本、图表类别/系列/数值、表格载荷和语义注释 | 强制时阻断；自动数字筛查不是事实真实性证明 |
| G-PRE | 本次 PPTX 经 LibreOffice/poppler 渲染后的页数、文本墨迹边界和字体记录 | 强制时阻断；只证明当前渲染环境 |
| G-DECK | 固定种子页面抽样、必要页覆盖、整稿人工检查记录 | 非强制未执行为 `NOT VERIFIED` |
| G-TGT | 指定 Microsoft PowerPoint 版本、OS、字体及关键页人工检查 | 未配置为 `NOT VERIFIED`；强制时阻断 |
| G-COMPLY | 素材文件、登记完整性、授权、隐私、生成属性与证据用途 | 强制时阻断；未决授权/隐私不能豁免 |

`G-PRE` 通过不等于 `G-TGT` 通过；`G-DECK` 通过不等于独立盲评通过；自动测试通过也不等于设计质量通过。

## 4. 执行命令

安装 wheel 后使用公开入口，避免依赖仓库内隐式模块路径：

```bash
ppt-vad plan.json output.pptx --assets . --strict
ppt-vad-verify plan.json output.pptx \
  --config dev/acceptance.json --render-dir empty-render-dir \
  --report delivery-report.json
```

验证器退出码：`0` 表示 `PASS — Delivery` 或 `PASS — Accepted Risk`；`1` 表示 `BLOCK — Delivery`；`2` 表示仅剩非强制项目时的 `NOT VERIFIED`。配置或输入冲突会进入报告并阻断，不得把它解释成通过。报告保留 plan/PPTX/config/policy SHA、每门证据、冲突、抽样种子和验证环境。

人工复核写入配置的 `reviews`，风险接受写入 `risk_approvals`；二者都不是让自动门“看起来通过”的通用开关。`--render-dir` 必须是空目录，预览渲染依赖 LibreOffice、poppler（`pdftoppm`、`pdfinfo`、`pdftotext`、`pdffonts`）；缺失时应明确失败或标记未验证，不得沿用旧 PNG。

## 5. 风险接受与不可豁免项

风险接受只适用于配置为**非强制**且 `scripts/policy.json` 允许的风险。必须同时记录：

- 涉及门 ID 与具体范围；
- 影响、已有替代证据与剩余风险；
- 缓解措施、责任人、批准人和失效条件；
- 独立 reviewer 与可定位 evidence。

缺少任一关键记录、批准者越权、用口头“整体可接受”代替具体范围，均无效。硬结构错误、事实/数据错误、授权或隐私阻断、权威源冲突及 MUST FIX 不可通过风险接受变成通过。

## 6. 状态与对外声明

- **PASS — Delivery**：所有强制门通过，且无不可豁免问题；只在证据实际覆盖的范围内声明。
- **PASS — Accepted Risk**：所有强制门通过，仅非强制风险按完整记录接受；对外仍须披露未验证范围。
- **NOT VERIFIED**：仅非强制检查尚未完成且没有有效风险接受；不是通过。
- **BLOCK — Delivery**：任一强制门未完成/未通过，或存在硬错误、事实/授权/隐私/权威冲突、MUST FIX。

“内部预览”“质量门通过”“目标环境通过”“独立盲评通过”是四种不同声明。没有 Microsoft PowerPoint 实测时，应写“LibreOffice 预览已验证；PowerPoint 尚未验证”，不能简写为“PPT 全环境无问题”。

## 7. 留档与复现

每次候选版本至少保存：输入 plan、编译日志、PPTX、交付报告、配置与策略 SHA、渲染器/字体环境、PDF/PNG、人工检查记录，以及（若执行）盲评原始表。失败日志也必须保留，不能只留最后一次绿色输出。

版本对照使用同一案例、权重、DPI 和字体条件；基线文件不得覆盖。独立评审应去除版本身份并随机顺序；自评只能标为自评。详细案例和评分锚点见 `dev/evals/eval_cases.json` 与 `dev/evals/rubric.md`。
