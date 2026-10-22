# Changelog

本项目使用语义化版本。机器契约变化必须同时更新 Schema、消费者、测试和迁移说明。

## 2.0.0 — 2026-10-22

### Added

- 六门发布模型、外部 release evidence 散列绑定和确定性交付状态；
- `handoff`：从 plan 派生最小 THINK→BUILD 交接；
- `impact`：比较前后 deck JSON 并给出保守验证范围；
- `evidence-init`：预填制品散列、实际图片和字体，所有人工门默认 FAIL；
- `inspect`：只读提取既有 PPTX 的可观察对象和保守工程风险；
- Draft / Review / Release 分阶段交付包及 SHA-256 清单；
- manifest、release evidence、inspection、benchmark case Schema；
- 设计失败案例、改稿工作流和独立基准治理协议；
- 20/50 页工程压力评估；
- 不可信 PPTX 的压缩包、页数、对象数和图片像素资源预算；
- 图片替代文本、Z 序、链接、备注、母版版式和字体诊断；
- inspection v1→v2 迁移器、固定时间戳/权限的确定性 ZIP、统一 `release_check.py`；
- 稳定语义锚点及文档引用测试。

### Changed

- `check`、`impact` 对 `.py` 作者稿默认拒绝执行；Draft/Spec 必须显式传入 `--allow-author-code`，Release 阶段无条件禁止执行作者代码；
- 长稿联系表改为策略控制的流式缩略图，单页原始证据不降分辨率；
- `package` 现在校验 manifest Schema、PPTX SHA-256、release evidence SHA-256、允许文件类型、符号链接、输出覆盖和复制后散列；
- manifest 预览路径写为绝对路径；
- release evidence 中素材 `generated` 从布尔值改为 `yes / no / unknown`；`unknown` 不能通过合规门；
- 生成素材要求必要专业复核为 `pass`；一般素材只要求具名复核者确认本项目预期用途已获授权，不强制无关权利；
- 判断评估增加强顺序用例，消融词表全部获得独立覆盖。

### Fixed

- 比率主张不再被 BUILD 改写为绝对差；
- BUILD 阻断实质未决问题；
- manifest 区分工程状态、交付状态和六道质量门；
- 章节重排不再使源码和评估中的裸章节号引用失效；
- 图片测试夹具和若干无消费者实现已清理或修复。

### Migration

1. 执行可信作者 Python 稿：

   ```bash
   python scripts/pvad.py check author_deck.py output.pptx --allow-author-code
   ```

   该代码不在沙箱中运行；来源不可信时必须转换为 `deck.json`。

2. 旧 release evidence 若使用 `generated: true/false`，需迁移为：

   ```yaml
   generated: yes   # 或 no / unknown
   ```

3. 旧 manifest 没有 `security` 或不符合 `schema/manifest.schema.json` 时，必须重新运行 `check`，不能直接用于 `package`。

4. `package` 输出目录必须不存在或为空；PPTX 被修改后必须重新运行 `check`，旧 evidence、manifest、摘要和 Release 包全部失效。

5. 旧 `pvad-inspection-v1` 可运行 `python scripts/migrate.py inspection OLD.json --out NEW.json` 升级为 v2；迁移后仍须人工复核读取顺序和目标字体。

6. Python 最低支持版本为 3.10。目标 PowerPoint 验证仍是外部门，不由 CI 推定通过。

### Evidence limits

- 内部渲染不是 Microsoft PowerPoint 证据；
- 合成压力测试不证明视觉质量或真实业务能力；
- 内部判断用例是回归集，不是独立盲标；
- 没有真实评审者、目标环境和合规责任人时，交付继续保持 `BLOCK — Delivery`。
