# Changelog

本项目使用语义化版本。机器契约变化必须同时更新 Schema、消费者、测试和迁移说明。

## 2.1.1 — 2026-10-22

预检可重复执行、阈值只有一份定义。不改设计取向与判据取值。

### Fixed

- **预检不能重跑**：`scripts/release_check.py` 复用同一个 `--out` 目录时，`package_skill.py`
  因目录非空在第 7 阶段失败——前 6 个阶段白跑，且旧 `manifest.json` 与旧日志可能被当成本轮证据。
  现在开跑前只清理本工具自己的产物（`package/`、`pvad-skill.zip`、`manifest.json`、`*.stdout.txt`/`*.stderr.txt`），
  其他文件保留，并在 manifest 的 `reset` 字段如实列出清理项。
- **预检证据不绑定代码版本**：manifest 现在写入开跑前采集的 `git.commit` / `git.dirty`
  （未安装 git 时如实写 `UNKNOWN`），并记录 `package_zip_sha256` 与 `package_zip_bytes`；
  打包阶段退出码为 0 却没有 zip 时整次判 FAIL，不再用退出码冒充产物。
- **Schema 与消费者的必填口径不一致**：`thresholds.size_pt_tolerance`、`thresholds.chart_scale_skew_ratio`
  被 `fidelity.py` / `guard.py` 以 `policy.num()`（缺键即报错）读取，却没进 Schema 的 `required`——
  缺键策略能通过校验、在运行时才炸。两者已补进 `required`（策略文件本身早已给出，取值不变）。
- **Python 里另立阈值默认值**：`fidelity.py` 的 `SIZE_PT_TOL = 0.75` 从未被使用，
  `POSITION_TOL` / `SIZE_TOL` 的元组第二项（`1.5`）也从不生效。阈值改为只从 Policy 读取，删除这些影子常量。

### Added

- `TestPolicySchemaCoversConsumers`：代码按必填读取的每个 `thresholds.*` 键都必须在 Schema 的 `required` 里；
  Python 不得再出现带数值默认值的 `*TOL*` / `*THRESHOLD*` 常量。
- `TestReleasePreflightRepeatable`：清场只删自己的产物、`git_state` 绑定或如实 UNKNOWN、
  manifest 的包指纹与本次 zip 一致、同目录重跑成功、无产物必判 FAIL。

## 2.1.0 — 2026-10-22

技能包层面的收敛：不改设计取向，只让判断被完整传达、让验证说真话。

### Fixed

- **架构命名**：`SKILL.md` 标题从「四层架构」改为「五层架构与权威源」，与表内 5 行一致；
  同步 6 处引用、guard 的 disclaimer 文案、`policy/prompts.yaml` 注释与 2 个测试 docstring；
  补「层间只允许单向依赖」一句。
- **测试套件不绿**：补回缺失的 `.github/workflows/ci.yml`（3.10 / 3.13 矩阵），两个
  `TestCiContract` 用例恢复，且 CI 仍显式声明不验证目标 PowerPoint 和外部门。
- **图表单位丢失**：brief 的 `chart.unit` 进入 Schema 与 BUILD；编译写出 `unit_suffix`，
  回读比对单位丢失即 MISMATCH。此前「占比图」会继承整稿的「万元」并被 release 判 PASS。
- **柱形数据标签**：显式 `outEnd`，标签不再压在柱顶被填充色吃掉。
- **死阈值**：`guard._check_density` 读了 `thresholds.numeric_density_high` 却用内联常量 6；
  改为消费新键 `thresholds.sparse_max_numbers`（Schema + 语义同步）。
- **断行碎片无告警**：新增 `LEAD_FRAGMENT`（续行以单个汉字开头后接数字/外文 = 紧密词组被拆开），
  阈值 `typography.orphan.min_lead_chars`。基线案例 s02 标题的「门店」被拆一事现在会被报出。
- **契约缺项**：`contract.md` 补 `EVIDENCE_NOT_LAID_OUT`、`SOURCE_ZONE_COLLISION` 两个阻断码和
  7 个告警码，修正 `5.0` 编号；`policy.md` 的 `make eval` 改为真实命令。
- **陈旧证据**：四份评测报告重新生成并写入 `policy_sha256`；`stress_eval.py` 之前直接
  `ModuleNotFoundError`，已修。治理登记表与验收配置的 `repository_revision` 更新到当前基线。
- **死代码**：`scripts/ evals/ tests/` 共 27 处 pyflakes 问题（未用导入/未用局部变量/变量遮蔽）清零；
  其中 3 个测试模块对 `support` 的导入改为真实用途（仓库夹具自检），不再依赖裸副作用导入。

### Changed

- `SKILL.md` 不再复制参考文档的细节：§4.3/§4.4/§5.2/§8.1 改为「保留可执行规则 + 指向权威文档」，
  冲突优先级只在 §3.1 定义一次，排版优先级明确「只在排版内部生效」。
- 三份参考文档补「权威范围 / 不负责」声明，避免同一原则两处定义。
- `verification.md` 证据表由 6 层扩为 7 层，与 `SKILL.md`「证据层级」同一份清单。

### Added

- 测试能力补强（不是放宽）：错误码 ↔ `contract.md` 一致、文档引用文件必须存在、
  文档命令必须有入口、`policy.*` 读取结果必须被使用（AST 检查）、评测报告必须绑定当前策略指纹。
- 反向验证用例：单位丢失 → fidelity MISMATCH；续行单字碎片 → 告警；CI 缺失 → 测试失败（已修复）。

### Evidence limits

- 仍然没有目标 PowerPoint、整稿人工复核与素材合规证据：`delivery_status` 保持 `BLOCK — Delivery`；
- 新增告警是「可被看见」，不是「已修好」：基线 s02 标题的断行问题仍由作者改写处理；
- 评测报告只证明离散判断/度量模型内部一致，不证明视觉质量。

### Migration

- 需要图表单位的 brief 补 `chart.unit`（可选字段，旧 brief 行为不变）；
- 策略新增两个必填键 `thresholds.sparse_max_numbers`、`typography.orphan.min_lead_chars`，
  自定义策略目录必须同步，否则 `policy.load` 直接报 `策略结构非法`；
- 修改策略后必须重新生成 `evals/reports/*.json`，否则 `TestEvidenceMatchesCurrentVersion` 失败。

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
