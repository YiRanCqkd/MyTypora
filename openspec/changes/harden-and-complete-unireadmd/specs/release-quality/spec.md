## ADDED Requirements
### Requirement: Evidence-Gated Release Quality
系统 MUST 以固定工具链、七文件证据集、覆盖率、审计、性能和最终审核作为发布门禁。

#### Scenario: Green 验证可重放
- **WHEN** Final verifier 读取每个任务 manifest
- **THEN** 只重放幂等 Green 命令，比较退出码和哈希，并拒绝缺失或篡改证据

### Requirement: Scoped Release Change List
系统 MUST 生成从 Task 1 base SHA 到当前 HEAD 的稳定 change list，并按 C1-C6 scope-map 校验。

#### Scenario: Scope OUT 文件被拒绝
- **WHEN** changed-files 包含未映射文件或 Scope OUT glob
- **THEN** scope verifier 返回 `REJECT:SCOPE_OUT_FILE` 或对应拒绝码

## MODIFIED Requirements
### Requirement: Quality Gate Baseline
当前项目 MUST 提供前端质量命令、Electron 打包命令和基础 CI 工作流，并将其收敛到统一 helper
和发布门禁。

#### Scenario: 开发者运行发布质量命令
- **WHEN** 开发者执行最终 quality 或 release verifier
- **THEN** 项目按固定顺序运行类型检查、Lint、测试、构建、审计、E2E、性能和 package smoke

### Requirement: Toolchain Baseline
当前项目 MUST 记录本地 Task 1 工具链，并通过 resolver 返回 Node、npm CLI 和 Git 的绝对路径；
候选缺失、重复、版本或哈希变化 MUST 返回固定错误码。

#### Scenario: 解析固定工具链
- **WHEN** resolver 使用真实 baseline 运行
- **THEN** 返回 Node.js 24.8.0、npm CLI 和 Git 2.47.1.windows.2 的绝对路径

#### Scenario: 工具链 baseline 被篡改
- **WHEN** 临时 baseline 副本中的版本或 SHA-256 被修改
- **THEN** resolver 返回非 0，并输出 `ERR_TOOLCHAIN_BASELINE_CHANGED`
