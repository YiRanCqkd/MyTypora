# release-quality Specification

## Purpose
记录 UniReadMD 当前质量命令、Electron 打包、CI 工作流和 Task 1 固定工具链基线，供后续任务
收紧覆盖率、审计、性能、真实 Electron 和发布证据门禁。

## Requirements
### Requirement: Quality Gate Baseline
当前项目 MUST 提供前端质量命令、Electron 打包命令和基础 CI 工作流。

#### Scenario: 开发者运行基础质量命令
- **WHEN** 开发者执行现有质量脚本
- **THEN** 项目按当前脚本顺序运行类型检查、Lint、测试和构建

### Requirement: Toolchain Baseline
当前项目 MUST 记录本地 Task 1 工具链，并通过 resolver 返回 Node、npm CLI 和 Git 的绝对路径。

#### Scenario: 解析固定工具链
- **WHEN** resolver 使用真实 baseline 运行
- **THEN** 返回 Node.js 24.8.0、npm CLI 和 Git 2.47.1.windows.2 的绝对路径
