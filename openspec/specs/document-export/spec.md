# document-export Specification

## Purpose
记录 UniReadMD 当前 HTML、PDF、PNG、JPEG 和可选 Pandoc 文档导出入口的能力基线，供后续任务
重建统一、带样式且受权限约束的导出管线。

## Requirements
### Requirement: Export Pipeline Baseline
当前应用 MUST 支持 HTML、PDF、PNG、JPEG 以及可选 Pandoc 文档导出入口。

#### Scenario: 用户选择导出格式
- **WHEN** 用户从菜单选择导出格式
- **THEN** 应用调用对应导出 IPC 并返回旧合同字段

### Requirement: Export Capability Baseline
当前应用 MUST 报告内建导出格式和 Pandoc 可用性。

#### Scenario: 查询导出能力
- **WHEN** renderer 查询导出能力
- **THEN** 主进程返回内建格式、Pandoc 状态和可用 Pandoc 格式
