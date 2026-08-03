## ADDED Requirements
### Requirement: Unified Styled Export
系统 MUST 使用统一导出文档生成器，保留预览样式、图片、公式、Mermaid 和长文布局。

#### Scenario: 内建导出与预览一致
- **WHEN** 用户导出 HTML、PDF、PNG 或 JPEG
- **THEN** 导出内容包含净化资源，并与应用预览在固定像素阈值内一致

### Requirement: Optional Pandoc Lane
系统 MUST 将 DOCX、ODT 和 RTF 视为可选 Pandoc 能力，并固定验证 Pandoc 3.10。

#### Scenario: Pandoc 缺失时文档导出禁用
- **WHEN** 系统未找到唯一的 Pandoc 3.10 候选
- **THEN** DOCX、ODT 和 RTF 显示禁用状态并返回 `ERR_PANDOC_UNAVAILABLE`

## MODIFIED Requirements
### Requirement: Export Pipeline Baseline
当前应用 MUST 支持 HTML、PDF、PNG、JPEG 以及可选 Pandoc 文档导出入口，并确保目标路径
只来自对话框且格式扩展名匹配。

#### Scenario: 用户选择导出格式
- **WHEN** 用户从菜单选择 HTML、PDF、PNG、JPEG、DOCX、ODT 或 RTF
- **THEN** 应用调用对应导出 IPC，验证路径和大小限制，并返回旧合同字段

### Requirement: Export Capability Baseline
当前应用 MUST 报告内建导出格式和 Pandoc 可用性，并把缺能力、超限和取消场景表达为稳定状态。

#### Scenario: 查询导出能力
- **WHEN** renderer 查询导出能力
- **THEN** 主进程返回内建格式、Pandoc 状态、可用 Pandoc 格式和非敏感 trace 日志
