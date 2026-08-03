# document-integrity Specification

## Purpose
记录 UniReadMD 当前文档内容、文件路径、最近文件、会话快照、脏状态和保存入口的完整性基线，
供后续任务补齐未保存保护、会话恢复和大文件边界。

## Requirements
### Requirement: Document State Baseline
当前应用 MUST 维护编辑内容、文件路径、最近文件、会话快照和脏状态提示的基础状态。

#### Scenario: 打开文档后更新状态
- **WHEN** 用户通过打开文件流程载入 Markdown
- **THEN** 当前内容、文件路径和最近文件状态被更新

### Requirement: Save Flow Baseline
当前应用 MUST 支持手动保存和另存为，并返回旧合同字段。

#### Scenario: 手动保存返回文件路径
- **WHEN** 用户保存当前 Markdown
- **THEN** 主进程返回保存后的文件路径和内容状态
