# document-workspace Specification

## Purpose
记录 UniReadMD 当前 CodeMirror 编辑、Markdown 预览、视图切换、搜索、侧栏、基础设置和
预览编辑入口的工作区基线，供后续任务恢复核心编辑体验并补齐文件工作区。

## Requirements
### Requirement: Workspace and Editor Baseline
当前应用 MUST 提供 CodeMirror Markdown 编辑区、预览区、视图切换、搜索、侧栏和基础设置入口。

#### Scenario: 用户切换源码和预览
- **WHEN** 用户在源码、预览或双栏视图之间切换
- **THEN** 同一份 Markdown 文档内容继续作为编辑和预览来源

### Requirement: Preview Editing Baseline
当前应用 MUST 保留预览内联编辑、表格单元格编辑、Focus 和 Typewriter 的现有入口与快捷键。

#### Scenario: 预览编辑入口存在
- **WHEN** 用户在预览中触发可编辑块
- **THEN** 应用展示用于编辑 Markdown 片段的临时控件
