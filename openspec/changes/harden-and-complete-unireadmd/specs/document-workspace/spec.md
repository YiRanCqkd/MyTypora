## ADDED Requirements
### Requirement: Typora-Style In-Place Block Editing
系统 MUST 提供计划内 Typora 式原地块编辑，并以 CodeMirror Markdown 作为唯一数据源。

#### Scenario: 原地编辑写回 Markdown
- **WHEN** 用户在预览中编辑段落、标题、列表、引用、代码块或表格单元格
- **THEN** 应用通过 CodeMirror transaction 更新 Markdown，并保留统一撤销栈

### Requirement: Safe Workspace and Settings
系统 MUST 提供新建文件、递归文件夹树、排序、视图、拼写和安全用户 CSS 设置。

#### Scenario: 文件夹树扫描受限
- **WHEN** 用户打开包含深层目录、符号链接或大量 Markdown 文件的文件夹
- **THEN** 主进程按深度、数量、扩展名和 symlink 策略返回结果或稳定错误码

## MODIFIED Requirements
### Requirement: Workspace and Editor Baseline
当前应用 MUST 提供 CodeMirror Markdown 编辑区、预览区、视图切换、搜索、侧栏、文件夹树和
设置入口，且所有编辑视图共享同一 Markdown 数据源。

#### Scenario: 用户切换源码、预览和原地编辑
- **WHEN** 用户在源码、预览、双栏或原地块编辑之间切换
- **THEN** 同一份 CodeMirror Markdown 文档继续作为编辑、预览、保存和导出的来源

### Requirement: Preview Editing Baseline
当前应用 MUST 保留预览内联编辑、表格单元格编辑、Focus 和 Typewriter 的现有入口与快捷键，
并补齐撤销、焦点、滚动和失败处理。

#### Scenario: 预览编辑不创建第二文档模型
- **WHEN** 用户提交或取消预览编辑控件
- **THEN** 应用写回或丢弃临时控件内容，不生成新的富文本内核或独立状态源
