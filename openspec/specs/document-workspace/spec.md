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

### Requirement: 表格单元格原位可视编辑
应用 MUST 在普通表格单元格内保留样式直接编辑，通过 CodeMirror 单一源码保存内容，
并保留中文输入、列对齐、转义竖线、导航、撤销和保存流程。

#### Scenario: 单击格式化单元格
- **WHEN** 用户单击含正文、加粗或链接等普通行内格式的单元格
- **THEN** 原单元格成为输入区域，不出现独立编辑框，进入时保持原有字号、行高和位置

#### Scenario: 导航与保存
- **WHEN** 用户输入后按 Tab 或 Enter 导航，再保存并重新打开
- **THEN** 输入内容和表格结构保持一致，最后一行导航可按既有规则追加行

#### Scenario: 复杂内容兼容
- **WHEN** 单元格含图片或公式，或用户按 Alt 单击单元格
- **THEN** 保留已有源码编辑方式，不破坏原有 Markdown 内容

### Requirement: 一致且稳定的工作区右键菜单
应用 MUST 统一工作区右键菜单的字体、间距、背景、阴影和分隔线，按实际尺寸限制弹出位置，
并保留原有命令、主题适配及键盘入口。

#### Scenario: 在不同边缘打开菜单
- **WHEN** 用户在工作区顶部、中部或靠近底边右键打开插入、表格或图片菜单
- **THEN** 菜单靠近点击位置且完整可用，不因预估高度而无故向上偏移或被裁剪

#### Scenario: 关闭与互斥
- **WHEN** 用户按 Escape、点击外部、滚动或切换编辑模式
- **THEN** 当前菜单关闭，打开另一类菜单时不会保留旧菜单
