## ADDED Requirements

### Requirement: Stable Preview Editing Layout
系统 SHALL 在进入预览编辑时保持所在内容的字体和行高，并完整展示编辑内容。

#### Scenario: 长段落和标题编辑
- **WHEN** 用户编辑标题或长段落，输入换行或调整窗口宽度
- **THEN** 编辑字号一致、内容不被截断，后续正文保留布局空间

#### Scenario: 短格式片段编辑
- **WHEN** 用户编辑短加粗或斜体片段
- **THEN** 编辑占位按源码宽度调整，后续文字正常显示，提交和取消清除临时布局

### Requirement: Empty Document Writing
系统 SHALL 提供空白文档的直接预览输入入口，保持 Markdown 唯一数据源。

#### Scenario: 新文档开始写作
- **WHEN** 用户点击空白文档提示并输入
- **THEN** 内容写回 CodeMirror，可撤销，提示文字不进入保存或导出内容

### Requirement: Readable Responsive Workspace
系统 SHALL 提供适合连续阅读的正文行宽和行距，并在窄窗口保持紧凑工具栏。

#### Scenario: 窄窗口使用次要工具
- **WHEN** 工作区宽度不足以展示全部工具，用户打开更多工具
- **THEN** 常用工具仍可见，次要操作可访问，Escape 关闭菜单并恢复焦点

#### Scenario: 深色状态栏可读性
- **WHEN** 用户在深色主题下选中或悬停状态栏按钮
- **THEN** 按钮背景沿用深色主题变量，文字对比度符合 WCAG AA 要求

#### Scenario: 预览编辑期间切换专注模式
- **WHEN** 用户在宽屏预览编辑中切换专注模式
- **THEN** 正文与编辑框保持原有行宽，键盘焦点保留在编辑框内
