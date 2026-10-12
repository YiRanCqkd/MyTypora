## ADDED Requirements

### Requirement: Responsive Preview Editing Geometry
系统 SHALL 让正文和编辑区域随可用工作区展开，并保持渲染编辑的位置与选区稳定。

#### Scenario: 放大窗口与切换视图
- **WHEN** 用户放大窗口或切换源码、预览、专注模式
- **THEN** 内容使用可用工作区宽度并保留边距，不再受固定窄宽度限制

#### Scenario: 容器变化时继续编辑
- **WHEN** 编辑中的容器尺寸发生变化，即使没有窗口 resize 事件
- **THEN** 编辑框与正文锚点重新对齐，保留输入内容、焦点、选区及可见光标

#### Scenario: 扩展选区与滚轮
- **WHEN** 用户在长编辑内容中扩展选区或滚动鼠标
- **THEN** 滚动跟随选区活动端，一次普通滚轮只滚动一次，Ctrl 缩放不混入普通滚动
