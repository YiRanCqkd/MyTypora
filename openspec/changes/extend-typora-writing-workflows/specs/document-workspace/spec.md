## ADDED Requirements

### Requirement: Inline Fragment Editing
系统 SHALL 在编辑常见行内格式时只展开对应源码片段，保留相邻正文的渲染和源码。

#### Scenario: 编辑加粗片段
- **WHEN** 用户点击段落中的加粗文本并输入
- **THEN** 仅该片段展开编辑，原段落其他格式和文本保持，取消及撤销正确恢复

### Requirement: Selection Statistics
系统 SHALL 统计当前有效选区的字数与字符数，并提供统计单位和阅读速度配置。

#### Scenario: 源码与预览选区
- **WHEN** 用户在源码或预览选择文本
- **THEN** 状态栏显示选区统计，清除选区后恢复全文统计，文档切换不残留选区

### Requirement: Academic Math Compatibility
系统 SHALL 支持文档内公式编号和标签引用，并支持化学表达式。

#### Scenario: 前向引用与编号
- **WHEN** 文档在公式前引用其标签，并启用自动编号
- **THEN** 引用和对应公式显示一致编号，修改或切换文档后重新计算

### Requirement: Authorized Image Collection
系统 SHALL 提供本地图片批量整理，保留原图片并保护代码示例与用户已有文件。

#### Scenario: 整理包含引用式图片的文档
- **WHEN** 用户选择已授权工作区内的整理目录
- **THEN** 复制实际图片、更新 Markdown 引用，失败不覆盖正文或既有文件

### Requirement: Extended Export Options
系统 SHALL 提供固定格式的扩展导出及受限选项，保持既有导出兼容。

#### Scenario: 转换失败
- **WHEN** 扩展导出失败
- **THEN** 显示失败原因，保留原输出文件并清理临时文件

### Requirement: Writing Theme Profiles
系统 SHALL 支持写作主题和明暗独立选择，并保持自定义 CSS 能力。

#### Scenario: 系统主题变化
- **WHEN** 用户启用跟随系统且系统配色改变
- **THEN** 应用对应明暗写作主题，重新启动后保持选择

### Requirement: Bounded Document Rendering
系统 SHALL 缓存符合条件的大文档正文渲染，并延后屏幕外图表渲染。

#### Scenario: 导出屏幕外图表
- **WHEN** 用户导出包含尚未显示图表的文档
- **THEN** 导出等待全部图表完成，交互加载的延后策略不造成图表缺失
