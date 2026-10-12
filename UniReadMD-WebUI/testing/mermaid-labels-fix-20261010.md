# Mermaid 节点文字缺失修复记录

日期：2026-10-10。

## 问题与原因

使用 `flowchart TD` 和中英文节点复现：矩形、菱形、连线文字与分组标题正常，
但四个节点的文字均为空。上一版截图和 SVG 保存在 `release/mermaid-labels-20261010`。

本地安装的 Mermaid 11.16.1 部分节点直接读取全局 `htmlLabels`；应用此前只设置
`flowchart.htmlLabels: false`，导致节点仍使用 foreignObject 内的 HTML 文字。
应用的安全过滤会移除 foreignObject，因此出现有节点形状但无文字的现象。
[Mermaid 官方配置说明](https://mermaid.js.org/config/schema-docs/config-properties-htmllabels.html)
也明确要求使用全局 htmlLabels，图表专用设置已弃用。

回归检查同时发现：源码块定位只接受 HTMLElement，点击 SVG 节点文字无法进入编辑。
最初测试点击矩形中心被新显示的文字挡住，改为用户实际点击文字后，复现了这一遗漏。

## 修改

- 统一在 Mermaid 初始化适配层设置全局 `htmlLabels: false`，使用 SVG text 输出节点文字。
- 保留库的全部默认 secure 配置，并锁定 htmlLabels，文档指令及配置头不会改回 HTML 标签。
- 保留原有 strict、安全过滤、CSP 和本地文件授权。
- 点击映射接受 SVG Element，最终映射必须仍是外层 HTML 源码块。

这是恢复既有图表显示和源码块编辑行为的缺陷修复，按 OpenSpec 的 bug fix 例外直接实施。

## 验证

已保存单元及浏览器 Red 证据。新增检查覆盖中英文节点、连线、分组、浅深主题、
文档内 init 与配置头、新节点形状、状态图、类图、时序图、思维导图、点击编辑与 HTML 导出。
类型、Lint、生产构建和包体积检查通过。
新增 Mermaid 专项 5 项已通过，包括多行、加粗标签与实际点击 SVG 文字编辑。
核心过滤及点击映射定向单元 11 项通过，类型、Lint、生产构建及包体积检查通过。
测试初轮的主题选择器使用精确 label 匹配，与实际控件命名不符；改用现有用例中的
combobox 角色定位后通过，未改变产品或断言、未延长超时。各轮日志保留供复核。
完整浏览器回归 88 项通过，五张既有视觉基线未更新、未放宽比较容差。
主进程通配回归初轮 70 项中有 2 项打包一致性失败，因为未指定本轮产物目录，
默认读取了早期的 `release/electron/win-unpacked`；本轮独立打包检查 3 项已通过。
保留失败日志，指定 `UNIREADMD_PACK_DIR` 为本轮目录后重跑原有全部断言。
重跑全部主进程及打包检查 70/70 通过；最终前端 29 个文件、160 项单元全部通过。
核心覆盖率为语句 96.25%、分支 92.61%、函数 100%、行 96.21%，原有 90% 门槛未改变。
最终真实安装产物原生 11 项、剪贴板 3 项、写作 9 项、Mermaid 3 项，共 26 项通过。
Mermaid 检查读取中文文件、切换深色主题、点击节点文字编辑并保存，再通过真实 IPC
导出 HTML，使用独立 Chromium 离线打开导出文件，检查 SVG 文字可见且具有实际尺寸。
浅深主题与离线导出截图已审看；外部对象及事件仍被过滤。
本轮真实 6,575,016 字节文档打开 2599ms、搜索 139ms、编辑 470ms，原生门槛通过。

## 发布

独立产物目录：`release/nsis-20261010-mermaid-labels`。
构建新的 NSIS，并在临时目录验证实际文件打开、主题切换、编辑保存、HTML 导出和离线打开。
原有正式版保持不覆盖。
安装包为 `UniReadMD-Setup-1.0.0-x64.exe`，109,706,169 字节。
SHA-256：`cdaa558d777ffec3c4989f7c025c1a42f38e767f411550d726cc0e2a69968485`。
7-Zip 完整性检查退出码为 0。
首轮安装验证已有 23 项通过；新增 Mermaid 检查确认文字存在后，在隐藏窗口截图超时。
按既有原生脚本的方式，在取图前 showInactive，保证窗口合成画面但不抢焦点。
此项只修改测试脚本，未延长 20 秒超时。首轮仍正常卸载并还原原安装，日志保存在
`validation-20261010/first-attempt`，随后重新完整验证同一安装包。
最终安装和卸载退出码均为 0，父验证脚本退出码为 0。
临时安装的 EXE 与 app.asar 与当前打包内容哈希一致。
原正式版 EXE、app.asar、全局安装记录、缓存及快捷方式已核对还原，系统剪贴板已还原。
最终证据见 `validation-20261010/verification.log`、`native-mermaid-result.json`、
`native-mermaid-light.png`、`native-mermaid-dark.png`、`native-mermaid-export.png`。

构建准备时误建了空目录 `frontend/release/mermaid-labels-20261010`。
自动审批检查拒绝了清理命令，仅返回 blocked by policy；只读确认目录为空，
保留该目录继续工作，不重试删除，不影响实际证据目录或发布产物。
