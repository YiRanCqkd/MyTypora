# Design

## Context
UniReadMD 当前将 Markdown 渲染、编辑、会话、工作区、设置和导出大量集中在
`frontend/src/App.vue`，Electron 文件、导出、菜单和窗口逻辑集中在 `electron/main.cjs`。
发布硬化需要先建立测试保护，再拆分模块，最后收紧安全和发布门禁。

Task 1 已固定本地工具链解析方式：Node 和 npm 来自 `OpenJS.NodeJS` 24.8.0 的同一安装根，
Git 仅使用 `D:\Program Files\Git\cmd\git.exe`，OpenSpec 仅通过固定 CLI 路径调用。

## Goals
- 安全打开不可信 Markdown，不让 renderer 获得任意本地文件能力。
- 所有可能丢稿的文档切换都必须有可验证的保护和恢复语义。
- 恢复 Typora 式核心编辑体验，但 CodeMirror Markdown 始终是唯一数据源。
- 导出结果与预览样式、资源、公式和图表保持一致。
- 发布门禁覆盖质量、依赖、性能、真实 Electron、安装包和最终证据哈希。

## Non-Goals
- 不实现云同步、多人协作、Git 集成、AI、插件市场、PWA 或完整国际化。
- 不更换富文本内核，不建立第二套文档模型。
- 不通过降低安全或测试阈值换取通过。

## Current Observations
- `App.vue` 当前启用 markdown-it 原始 HTML，并通过 `v-html` 渲染预览。
- `main.cjs` 当前注册文件、设置、导出和菜单 IPC，并在主窗口关闭了部分安全选项。
- `preload.cjs` 暴露旧桥接入口，后续必须兼容这些入口的名称和返回结构。
- `App.vue` 已有 CodeMirror、预览内联编辑、表格编辑、Focus 和 Typewriter 相关代码。
- 导出入口已包含 HTML、PDF、PNG、JPEG、DOCX、ODT、RTF 和 Pandoc 能力探测。
- 当前质量脚本存在仓库根与前端目录混用问题，发布门禁尚未覆盖最终验收。

## Architecture Decisions
### Security Boundary
- DOMPurify 固定 HTML 白名单、属性白名单和协议白名单。
- Mermaid 使用 strict security level，输出 SVG 再净化，禁止 `foreignObject`。
- CSP 固定为最小应用 origin、data/blob 图片、内联样式和无外部连接能力。
- BrowserWindow 固定 `contextIsolation:true`、`nodeIntegration:false`、`sandbox:true` 和
  `webSecurity:true`。
- `will-navigate` 和 `setWindowOpenHandler` 拒绝非应用 origin；合法外链交给系统浏览器。
- IPC 同时校验 `event.sender.id`、`event.senderFrame.url` 和启动时登记的应用 origin。
- 主进程按 `webContents.id` 维护 canonical realpath 授权集合，窗口关闭立即清理。

### Document Integrity
- 每次打开、新建、切换、关闭、外部启动和第二实例接管前检查 dirty 状态。
- Save/Discard/Cancel 必须稳定返回，并记录非敏感 trace id。
- 会话快照限制为 5 MiB；文档读写和导出分别执行计划中的 20 MiB 边界。
- 自动保存默认关闭，启用后使用原子写入和失败恢复提示。

### Workspace and Typora-Style Editing
- CodeMirror Markdown 是唯一真实文档；预览原地块编辑只创建临时控件。
- 原地块编辑提交时转换为 CodeMirror transaction，撤销栈必须与源码编辑一致。
- 块范围由 Markdown 源映射和 `data-source-start/end/token` 生成，不引入富文本 AST 内核。
- Task 8 只恢复段落、标题、列表、引用、代码块和表格等计划内块编辑范围。
- 文件夹树扫描固定深度、数量、扩展名和 symlink 策略；设置入口走受限 IPC。
- 用户 CSS 由主进程读取为文本，PostCSS 净化后返回内容，不向 renderer 暴露路径。

### Export Pipeline
- 导出生成器复用净化后的预览 HTML 和应用 CSS 清单，不维护第二套 Markdown 解析逻辑。
- HTML、PDF、PNG、JPEG 由内建 Chromium lane 处理，并等待字体、图片、Math 和 Mermaid。
- DOCX、ODT、RTF 只在固定 Pandoc 3.10 lane 可用时启用。
- 导出目标路径只来自对话框，格式和扩展名必须匹配。
- 超限、缺图片、缺 Pandoc、取消对话框和写入失败都返回固定错误码。

### Release Quality
- `resolve-toolchain.ps1` 是后续 Node/npm/Git 调用入口，禁止 PATH shim 和 Electron 内置 Node。
- 任务 4 创建 `invoke-node-tool.ps1` 后，Task 2-13 和 F1-F4 统一通过 helper 调用工具。
- 质量门禁依次覆盖 typecheck、Lint、coverage、单测、Electron 测试、build、audit、E2E、
  perf 和 package smoke。
- Final verifier 只读取既有产物和 manifest 的 Green 命令，不重放历史 Red 命令。

## Compatibility
- 旧 IPC channel、preload 方法、菜单命令、最近文件和导出返回字段必须保持兼容。
- 保存的设置 JSON 必须兼容旧 key，并对非法 key 给出稳定错误码。
- 依赖升级不能改变既有 Markdown 渲染、KaTeX、Mermaid 和导出合同。
- CI 不写入本机 D 盘路径；只记录 runner 工具链到环境证据。

## Logging and Evidence
- 应用日志只记录 trace id、阶段、错误码、格式、字节数和非敏感路径摘要。
- 证据 JSON 记录命令、退出码、断言、fixture、时间和文件哈希。
- 不记录完整环境变量、token、auth、Cookie、证书密码或用户秘密。

## Rollback Plan
- 每个 Todo 单独提交，回滚只撤销该 Todo 文件和测试。
- 发布关键任务回滚后，删除或作废所有下游证据哈希和 release ready 产物。
- 回滚态只能用于开发恢复，不能形成安装包或发布声明。

## Verification Plan
- 每个 Todo 保存 Red、Green、happy QA、failure QA、environment、manifest 和 sha256。
- Task 1 使用 OpenSpec validate/show、resolver baseline 篡改、base SHA 祖先和编码换行验证。
- Task 2-13 逐步引入特征测试、实现、依赖、E2E、真实 Electron、性能和最终审核。
- F1-F4 全部自动批准后才允许 FA 写入 `release-ready.json`。
