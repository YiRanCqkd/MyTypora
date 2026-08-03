# 项目上下文

## 目的
UniReadMD 是一个 Windows 桌面 Markdown 编辑器，目标是在 Electron 外壳中提供接近
Typora 的编辑、预览、文件工作区和导出体验。当前发布硬化工作的目标是让应用能够安全打开
不可信 Markdown，保护未保存内容，并通过可重放的本地和 CI 质量门禁。

## 技术栈
- Electron Main/Preload 提供窗口、菜单、文件、设置和导出能力。
- Vue 3、TypeScript、Vite、Pinia 和 CodeMirror 6 构成前端工作台。
- Markdown 渲染基于 markdown-it 及现有数学、目录、任务列表和 Mermaid 扩展。
- 测试组合为 Vitest、Playwright、Node `node:test` 和真实 Electron 冒烟。
- Task 1 工具链固定为 Node.js 24.8.0、npm CLI 11.6.0、Git 2.47.1.windows.2
  和 OpenSpec 0.12.0。

## 项目约定

### 代码风格
- 交互、Markdown 文档和注释使用简体中文。
- 文件使用 UTF-8 无 BOM 和 CRLF。
- TypeScript、JavaScript、Vue 和 PowerShell 变更遵守 2 空格缩进。
- 代码优先使用单引号，单行目标长度不超过 100 个字符。
- 多行对象、数组和参数结构保留尾随逗号。

### 架构模式
- Electron 主进程持有本地文件权限，renderer 不得自行声明路径授权。
- `App.vue` 和 `electron/main.cjs` 应逐步收敛为装配层。
- 前端能力按 markdown、editor、workspace、settings 和 export 分层。
- Electron 能力按 authorization、files、exports、settings 和 menus 分层。
- CodeMirror Markdown 文档是唯一数据源；不引入新的富文本内核。

### 测试策略
- 每个实施任务先保存 Red 证据，再保存 Green 证据。
- 单测、类型检查、Lint、E2E、Electron 测试、构建、审计和性能都进入发布门禁。
- 证据目录固定为七个文件，并用 `sha256.txt` 覆盖其余六个文件。

### Git 工作流
- 每个 Todo 使用一个可独立回滚的约定式提交。
- 本地 Git 固定为 `D:\Program Files\Git\cmd\git.exe`。
- 不提交与当前任务无关的既有工作区修改。

## 领域上下文
- 用户会打开本地 `.md` 和 `.markdown` 文件，并可能从文件夹树批量浏览文档。
- Markdown 内容可能不可信，不能假设 HTML、链接、图片或 Mermaid 内容安全。
- 导出需要保留预览样式、公式、图片和图表，并支持内建与可选 Pandoc 两条路径。
- 文件、设置和导出路径授权必须与窗口和 sender origin 绑定。

## 重要约束
- 不实现云同步、多人协作、Git 集成、AI、插件市场、PWA 或完整国际化。
- 不通过放宽 CSP、Electron 安全选项、路径授权或测试阈值换取通过。
- 不把 token、auth、完整环境变量、Cookie 或秘密写入文档、日志、证据或提交。
- Windows 本地工具链必须从注册表和固定绝对路径解析，不能依赖当前进程 PATH。

## 外部依赖
- OpenSpec CLI 通过固定 Node 调用
  `D:\DevTools\npm-global\node_modules\@fission-ai\openspec\bin\openspec.js`，安装缓存固定为
  `D:\DevTools\npm-cache`，不得调用 PATH shim 或 C 盘全局包。
- Pandoc 是可选外部能力，仅在导出测试 lane 中按固定版本 3.10 探测。
- CI 中 Git 使用 runner 自带版本并记录环境证据，不写入本机 D 盘路径。
