# Phase2 前端与 Electron 骨架

## 目录说明
- `frontend/`：Vue3 + TypeScript + Vite + Pinia + Vitest + ESLint
- `electron/`：Electron Main/Preload，负责窗口与受限 IPC

## 快速启动
1. 在 `UniReadMD-WebUI` 目录执行 `npm install`
2. 在 `UniReadMD-WebUI/frontend` 目录执行 `npm install`
3. 回到 `UniReadMD-WebUI` 目录执行 `npm run dev`

## 质量门禁
- 本地执行：`npm run quality:frontend`
- CI 工作流：`.github/workflows/frontend-quality.yml`

## 打包命令
- `npm run pack:win`：生成并同步 Windows NSIS 安装包到 `release/electron/UniReadMD-Setup.exe`
- `npm run pack:dir`：生成 Windows unpacked 目录（用于快速验证）
- `npm run dist:nsis`：生成带版本号的 Windows NSIS 安装包
- `npm run dist:win`：等同于 `dist:nsis`

产物目录：`release/electron/`

已验证：
- `npm run pack:win` 在当前环境用于生成安装器交付物

## 图标与启动
- Windows 全部可执行文件图标统一使用：`res/unireadmd.ico`
  - 主程序 exe（`win-unpacked/UniReadMD.exe`）
  - NSIS 安装器 exe
  - NSIS 卸载器 exe
- 可通过脚本按指定 exe 直接启动窗口：
  - 默认启动：`start_from_exe.bat`
  - 指定 exe 启动：`start_from_exe.bat "I:\\path\\to\\UniReadMD-WebUI\\release\\electron\\win-unpacked\\UniReadMD.exe"`

注意：
- `dist:*` 首次执行会下载 `nsis` 工具链，若网络受限会导致构建失败
- 可通过内网镜像或缓存方式解决（例如设置 `ELECTRON_BUILDER_BINARIES_MIRROR`）

## 签名配置
使用 Electron Builder 标准环境变量：
- `CSC_LINK`：证书路径或 base64 内容
- `CSC_KEY_PASSWORD`：证书密码

PowerShell 示例：
```powershell
$env:CSC_LINK='C:\\cert\\unireadmd.pfx'
$env:CSC_KEY_PASSWORD='your-password'
npm run dist:win
```

## 当前能力
- Electron 加载前端页面（开发/生产入口自动切换）
- 受限 IPC：
  - `app:getVersion`
  - `app:logEvent`
  - `file:openMarkdown`
  - `file:openFromPath`
  - `file:saveMarkdown`
  - `export:html`
  - `window:minimize`
  - `window:toggleMaximize`
  - `window:isMaximized`
  - `window:close`
  - `window:toggleDevTools`
  - `settings:getAll`
  - `settings:get`
  - `settings:set`
  - `settings:selectUserCssFile`
- 前端页面：Files / Outline / Search 侧栏骨架
- 基础 Markdown 渲染与文件打开保存/导出流程
- 主进程与渲染层 trace id 贯通日志（`userData/logs/runtime.log`）
