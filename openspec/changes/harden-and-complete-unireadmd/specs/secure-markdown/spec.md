## ADDED Requirements
### Requirement: Trusted Markdown Sanitization
系统 MUST 对所有 Markdown 预览 HTML、Mermaid SVG、图片资源和外链执行固定白名单净化。

#### Scenario: 恶意 Markdown 被阻断
- **WHEN** 文档包含 `script`、`iframe`、事件属性、危险协议、远端图片或 data SVG
- **THEN** renderer 不执行脚本、不加载危险资源，并记录稳定安全错误码

### Requirement: Window Navigation and IPC Isolation
系统 MUST 限制应用窗口 origin、外部导航、新窗口、权限请求和 preload IPC sender。

#### Scenario: 远端页面不能继承 preload
- **WHEN** Markdown 或外部输入尝试导航到非应用 origin 或伪造 IPC sender
- **THEN** 主进程拒绝请求并返回 `ERR_UNTRUSTED_SENDER` 或对应安全错误码

## MODIFIED Requirements
### Requirement: Markdown Rendering and Native Bridge Baseline
当前应用 MUST 支持 Markdown 预览、受限 preload bridge、菜单命令和本地文件打开保存的
既有入口，并在不可信内容、导航和 IPC 调用中强制执行安全边界。

#### Scenario: 净化后 Markdown 能预览
- **WHEN** 用户打开包含常用 Markdown、允许 HTML、公式、Mermaid 和本地授权图片的文档
- **THEN** 应用展示净化后的预览内容，保留旧桥接入口并阻断未授权本地文件访问

#### Scenario: 危险桥接调用被拒绝
- **WHEN** renderer 从未登记 origin、已销毁窗口或远端 frame 调用本地文件 IPC
- **THEN** 主进程拒绝调用、返回稳定错误码并写入非敏感 trace 日志

### Requirement: External Navigation Baseline
当前应用 MUST 通过 Electron BrowserWindow 加载前端入口，并只允许应用 origin 留在窗口内。

#### Scenario: 合法外链交给系统浏览器
- **WHEN** 用户点击 `https`、`mailto` 或 `tel` 外链
- **THEN** 应用拒绝窗口内导航，并只通过系统浏览器或系统处理器打开
