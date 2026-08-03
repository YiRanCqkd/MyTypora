# secure-markdown Specification

## Purpose
记录 UniReadMD 当前 Markdown 预览、preload bridge、窗口导航和本地文件入口的安全基线，供
后续任务在不破坏旧调用方的前提下逐步收紧。

## Requirements
### Requirement: Markdown Rendering and Native Bridge Baseline
当前应用 MUST 支持 Markdown 预览、受限 preload bridge、菜单命令和本地文件打开保存的
既有入口，并保持这些入口对旧调用方可用。

#### Scenario: 既有 Markdown 能预览
- **WHEN** 用户打开或输入 Markdown 文档
- **THEN** 应用展示预览内容并保留旧桥接入口

### Requirement: External Navigation Baseline
当前应用 MUST 通过 Electron BrowserWindow 加载前端入口，并允许菜单、preload 与外部链接流程
继续工作。

#### Scenario: 桌面窗口加载应用
- **WHEN** 应用启动主窗口
- **THEN** renderer 可通过 preload 调用注册的 IPC channel
