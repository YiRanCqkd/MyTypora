## ADDED Requirements
### Requirement: Dirty Document Guard
系统 MUST 在打开、新建、关闭、外部启动和第二实例接管前保护未保存 Markdown。

#### Scenario: 脏文档切换需要用户决策
- **WHEN** 当前文档存在未保存修改且用户触发破坏性切换
- **THEN** 应用提供 Save、Discard、Cancel 三种路径并按选择继续或中止

### Requirement: Session and Autosave Recovery
系统 MUST 对会话、自动保存、大文件和写入失败提供可恢复语义。

#### Scenario: 自动保存失败可恢复
- **WHEN** 自动保存启用但原子写入失败
- **THEN** 应用保留内存内容、报告稳定错误码，并允许用户手动恢复或另存

## MODIFIED Requirements
### Requirement: Document State Baseline
当前应用 MUST 维护编辑内容、文件路径、最近文件、会话快照和脏状态提示，并对所有状态切换
执行明确失败语义。

#### Scenario: 打开文档前保护当前状态
- **WHEN** 用户通过打开文件流程载入 Markdown 且当前文档未保存
- **THEN** 应用先完成脏状态决策，再更新内容、文件路径和最近文件状态

### Requirement: Save Flow Baseline
当前应用 MUST 支持手动保存、另存为、自动保存和原子写入，并返回旧合同字段与稳定错误码。

#### Scenario: 保存失败不丢内存稿
- **WHEN** 主进程保存失败、目标超限或路径未授权
- **THEN** 应用保留当前 Markdown 内容，返回错误码并允许用户重试或另存
