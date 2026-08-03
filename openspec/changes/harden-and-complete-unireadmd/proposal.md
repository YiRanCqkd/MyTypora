# harden-and-complete-unireadmd

## Why
UniReadMD 已具备 Markdown 编辑、预览、文件和导出骨架，但不可信 Markdown、renderer 文件权限、
未保存内容、导出一致性和发布质量门禁仍存在发布阻断风险。

本变更把这些风险拆成可验证的 Task 2-13，并以 OpenSpec、固定工具链和严格证据集作为后续
实施的入口门禁。

## What Changes
- C1：封闭 Markdown、导航、preload IPC 与本地文件系统之间的攻击链。
- C2：补齐脏文档保护、会话恢复、自动保存和大文件边界。
- C3：恢复渲染态内联编辑、Focus、Typewriter、快捷键和撤销流程。
- C4：补齐新建文件、递归文件夹树、排序、设置和安全用户 CSS。
- C5：重建统一导出管线，保留样式、资源、公式、图表和长文图片。
- C6：拆分超大模块，收紧依赖、CI、覆盖率、性能、安装包和最终发布门禁。

## Scope
### In
- OpenSpec change、五个 capability delta、Task 1 工具链 baseline 和 resolver。
- Markdown 净化、CSP、Electron 安全选项、导航、外链、IPC sender 和路径授权。
- 文档状态、脏切换保护、会话、自动保存、大文件限制和稳定错误码。
- CodeMirror Markdown 唯一数据源下的 Typora 式原地块编辑、文件工作区和设置入口。
- HTML、PDF、PNG、JPEG、DOCX、ODT、RTF 导出与可选 Pandoc lane。
- 模块拆分、依赖升级、测试、覆盖率、真实 Electron、安装包、性能和最终审核。

### Out
- 云同步、多人协作、Git 版本管理、AI 写作助手、插件市场、PWA 和完整国际化。
- 与发布阻断修复无关的品牌重设计、后端服务或移动端专项建设。
- 通过放宽 CSP、Electron 安全选项、路径授权、Lint、覆盖率或快照阈值换取通过。

## Impact
- Affected specs: `secure-markdown`、`document-integrity`、`document-workspace`、
  `document-export`、`release-quality`。
- Affected code: `UniReadMD-WebUI/frontend/src/App.vue`、
  `UniReadMD-WebUI/electron/main.cjs`、`UniReadMD-WebUI/electron/preload.cjs`、
  `UniReadMD-WebUI/frontend/src/services/bridge.ts`、质量脚本、CI 和 release 产物。
- Affected tests: Vitest、Playwright、Node `node:test`、真实 Electron、审计、覆盖率、性能、
  package smoke 和 OpenSpec validate。

## Acceptance Criteria
- AC1（映射任务 2、4）：`App.vue` 拆分后只负责装配，前端质量命令、类型检查、Lint 和
  已有单测按固定 helper 通过。
- AC2（映射任务 3、4）：`main.cjs` 拆分后只负责装配，IPC 合同测试和 Electron 原生测试通过。
- AC3（映射任务 5）：根与前端依赖升级到固定版本，四份 audit JSON 的 critical/high 均为 0。
- AC4（映射任务 6）：恶意 Markdown、远端导航、伪造 sender 和未授权路径均返回稳定错误码。
- AC5（映射任务 7）：所有破坏性文档切换都有 Save/Discard/Cancel 守卫，失败会话可恢复。
- AC6（映射任务 8）：预览原地块编辑、表格编辑、Focus 和 Typewriter 均以 CodeMirror
  Markdown 为唯一数据源，不引入新的富文本内核。
- AC7（映射任务 9）：新建文件、递归文件夹树、排序、设置和安全用户 CSS 有完整 E2E 覆盖。
- AC8（映射任务 10）：HTML、PDF、PNG、JPEG、DOCX、ODT、RTF 导出保留样式和资源，
  超限或缺能力时返回固定错误码。
- AC9（映射任务 11）：设计令牌、可访问性、键盘流和视觉基线通过自动化检查。
- AC10（映射任务 12）：源码和 unpacked Electron 的打开、保存、设置、导出和安装包冒烟全部通过。
- AC11（映射任务 13）：性能、包体、覆盖率、audit、E2E 和 package smoke 进入统一发布门禁。
- AC12（映射任务 13）：scope-map、changed-files、F1-F4 和 FA 只能在全部证据哈希通过后批准。

## Rollback
- Task 1-4 属于开发基础回滚；回滚时删除对应 OpenSpec change 或基础脚本，不删除已有规范。
- Task 5-13 属于发布关键回滚；任一回滚必须作废其下游任务和最终证据。
- 回滚后不得保留 `gates-ready.json`、`release-ready.json` 或已失效的安装包发布资格。
- 恢复发布资格必须重新实施被回滚任务，并重跑全部下游证据和最终审核。

## Safety and Evidence
- 不写入 token、auth、完整环境变量、Cookie 或秘密。
- 每个任务证据目录必须且只能包含七个规定文件，`sha256.txt` 覆盖其余六个文件。
- Red 阶段只记录历史失败命令与哈希；Green 阶段只记录幂等验证命令。
- 日志必须携带 trace id、稳定错误码、阶段、目标文件类型和必要的非敏感上下文。
