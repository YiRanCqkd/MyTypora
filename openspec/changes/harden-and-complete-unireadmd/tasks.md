# Tasks

## 1. Frontend Decomposition and Baseline Gates
- [ ] 1.1 Task 2：用特征测试锁定前端现有渲染、编辑、会话、工作区、设置和导出行为。
- [ ] 1.2 Task 2：从 `App.vue` 提取 markdown、editor、workspace、settings 和 export 模块。
- [x] 1.3 Task 4：修复 CI、Lint、coverage、source-size 和 E2E baseline 基础设施。

## 2. Electron Decomposition and Contracts
- [ ] 2.1 Task 3：为设置、文件、授权、导出和菜单建立 Node `node:test` 合同测试。
- [ ] 2.2 Task 3：从 `main.cjs` 提取 authorization、files、exports、settings 和 menus 模块。
- [ ] 2.3 Task 3：保持现有 IPC channel、preload 方法和返回结构兼容。

## 3. Dependency and Security Hardening
- [x] 3.1 Task 5：升级并锁定 Electron、DOMPurify、MarkdownIt、Mermaid、Vite 和 Vitest 等依赖。
- [ ] 3.2 Task 6：固定 Markdown HTML 白名单、属性白名单、协议白名单和 Mermaid strict 策略。
- [ ] 3.3 Task 6：收紧 CSP、BrowserWindow 安全选项、导航、外链、sender 校验和路径授权。
- [ ] 3.4 Task 6：为 XSS、远端导航、伪造 sender、未授权路径和危险资源添加 Red/Green 证据。

## 4. Document Integrity
- [ ] 4.1 Task 7：为打开、新建、关闭、外部启动和第二实例接管增加 Save/Discard/Cancel 守卫。
- [ ] 4.2 Task 7：补齐会话快照、自动保存、原子写入、大文件限制和恢复提示。
- [ ] 4.3 Task 7：验证取消、保存失败、超限内容和损坏会话的稳定错误码。

## 5. Editor and Workspace Completion
- [ ] 5.1 Task 8：恢复 Typora 式原地块编辑、表格编辑、Focus、Typewriter、快捷键和撤销流程。
- [ ] 5.2 Task 8：确保所有预览编辑提交都写回 CodeMirror Markdown，不引入新的富文本内核。
- [ ] 5.3 Task 9：补齐新建文件、递归文件夹树、排序、视图设置、拼写和安全用户 CSS 设置。
- [ ] 5.4 Task 9：验证 symlink 环、无权限目录、危险 CSS、非 CSS 文件和脏文档新建失败路径。

## 6. Export Completion
- [ ] 6.1 Task 10：构建统一导出文档生成器，复用净化预览 HTML 和应用 CSS 清单。
- [ ] 6.2 Task 10：验证 HTML、PDF、PNG、JPEG 的样式、资源、公式、Mermaid 和长文尺寸。
- [ ] 6.3 Task 10：固定 Pandoc 3.10 lane，并验证 DOCX、ODT、RTF 可用和缺失场景。

## 7. Release Quality and Final Verification
- [ ] 7.1 Task 11：建立设计令牌，修复可访问性、键盘流、reduced-motion 和视觉基线。
- [ ] 7.2 Task 12：增加源码与 unpacked Electron 的真实原生流程和安装包冒烟验证。
- [ ] 7.3 Task 13：收紧性能、包体、覆盖率、audit、E2E、package smoke 和总发布门禁。
- [ ] 7.4 Task 13：生成 `scope-map.json`、`changed-files.json`、`gates-ready.json` 和最终审核输入。
- [ ] 7.5 F1-F4：执行计划合规、代码质量、真实 Electron QA 和范围一致性自动审核。
- [ ] 7.6 FA：仅在 F1-F4 全部通过后生成 `release-ready.json`。
