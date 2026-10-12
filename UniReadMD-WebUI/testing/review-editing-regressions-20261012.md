# 可视编辑审查问题修复验证

日期：2026-10-12。版本：1.1.0。

## 修复结果

| 审查问题 | 修复方式 | 验证内容 |
| --- | --- | --- |
| 段内换行退格误合并前一段 | 检查格式节点内的文本和换行；删除段内换行后跳过跨段逻辑 | 普通正文、加粗、斜体、行内代码、连续换行；真实段首仍可合并 |
| 表格右键使用冻结预览中的旧范围 | 提交原位编辑时换算目标范围，刷新后重新定位并解析单元格 | 待提交和已提交输入、多张相同表格、正文行数减少、表格间空行保留 |
| 合法表格点击抛出异常 | 同时检查 null 和 undefined；保留省略竖线的数据行 | 可直接编辑并保存，其他数据行不丢失；无法取得值时回退源码入口 |

本次属于既有功能缺陷修复，按 OpenSpec 的缺陷修复规则直接实施。

## 验证结果

- 类型检查通过：`node node_modules/vue-tsc/bin/vue-tsc.js --noEmit`。
- 代码检查通过：`node node_modules/eslint/bin/eslint.js --ext .ts,.vue src`。
- 全量单元测试通过：35 个文件、206 项测试。
- 浏览器回归通过：32 项，其中新增 11 项审查问题回归用例。
- 表格工具定向测试通过：22 项。
- 表格工具覆盖率：语句 92.85%、分支 93.10%、函数 100%、行 93.37%。
- 本次修改的 7 个源码及测试文件均通过 UTF-8 无 BOM、CRLF 换行检查。

浏览器回归执行命令（在 frontend 目录运行）：

```powershell
node node_modules/@playwright/test/cli.js test `
  review-editing-regressions.spec.ts paragraph-enter.spec.ts `
  table-menu-experience.spec.ts visual-editing-boundaries.spec.ts `
  --workers=2 --reporter=line
```

定向覆盖率报告位于 `frontend/coverage-review-editing/coverage-summary.json`。

## 交付范围

本次已更新源码、回归用例和前端构建产物，并重新生成 NSIS 安装包。

- 发布目录：`release/nsis-20261012-1.1.0-review-fixes/`。
- 安装包：`UniReadMD-Setup-1.1.0-x64.exe`。
- 大小：109,718,805 字节，约 104.6 MB。
- SHA256：`f1f38302afe46215d908fa2229c59fb9857ecc5ae3905bda840a50efee84497c`。
- 应用版本 1.1.0，Windows 文件版本 1.1.0.0。
- 安装包压缩内容完整性检查通过。
- 实际安装内容检查 3 项通过，前端、主进程与当前构建逐文件一致。
- 安装版回归 8 项通过，覆盖此次审查问题及代码块、正文单次回车、保存和重启。
- 测试副本安装、卸载完成；原有安装、缓存及快捷方式均已校验保护或还原。

发布目录中的 `artifact-manifest.json`、`verification-result.json` 和 `SHA256.txt`
记录最终产物身份及验证结果。所有安装版回归证据均晚于最终安装包生成时间。
