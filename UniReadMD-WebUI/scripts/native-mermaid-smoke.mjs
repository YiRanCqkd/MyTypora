import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { once } from 'node:events';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { showWorkspaceControls } from './native-workspace-controls.mjs';

const require = createRequire(import.meta.url);
const { _electron: electron, chromium, expect } =
  require('../frontend/node_modules/@playwright/test');
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-mermaid-'));
const evidence = process.env.UNIREADMD_SMOKE_EVIDENCE;
const executablePath = process.env.UNIREADMD_SMOKE_EXECUTABLE || require('electron');
const sourceFile = path.join(directory, '中文节点图表.md');
const exportedFile = path.join(directory, '节点图表.html');
const labels = ['开始 Start', '处理 Process', '判断 Decision', '结束 End',
  '架构分组', '下一步 Next'];
const source = '```mermaid\nflowchart LR\nsubgraph 架构分组\n'
  + 'A[开始 Start] -->|下一步 Next| B[处理 Process]\n'
  + 'B --> C{判断 Decision}\nC --> D[结束 End]\nend\n```';
const checks = [];
let application;
let browser;

async function expectLabels(page, expected) {
  const svg = page.locator('.mermaid-output svg');
  await expect(svg).toBeVisible();
  for (const label of expected) {
    const text = svg.locator('text').filter({ hasText: label }).first();
    await expect(text).toBeVisible();
    const bounds = await text.boundingBox();
    assert.ok(bounds.width > 0 && bounds.height > 0);
  }
  await expect(svg.locator('foreignObject, script, [onload]')).toHaveCount(0);
}

try {
  await fs.writeFile(sourceFile, source, 'utf8');
  const environment = { ...process.env, UNIREADMD_TEST_MODE: '1',
    UNIREADMD_TEST_USER_DATA: path.join(directory, 'profile') };
  delete environment.ELECTRON_RUN_AS_NODE;
  delete environment.ELECTRON_START_URL;
  application = await electron.launch({ executablePath, env: environment, timeout: 60_000,
    args: process.env.UNIREADMD_SMOKE_EXECUTABLE ? [] : [path.join(root, 'electron/main.cjs')] });
  const page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  // 测试模式默认隐藏窗口，截图前显示但不抢占用户焦点，保证原生合成器能输出画面。
  if (evidence) await application.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0].showInactive();
  });
  await expect(page.locator('.preview')).toBeVisible();
  await showWorkspaceControls(application, page);
  await application.evaluate(({ dialog, Menu }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    Menu.getApplicationMenu().items[0].submenu.items.find((item) => item.label === '打开...')
      .click();
  }, sourceFile);
  await expectLabels(page, labels);
  if (evidence) await page.screenshot({ path: path.join(evidence, 'native-mermaid-light.png') });
  checks.push('真实中文 Markdown 文件中的流程图节点、连线及分组文字可见');

  await page.getByText('偏好设置', { exact: true }).click();
  await page.getByRole('combobox', { name: '主题', exact: true }).selectOption('dark');
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await expectLabels(page, labels);
  if (evidence) await page.screenshot({ path: path.join(evidence, 'native-mermaid-dark.png') });
  checks.push('真实安装产物主题切换后节点文字仍可见');

  await page.locator('.mermaid-output .node text').filter({ hasText: '开始 Start' }).click();
  const input = page.locator('.preview-inline-editor-input');
  await expect(input).toBeVisible();
  await input.fill((await input.inputValue()).replace('开始 Start', '修改起点 Updated'));
  await input.press('Control+Enter');
  await expectLabels(page, ['修改起点 Updated', ...labels.slice(1)]);
  await page.keyboard.press('Control+s');
  await expect.poll(() => fs.readFile(sourceFile, 'utf8')).toContain('修改起点 Updated');
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath });
  }, exportedFile);
  await page.getByRole('combobox', { name: '导出格式' }).selectOption('html');
  await page.getByRole('button', { name: '导出', exact: true }).click();
  await expect.poll(() => fs.readFile(exportedFile, 'utf8').catch(() => ''))
    .toContain('修改起点 Updated');
  browser = await chromium.launch();
  const exportedPage = await browser.newPage({ viewport: { width: 1400, height: 950 } });
  await exportedPage.goto(pathToFileURL(exportedFile).href);
  await expectLabels(exportedPage, ['修改起点 Updated', ...labels.slice(1)]);
  if (evidence) {
    await exportedPage.screenshot({ path: path.join(evidence, 'native-mermaid-export.png') });
    await fs.copyFile(exportedFile, path.join(evidence, 'native-mermaid-export.html'));
  }
  checks.push('图表编辑保存及真实 HTML 导出，离线重新打开后节点文字可见');
} finally {
  await browser?.close();
  if (application) {
    const child = application.process();
    const exited = child.exitCode !== null ? Promise.resolve() : once(child, 'exit');
    await application.evaluate(({ app }) => app.exit(0)).catch(() => {});
    await exited;
  }
  const target = await fs.realpath(directory);
  assert.equal(path.dirname(target), await fs.realpath(os.tmpdir()));
  await fs.rm(target, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
}

const result = { ok: true, executablePath, checks };
if (evidence) await fs.writeFile(path.join(evidence, 'native-mermaid-result.json'),
  JSON.stringify(result, null, 2) + '\r\n', 'utf8');
console.log(JSON.stringify(result, null, 2));
