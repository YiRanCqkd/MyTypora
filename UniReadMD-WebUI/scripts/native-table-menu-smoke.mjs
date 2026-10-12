import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { once } from 'node:events';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const { _electron: electron, expect: baseExpect } =
  require('../frontend/node_modules/@playwright/test');
const expect = baseExpect.configure({ timeout: 20_000 });
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-table-menu-'));
const evidence = process.env.UNIREADMD_SMOKE_EVIDENCE;
const executablePath = process.env.UNIREADMD_SMOKE_EXECUTABLE || require('electron');
const checks = [];
let application;

async function stop(gracefully = false) {
  if (!application) return;
  const child = application.process();
  const exited = child.exitCode !== null ? Promise.resolve() : once(child, 'exit');
  if (gracefully) {
    await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close());
    let timer;
    try {
      await Promise.race([exited, new Promise((_, reject) => {
        timer = setTimeout(() => reject(new Error('正常关闭应用超时')), 20_000);
      })]);
    } finally { clearTimeout(timer); }
  } else {
    await application.evaluate(({ app }) => app.exit(0)).catch(() => {});
    await exited;
  }
  application = undefined;
}

try {
  const environment = { ...process.env, UNIREADMD_TEST_MODE: '1',
    UNIREADMD_TEST_USER_DATA: path.join(directory, 'profile') };
  delete environment.ELECTRON_RUN_AS_NODE;
  delete environment.ELECTRON_START_URL;
  const launch = () => electron.launch({ executablePath, env: environment, timeout: 60_000,
    args: process.env.UNIREADMD_SMOKE_EXECUTABLE ? [] : [path.join(root, 'electron/main.cjs')] });
  const source = path.join(directory, '表格原位编辑.md');
  const original = '# 表格回归\n\n| 名称 | 内容 |\n| :---: | ---: |\n| **加粗** | 原内容 |';
  await fs.writeFile(source, original, 'utf8');
  application = await launch();
  let page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  // 原生键盘保存需要真实焦点，测试窗口使用独立配置且不触碰用户窗口。
  await application.evaluate(({ BrowserWindow }) => {
    const window = BrowserWindow.getAllWindows()[0];
    window.show(); window.focus(); window.webContents.focus();
  });
  await expect(page.locator('.preview')).toBeVisible();
  await application.evaluate(({ dialog, BrowserWindow }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    BrowserWindow.getAllWindows()[0].webContents.send('menu:command', 'file-open');
  }, source);
  const first = page.locator('.preview td').first();
  const last = page.locator('.preview td').last();
  await expect(first.locator('strong')).toHaveText('加粗');
  const before = await first.boundingBox();
  await first.click();
  await expect(first).toHaveAttribute('contenteditable', 'true');
  await expect(page.getByTestId('preview-table-cell-editor')).toHaveCount(0);
  assert.deepEqual(await first.boundingBox(), before);
  await expect(first).toHaveCSS('text-align', 'center');
  await first.press('Tab');
  await expect(last).toBeFocused();
  checks.push('真实安装版表格直接编辑，保留加粗、尺寸、对齐及 Tab 导航');

  await last.fill('中文|内容');
  await last.click({ button: 'right' });
  const menu = page.getByTestId('preview-table-menu');
  await expect(menu).toBeVisible();
  await expect(menu).toHaveCSS('position', 'fixed');
  await expect(menu).toHaveCSS('font-size', '13px');
  await expect(menu).toHaveCSS('width', '224px');
  const menuBox = await menu.boundingBox();
  const hostBox = await page.getByTestId('editor-preview').boundingBox();
  assert.ok(menuBox.y >= hostBox.y && menuBox.y + menuBox.height <= hostBox.y + hostBox.height);
  if (evidence) await page.screenshot({ path: path.join(evidence, 'native-table-menu-light.png') });
  await menu.getByRole('button', { name: '列居中', exact: true }).click();
  await expect(last).toHaveText('中文|内容');
  await expect(last).toHaveCSS('text-align', 'center');
  await page.keyboard.press('Control+s');
  await expect.poll(() => fs.readFile(source, 'utf8')).toContain('中文\\|内容');
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  assert.match(await fs.readFile(source, 'utf8'), /\| :---: \| :---: \|/);
  checks.push('刚输入即右键不丢失文字，统一菜单定位正常，对齐与转义真实保存');

  await last.click();
  await last.press('Control+End');
  await last.press('Shift+Enter');
  await page.keyboard.insertText('次行');
  await expect(last).toContainText('次行');
  await last.press('Control+Enter');
  await expect(last.locator('br')).toHaveCount(1);
  await expect(page.locator('.cm-content')).toContainText('中文\\|内容<br>次行');
  await last.click();
  await expect(last).toHaveAttribute('contenteditable', 'true');
  await expect(page.getByTestId('preview-table-cell-editor')).toHaveCount(0);
  await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]
    .webContents.send('menu:command', 'file-save'));
  await expect.poll(() => fs.readFile(source, 'utf8')).toContain('中文\\|内容<br>次行');
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  checks.push('单元格内换行写回 br，保存后再次进入仍为原位编辑');

  await last.click();
  await last.press('Control+End');
  await page.keyboard.insertText('导出新增');
  const exported = path.join(directory, '表格.html');
  await application.evaluate(({ dialog, BrowserWindow }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath });
    BrowserWindow.getAllWindows()[0].webContents.send('menu:command', 'file-export-html');
  }, exported);
  await expect.poll(() => fs.readFile(exported, 'utf8').catch(() => ''))
    .toContain('导出新增');
  const { JSDOM } = require('jsdom');
  const exportedHtml = await fs.readFile(exported, 'utf8');
  const exportedDom = new JSDOM(exportedHtml.match(/<article\b[\s\S]*<\/article>/)?.[0]);
  assert.equal(exportedDom.window.document.querySelector(
    '[contenteditable], .preview-visual-editing',
  ), null);
  exportedDom.window.close();
  await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]
    .webContents.send('menu:command', 'file-save'));
  await expect(page.locator('.error-banner')).toHaveCount(0);
  await expect.poll(() => fs.readFile(source, 'utf8')).toContain('导出新增');
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  checks.push('活动单元格立即导出提交最新内容，离线 HTML 无编辑属性');

  await page.evaluate(() => { document.documentElement.dataset.theme = 'dark'; });
  await last.click({ button: 'right' });
  await expect(menu).toHaveCSS('background-color', 'rgb(38, 40, 45)');
  if (evidence) await page.screenshot({ path: path.join(evidence, 'native-table-menu-dark.png') });
  await page.keyboard.press('ArrowDown');
  await expect(menu.getByRole('button').nth(1)).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(menu).toHaveCount(0);
  checks.push('深色菜单可读，键盘导航与 Escape 关闭生效');

  await stop(true);
  application = await launch();
  page = await application.firstWindow();
  await expect(page.locator('.preview td').last()).toContainText('导出新增');
  await expect(page.locator('.preview td').first().locator('strong')).toHaveText('加粗');
  await expect(page.locator('.preview td').last().locator('br')).toHaveCount(1);
  checks.push('正常关闭及重启恢复表格内容、加粗和单元格换行');
} finally {
  await stop();
  const target = await fs.realpath(directory);
  assert.equal(path.dirname(target), await fs.realpath(os.tmpdir()));
  assert.ok(path.basename(target).startsWith('unireadmd-table-menu-'));
  await fs.rm(target, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
}
const result = { ok: true, executablePath, checks };
if (evidence) await fs.writeFile(path.join(evidence, 'native-table-menu-result.json'),
  JSON.stringify(result, null, 2) + '\r\n', 'utf8');
console.log(JSON.stringify(result, null, 2));
