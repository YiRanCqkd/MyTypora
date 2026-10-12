import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { once } from 'node:events';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const require = createRequire(import.meta.url);
const playwright = require('../frontend/node_modules/@playwright/test');
const { _electron: electron, chromium } = playwright;
const expect = playwright.expect.configure({ timeout: 20_000 });
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-visual-editing-'));
const evidence = process.env.UNIREADMD_SMOKE_EVIDENCE;
const executablePath = process.env.UNIREADMD_SMOKE_EXECUTABLE || require('electron');
const checks = [];
let application;
let browser;
async function stop() {
  if (!application) return;
  const child = application.process();
  const exited = child.exitCode !== null ? Promise.resolve() : once(child, 'exit');
  await application.evaluate(({ app }) => app.exit(0)).catch(() => {});
  await exited;
  application = undefined;
}
async function save(page) {
  // 使用原生菜单验证真实保存链路，避免桌面焦点影响系统加速键的派发。
  await application.evaluate(({ Menu }) => Menu.getApplicationMenu().items[0].submenu.items
    .find((item) => item.label === '保存').click());
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  await expect(page.locator('.error-banner')).toHaveCount(0);
}
try {
  const environment = { ...process.env, UNIREADMD_TEST_MODE: '1',
    UNIREADMD_TEST_USER_DATA: path.join(directory, 'profile') };
  delete environment.ELECTRON_RUN_AS_NODE;
  delete environment.ELECTRON_START_URL;
  const launch = () => electron.launch({ executablePath, env: environment, timeout: 60_000,
    args: process.env.UNIREADMD_SMOKE_EXECUTABLE ? [] : [path.join(root, 'electron/main.cjs')] });
  const source = path.join(directory, '原位编辑.md');
  await fs.writeFile(source, '# 原位标题\n\n正文 **加粗** 结尾\n\n- 第一项\n\n> 引用', 'utf8');
  application = await launch();
  let page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await expect(page.locator('.preview')).toBeVisible();
  // 原生菜单加速键需要活动窗口，不能依赖测试模式下隐藏窗口的偶然焦点状态。
  await application.evaluate(({ BrowserWindow }) => {
    const window = BrowserWindow.getAllWindows()[0];
    window.show();
    window.focus();
    window.webContents.focus();
  });
  await application.evaluate(({ dialog, BrowserWindow }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    BrowserWindow.getAllWindows()[0].webContents.send('menu:command', 'file-open');
  }, source);
  const paragraph = page.locator('.preview > p');
  await expect(paragraph).toHaveText('正文 加粗 结尾');
  await paragraph.click();
  await expect(paragraph).toHaveAttribute('contenteditable', 'true');
  await expect(paragraph.locator('strong')).toHaveText('加粗');
  await paragraph.press('End');
  await page.keyboard.insertText(' 中文续写');
  await expect(paragraph).toContainText('中文续写');
  await page.waitForTimeout(150);
  await expect(paragraph).toBeFocused();
  await save(page);
  await expect.poll(() => fs.readFile(source, 'utf8')).toContain('**加粗** 结尾 中文续写');
  checks.push('真实文档正文保持格式原位编辑，输入后光标稳定并保存到磁盘');

  await page.locator('.preview h1').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.insertText('标题下的新段落');
  await save(page);
  await expect.poll(() => fs.readFile(source, 'utf8'))
    .toContain('# 原位标题\n\n标题下的新段落\n\n正文');
  await page.locator('.preview li').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.insertText('第二项');
  await save(page);
  await expect.poll(() => fs.readFile(source, 'utf8')).toContain('- 第一项\n- 第二项');
  checks.push('标题换段及简单列表续项直接输入，保存后保持 Markdown 结构');

  const quote = page.locator('.preview blockquote p');
  await quote.click();
  await quote.evaluate((node) => {
    node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    node.textContent = '中文组合输入';
    node.dispatchEvent(new InputEvent('input', { bubbles: true, isComposing: true }));
  });
  await page.waitForTimeout(150);
  await expect(quote).toBeFocused();
  assert.doesNotMatch(await fs.readFile(source, 'utf8'), /中文组合输入/);
  await quote.evaluate((node) => node.dispatchEvent(new CompositionEvent('compositionend',
    { bubbles: true, data: '中文组合输入' })));
  await quote.press('Control+a');
  await quote.press('Control+b');
  await expect(quote.locator('b, strong')).toHaveText('中文组合输入');
  await save(page);
  await expect.poll(() => fs.readFile(source, 'utf8')).toContain('> **中文组合输入**');
  checks.push('中文组合事件保留编辑 DOM，结束输入后格式快捷键与保存一致');

  await page.locator('.preview li').last().click();
  await page.keyboard.press('End');
  await page.keyboard.insertText('可撤销文字');
  await page.keyboard.press('Control+z');
  await expect(page.locator('.preview li').last()).not.toContainText('可撤销文字');
  await page.keyboard.press('Control+Shift+z');
  await expect(page.locator('.preview li').last()).toContainText('可撤销文字');
  await save(page);
  if (evidence) {
    await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]
      .showInactive());
    await page.locator('.preview li').last().click();
    await page.screenshot({ path: path.join(evidence, 'native-visual-editing.png') });
    await page.keyboard.press('Control+Enter');
  }
  checks.push('原位编辑撤销重做沿用文档历史，保存后内容一致');

  const exportedFile = path.join(directory, '原位编辑.html');
  await application.evaluate(({ dialog, BrowserWindow }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath });
    BrowserWindow.getAllWindows()[0].webContents.send('menu:command', 'file-export-html');
  }, exportedFile);
  await expect.poll(() => fs.stat(exportedFile).then((stat) => stat.size).catch(() => 0))
    .toBeGreaterThan(1000);
  browser = await chromium.launch();
  const exported = await browser.newPage();
  await exported.goto(pathToFileURL(exportedFile).href);
  await expect(exported.locator('blockquote strong')).toHaveText('中文组合输入');
  await expect(exported.locator('[contenteditable="true"], .preview-visual-editing'))
    .toHaveCount(0);
  checks.push('真实 IPC 导出后离线展示格式正确，活动编辑属性未进入导出内容');

  await stop();
  application = await launch();
  page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await expect(page.locator('.preview')).toContainText('标题下的新段落');
  await expect(page.locator('.preview blockquote strong')).toHaveText('中文组合输入');
  await expect(page.locator('.preview li').last()).toContainText('可撤销文字');
  checks.push('独立应用重启恢复已保存文档，正文、引用格式与列表内容保留');
} finally {
  await browser?.close();
  await stop();
  const target = await fs.realpath(directory);
  assert.equal(path.dirname(target), await fs.realpath(os.tmpdir()));
  assert.ok(path.basename(target).startsWith('unireadmd-visual-editing-'));
  await fs.rm(target, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
}
const result = { ok: true, executablePath, checks };
if (evidence) await fs.writeFile(path.join(evidence, 'native-visual-editing-result.json'),
  JSON.stringify(result, null, 2) + '\r\n', 'utf8');
console.log(JSON.stringify(result, null, 2));
