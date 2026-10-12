import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { once } from 'node:events';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const playwright = require('../frontend/node_modules/@playwright/test');
const { _electron: electron } = playwright;
const expect = playwright.expect.configure({ timeout: 20_000 });
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-rendering-layout-'));
const evidence = process.env.UNIREADMD_SMOKE_EVIDENCE;
const executablePath = process.env.UNIREADMD_SMOKE_EXECUTABLE || require('electron');
const checks = [];
let application;
async function stop(gracefully = false) {
  if (!application) return;
  const child = application.process();
  const exited = child.exitCode !== null ? Promise.resolve() : once(child, 'exit');
  if (gracefully) {
    // 重启验证走真实关闭流程，确保防抖快照完成；清理失败现场时才直接退出。
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
  const source = path.join(directory, '正文布局回归.md');
  const original = '前一段\n\n首行正文\n次行正文\n末行正文\n\n后一段';
  await fs.writeFile(source, original, 'utf8');
  application = await launch();
  let page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await expect(page.locator('.preview')).toBeVisible();
  await application.evaluate(({ dialog, BrowserWindow }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    BrowserWindow.getAllWindows()[0].webContents.send('menu:command', 'file-open');
  }, source);
  await expect(page.locator('.preview > p')).toHaveCount(3);
  const geometry = () => page.locator('.preview > p').evaluateAll((nodes) => nodes.map((node) => ({
    top: node.getBoundingClientRect().top, height: node.getBoundingClientRect().height,
  })));
  const before = await geometry();
  const paragraph = page.locator('.preview > p').nth(1);
  await paragraph.click();
  await expect(paragraph).toHaveAttribute('contenteditable', 'true');
  await paragraph.press('Home');
  const entered = await geometry();
  const shadow = await paragraph.evaluate((node) => getComputedStyle(node).boxShadow);
  if (evidence) await fs.writeFile(path.join(evidence, 'native-rendering-observed.json'),
    JSON.stringify({ before, entered, shadow }, null, 2) + '\r\n', 'utf8');
  await expect(paragraph).toHaveCSS('box-shadow', 'none');
  await expect(paragraph).toHaveCSS('border-left-width', '0px');
  assert.deepEqual(entered, before, '进入普通正文编辑不能增加空白或移动相邻段落');
  await paragraph.press('Control+End');
  await page.keyboard.insertText('新增');
  await page.waitForTimeout(150);
  assert.deepEqual(await geometry(), before, '普通输入不能改变原有行距和段距');
  if (evidence) {
    await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]
      .showInactive());
    await paragraph.click();
    await page.screenshot({ path: path.join(evidence, 'native-rendering-layout.png') });
  }
  checks.push('正文只有一个插入光标，多行普通正文进入编辑及输入时保持原有几何位置');

  // 布局专项通过原生菜单保存，键盘保存由表格与图片专项独立验证。
  await application.evaluate(({ Menu }) => Menu.getApplicationMenu().items[0].submenu.items
    .find((item) => item.label === '保存').click());
  const expected = original.replace('末行正文', '末行正文新增');
  await expect.poll(() => fs.readFile(source, 'utf8')).toBe(expected);
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  await expect(page.locator('.preview > p')).toHaveCount(3);
  assert.deepEqual(await geometry(), before, '保存和重新渲染不能增加空段落');
  checks.push('真实磁盘保存保留原段落和单换行，未增加前后空行');

  const sourceToggle = page.locator('#toggle-sourceview-btn');
  const footer = page.locator('.status-bar');
  const footerBefore = (await footer.boundingBox());
  const controlBefore = (await sourceToggle.boundingBox());
  await sourceToggle.check();
  await expect(page.locator('.editor-host')).toBeVisible();
  await page.waitForTimeout(250);
  assert.deepEqual(await footer.boundingBox(), footerBefore, '源码切换不能移动状态栏');
  assert.deepEqual(await sourceToggle.boundingBox(), controlBefore, '源码开关位置应稳定');
  await expect(page.locator('#outline-btn')).toBeVisible();
  await sourceToggle.uncheck();
  await expect(page.locator('.preview')).toBeVisible();
  checks.push('真实安装版源码开关与状态栏位置稳定，侧栏入口始终可见');

  await stop(true);
  application = await launch();
  page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await expect(page.locator('.preview > p')).toHaveCount(3);
  await expect(page.locator('.preview > p').nth(1)).toContainText('末行正文新增');
  assert.equal(await fs.readFile(source, 'utf8'), expected);
  checks.push('应用重启后恢复正文原有换行与段落结构');
} finally {
  await stop();
  // 只清理本脚本创建的临时配置及样本文档，不能涉及正式安装或用户文档。
  const target = await fs.realpath(directory);
  assert.equal(path.dirname(target), await fs.realpath(os.tmpdir()));
  assert.ok(path.basename(target).startsWith('unireadmd-rendering-layout-'));
  await fs.rm(target, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
}
const result = { ok: true, executablePath, checks };
if (evidence) await fs.writeFile(path.join(evidence, 'native-rendering-layout-result.json'),
  JSON.stringify(result, null, 2) + '\r\n', 'utf8');
console.log(JSON.stringify(result, null, 2));
