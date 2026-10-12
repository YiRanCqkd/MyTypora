import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { once } from 'node:events';
import { fileURLToPath, pathToFileURL } from 'node:url';

const require = createRequire(import.meta.url);
const playwright = require('../frontend/node_modules/@playwright/test');
const { _electron: electron, chromium } = playwright;
const expect = playwright.expect.configure({ timeout: 20_000 });
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-image-workspace-'));
const evidence = process.env.UNIREADMD_SMOKE_EVIDENCE;
const executablePath = process.env.UNIREADMD_SMOKE_EXECUTABLE || require('electron');
let application;
let browser;
const checks = [];

async function stop() {
  if (!application) return;
  const child = application.process();
  const exited = child.exitCode !== null ? Promise.resolve() : once(child, 'exit');
  await application.evaluate(({ app }) => app.exit(0)).catch(() => {});
  await exited;
  application = undefined;
}
async function clickView(id) {
  await application.evaluate(({ Menu }, id) => {
    Menu.getApplicationMenu().getMenuItemById(id).click();
  }, id);
}
try {
  const environment = { ...process.env, UNIREADMD_TEST_MODE: '1',
    UNIREADMD_TEST_USER_DATA: path.join(directory, 'profile') };
  delete environment.ELECTRON_RUN_AS_NODE;
  delete environment.ELECTRON_START_URL;
  const launch = () => electron.launch({ executablePath, env: environment, timeout: 60_000,
    args: process.env.UNIREADMD_SMOKE_EXECUTABLE ? [] : [path.join(root, 'electron/main.cjs')] });
  application = await launch();
  let page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await expect(page.locator('.preview')).toBeVisible();
  if (process.env.UNIREADMD_SMOKE_EXECUTABLE) {
    assert.equal(await application.evaluate(({ app }) => app.getVersion()), '1.1.0');
  } else {
    assert.equal(JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8')).version,
      '1.1.0');
  }
  await expect(page.locator('.document-tools')).toBeHidden();
  await expect(page.locator('.preferences-panel')).toHaveCount(0);
  assert.deepEqual(await application.evaluate(({ Menu }) => {
    const menu = Menu.getApplicationMenu();
    return ['view-document-tools', 'view-preferences'].map((id) => menu.getMenuItemById(id).checked);
  }), [false, false]);
  checks.push('真实应用版本 1.1.0，首次启动两个区域默认隐藏且菜单未勾选');
  await clickView('view-document-tools');
  await expect(page.locator('.document-tools')).toBeVisible();
  await clickView('view-preferences');
  await expect(page.locator('.preferences-panel')).toBeVisible();
  await clickView('view-document-tools');
  await expect(page.locator('.document-tools')).toBeHidden();
  await stop();
  application = await launch();
  page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await expect(page.locator('.preferences-panel')).toBeVisible();
  await expect(page.locator('.document-tools')).toBeHidden();
  assert.deepEqual(await application.evaluate(({ Menu }) => {
    const menu = Menu.getApplicationMenu();
    return ['view-document-tools', 'view-preferences'].map((id) => menu.getMenuItemById(id).checked);
  }), [false, true]);
  checks.push('视图菜单两个显示开关独立生效，重启后状态与勾选一致');
  const sourceToggle = page.getByRole('checkbox', { name: '源码模式' });
  await sourceToggle.check();
  await expect(page.locator('.cm-content')).toBeVisible();
  await expect(page.locator('#outline-btn')).toBeVisible();
  await page.locator('#outline-btn').click();
  await expect(page.locator('.sidebar')).toBeHidden();
  await sourceToggle.uncheck();
  await expect(page.locator('#outline-btn')).toHaveText('显示侧栏');
  await page.locator('#outline-btn').click();
  await expect(page.locator('.sidebar')).toBeVisible();
  checks.push('源码圆形选择控件可取消，侧栏开关始终可用且状态独立');
  const source = path.join(directory, '图片比例.md');
  const image = path.join(directory, '图片.png');
  const content = '# 图片比例回归\n\n![同图](图片.png) 和 ![同图](图片.png)';
  const png = await application.evaluate(({ nativeImage }) => {
    return nativeImage.createFromBitmap(Buffer.alloc(800 * 400 * 4, 230),
      { width: 800, height: 400 }).toPNG().toString('base64');
  });
  await fs.writeFile(image, Buffer.from(png, 'base64'));
  await fs.writeFile(source, content, 'utf8');
  await application.evaluate(({ dialog, Menu }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    dialog.showMessageBox = async () => ({ response: 1 });
    Menu.getApplicationMenu().items[0].submenu.items.find((item) => item.label === '打开...')
      .click();
  }, source);
  const images = page.locator('.preview img');
  await expect(images).toHaveCount(2);
  await expect.poll(() => images.nth(1).evaluate((image) => image.naturalWidth)).toBe(800);
  await images.nth(1).click({ button: 'right' });
  const menu = page.getByRole('dialog', { name: '图片显示比例' });
  await expect(menu).toBeVisible();
  if (evidence) {
    await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].showInactive());
    await page.screenshot({ path: path.join(evidence, 'native-image-menu.png') });
  }
  await menu.getByRole('button', { name: '50%', exact: true }).click();
  await expect(images.nth(1)).toHaveAttribute('width', '400');
  await expect(images.first()).not.toHaveAttribute('width');
  await page.keyboard.press('Control+s');
  await expect.poll(() => fs.readFile(source, 'utf8')).toContain('width="400"');
  assert.match(await fs.readFile(source, 'utf8'), /!\[同图\]\(图片.png\) 和 <img/);
  checks.push('真实本地图片授权加载，右键只调整选中图片并保存到磁盘');
  const exportedFile = path.join(directory, '图片比例.html');
  await application.evaluate(({ dialog, BrowserWindow }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath });
    BrowserWindow.getAllWindows()[0].webContents.send('menu:command', 'file-export-html');
  }, exportedFile);
  await expect.poll(async () => fs.stat(exportedFile).then((stat) => stat.size).catch(() => 0))
    .toBeGreaterThan(1000);
  browser = await chromium.launch();
  const exported = await browser.newPage();
  await exported.goto(pathToFileURL(exportedFile).href);
  await expect(exported.locator('img').nth(1)).toHaveAttribute('width', '400');
  await expect.poll(() => exported.locator('img').nth(1)
    .evaluate((image) => image.naturalWidth)).toBe(800);
  if (evidence) await exported.screenshot({ path: path.join(evidence, 'native-image-export.png') });
  checks.push('真实 IPC 导出内嵌本地图片，离线打开后仍保留比例');
  const preview = page.locator('.preview');
  const bounds = await preview.boundingBox();
  await preview.click({ button: 'right',
    position: { x: bounds.width - 12, y: bounds.height - 20 } });
  await page.getByRole('menuitem', { name: '表格…' }).click();
  await page.getByRole('button', { name: '插入表格', exact: true }).click();
  await expect(page.locator('.preview table th')).toHaveCount(2);
  await expect(page.getByTestId('preview-table-cell-editor-input')).toBeFocused();
  checks.push('真实工作区空白右键插入表格并直接编辑表头');
} finally {
  await browser?.close();
  await stop();
  const target = await fs.realpath(directory);
  assert.equal(path.dirname(target), await fs.realpath(os.tmpdir()));
  await fs.rm(target, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
}
const result = { ok: true, executablePath, checks };
if (evidence) await fs.writeFile(path.join(evidence, 'native-image-workspace-result.json'),
  JSON.stringify(result, null, 2) + '\r\n', 'utf8');
console.log(JSON.stringify(result, null, 2));
