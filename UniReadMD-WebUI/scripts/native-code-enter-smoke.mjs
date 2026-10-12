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
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-code-enter-'));
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

async function save(page) {
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
  const source = path.join(directory, '代码与换行回归.md');
  const original = '正文\n\n```js\nconst value = "' + '长行abc'.repeat(80) + '";\n\n```';
  await fs.writeFile(source, original, 'utf8');
  application = await launch();
  let page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await application.evaluate(({ BrowserWindow }) => {
    const window = BrowserWindow.getAllWindows()[0];
    window.show(); window.focus(); window.webContents.focus();
  });
  await expect(page.locator('.preview')).toBeVisible();
  await application.evaluate(({ dialog, BrowserWindow }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    BrowserWindow.getAllWindows()[0].webContents.send('menu:command', 'file-open');
  }, source);
  const pre = page.locator('.preview pre');
  const before = (await pre.boundingBox());
  await pre.click({ position: { x: 5, y: 5 } });
  const input = page.locator('.preview-inline-code textarea');
  await expect(input).toBeFocused();
  await expect(page.locator('.preview-code-highlight .hljs-keyword')).toHaveText('const');
  const editor = (await page.locator('.preview-inline-code').boundingBox());
  assert.ok(Math.abs(editor.x - before.x) < 1, '代码背景留白也应打开完整代码块编辑');
  assert.ok(Math.abs(editor.width - before.width) < 1);
  await input.evaluate((node) => { node.scrollLeft = 320; });
  await expect.poll(() => page.locator('.preview-code-highlight')
    .evaluate((node) => node.scrollLeft)).toBe(320);
  if (evidence) await page.screenshot({ path: path.join(evidence, 'native-code-editor.png') });
  checks.push('代码块背景点击、完整尺寸、等宽高亮与横向滚动在实际应用中一致');

  await input.fill('```js\nconst value = "中文源码";\n\n```');
  await save(page);
  const codeSaved = '正文\n\n```js\nconst value = "中文源码";\n\n```';
  await expect.poll(() => fs.readFile(source, 'utf8')).toBe(codeSaved);
  await expect(pre.locator('code')).toHaveText('const value = "中文源码";\n\n');
  checks.push('真实菜单保存代码保留原始围栏与空行，并写入真实磁盘');

  const paragraph = page.locator('.preview p');
  await paragraph.click();
  await paragraph.press('End');
  const paragraphBefore = (await paragraph.boundingBox());
  const lineHeight = await paragraph.evaluate((node) =>
    parseFloat(getComputedStyle(node).lineHeight));
  await paragraph.press('Enter');
  assert.ok(Math.abs((await paragraph.boundingBox()).height
    - paragraphBefore.height - lineHeight) < 1, '一次回车只增加一行');
  await page.keyboard.insertText('下一行正文');
  if (evidence) await page.screenshot({ path: path.join(evidence, 'native-single-enter.png') });
  await save(page);
  const expected = codeSaved.replace('正文', '正文\n下一行正文');
  await expect.poll(() => fs.readFile(source, 'utf8')).toBe(expected);
  await expect(page.locator('.preview > p')).toHaveCount(1);
  checks.push('正文回车只换一行，继续输入与真实保存均无额外空白行');

  await stop(true);
  application = await launch();
  page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await expect(page.locator('.preview p')).toHaveText('正文\n下一行正文');
  await expect(page.locator('.preview pre code')).toContainText('中文源码');
  assert.equal(await fs.readFile(source, 'utf8'), expected);
  checks.push('正常关闭并重启后，正文单换行与代码源码保持一致');
} finally {
  await stop();
  // 仅删除本脚本创建的独立配置和样本，保留正式程序及用户文档。
  const target = await fs.realpath(directory);
  assert.equal(path.dirname(target), await fs.realpath(os.tmpdir()));
  assert.ok(path.basename(target).startsWith('unireadmd-code-enter-'));
  await fs.rm(target, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
}
const result = { ok: true, executablePath, checks };
if (evidence) await fs.writeFile(path.join(evidence, 'native-code-enter-result.json'),
  JSON.stringify(result, null, 2) + '\r\n', 'utf8');
console.log(JSON.stringify(result, null, 2));
