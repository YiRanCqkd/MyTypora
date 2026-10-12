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
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-review-editing-'));
const evidence = process.env.UNIREADMD_SMOKE_EVIDENCE;
const executablePath = process.env.UNIREADMD_SMOKE_EXECUTABLE || require('electron');
const checks = [];
const errors = [];
let application;
let page;

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

async function launch() {
  const environment = { ...process.env, UNIREADMD_TEST_MODE: '1',
    UNIREADMD_TEST_USER_DATA: path.join(directory, 'profile') };
  delete environment.ELECTRON_RUN_AS_NODE;
  delete environment.ELECTRON_START_URL;
  application = await electron.launch({ executablePath, env: environment, timeout: 60_000,
    args: process.env.UNIREADMD_SMOKE_EXECUTABLE ? [] : [path.join(root, 'electron/main.cjs')] });
  page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  page.on('pageerror', (error) => errors.push(error.message));
  await application.evaluate(({ BrowserWindow }) => {
    const window = BrowserWindow.getAllWindows()[0];
    window.show(); window.focus(); window.webContents.focus();
  });
  await expect(page.locator('.preview')).toBeVisible();
}

async function openDocument(name, content) {
  const filePath = path.join(directory, name + '.md');
  await fs.writeFile(filePath, content, 'utf8');
  await application.evaluate(({ dialog, BrowserWindow }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    BrowserWindow.getAllWindows()[0].webContents.send('menu:command', 'file-open');
  }, filePath);
  return filePath;
}

async function save(filePath, expected) {
  await application.evaluate(({ Menu }) => Menu.getApplicationMenu().items[0].submenu.items
    .find((item) => item.label === '保存').click());
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  await expect(page.locator('.error-banner')).toHaveCount(0);
  await expect.poll(() => fs.readFile(filePath, 'utf8')).toBe(expected);
}

try {
  await launch();
  const paragraphSource = 'first\n\n**second**';
  const paragraphFile = await openDocument('段落退格', paragraphSource);
  await expect(page.locator('.preview > p')).toHaveCount(2);
  const paragraph = page.locator('.preview > p').nth(1);
  await paragraph.click();
  await paragraph.press('Control+Home');
  await paragraph.press('Enter');
  await paragraph.press('Backspace');
  await expect(page.locator('.preview > p')).toHaveCount(2);
  await save(paragraphFile, paragraphSource);
  checks.push('加粗段落开头回车再退格不合并前段，真实保存保持原文');

  const table = '| A | B |\n| --- | --- |\n| one | two |';
  const tableFile = await openDocument('表格范围', `正文\n\n${table}\n\n${table}`);
  await expect(page.locator('.preview table')).toHaveCount(2);
  const leading = page.locator('.preview > p').first();
  await leading.click();
  await leading.press('End');
  await leading.press('Enter');
  // 同一事件循环内输入并右键，覆盖尚未防抖提交的文字。
  await leading.evaluate((node) => {
    document.execCommand('insertText', false, '新增正文');
    const cell = node.closest('.preview').querySelectorAll('table')[1].querySelector('td');
    cell.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true,
      cancelable: true, button: 2, clientX: 400, clientY: 250 }));
  });
  await page.getByTestId('preview-table-menu')
    .getByRole('button', { name: '在下方插入行', exact: true }).click();
  await expect(page.locator('.preview table').first().locator('tbody tr')).toHaveCount(1);
  await expect(page.locator('.preview table').last().locator('tbody tr')).toHaveCount(2);
  await save(tableFile, `正文\n新增正文\n\n${table}\n\n${table}\n|  |  |`);
  checks.push('正文输入后立即右键第二张相同表格，插入位置正确且保留空行');

  const shortFile = await openDocument('省略竖线', '| A |\n| --- |\none\ntwo');
  await expect(page.locator('.preview td')).toHaveCount(2);
  const cell = page.locator('.preview td').last();
  await cell.click();
  await expect(cell).toHaveAttribute('contenteditable', 'true');
  await cell.fill('修改值');
  const expected = '| A |\n| --- |\n| one |\n| 修改值 |';
  await save(shortFile, expected);
  await expect(page.locator('.preview tbody tr')).toHaveCount(2);
  assert.deepEqual(errors, []);
  checks.push('省略竖线的表格可原位编辑，真实保存保留全部数据行且无页面异常');

  await stop(true);
  await launch();
  await expect(page.locator('.preview td')).toHaveText(['one', '修改值']);
  assert.equal(await fs.readFile(shortFile, 'utf8'), expected);
  assert.deepEqual(errors, []);
  checks.push('正常关闭重启后恢复已保存表格，内容与磁盘一致');
} finally {
  await stop();
  // 仅清理本脚本创建的独立配置与样本，验证绝对路径后再递归删除。
  const target = await fs.realpath(directory);
  assert.equal(path.dirname(target), await fs.realpath(os.tmpdir()));
  assert.ok(path.basename(target).startsWith('unireadmd-review-editing-'));
  await fs.rm(target, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
}
const result = { ok: true, executablePath, checks };
if (evidence) await fs.writeFile(path.join(evidence, 'native-review-editing-result.json'),
  JSON.stringify(result, null, 2) + '\r\n', 'utf8');
console.log(JSON.stringify(result, null, 2));
