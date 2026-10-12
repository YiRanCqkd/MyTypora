import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { once } from 'node:events';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const { _electron: electron, expect } = require('../frontend/node_modules/@playwright/test');
const executablePath = process.env.UNIREADMD_SMOKE_EXECUTABLE;
const evidenceDirectory = process.env.UNIREADMD_SMOKE_EVIDENCE;
assert.ok(executablePath && evidenceDirectory, '需要指定测试程序及证据目录');
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-clipboard-'));
const documentPath = path.join(directory, '剪贴板.md');
let application;
let clipboardSaved = false;
let clipboardRestored = false;
const checks = [];

try {
  await fs.writeFile(documentPath, '# 剪贴板测试');
  const environment = { ...process.env, UNIREADMD_TEST_MODE: '1',
    UNIREADMD_TEST_USER_DATA: path.join(directory, '用户数据') };
  delete environment.ELECTRON_RUN_AS_NODE;
  delete environment.ELECTRON_START_URL;
  application = await electron.launch({ executablePath, args: [], env: environment });
  const page = await application.firstWindow();
  await expect(page.getByTestId('preview-host')).toContainText('UniReadMD');
  await application.evaluate(async ({ clipboard, ClipboardItem }) => {
    const saved = [];
    for (const item of await clipboard.read()) {
      if (!item.types.length) continue;
      const values = {};
      for (const type of item.types) values[type] = await item.getType(type);
      saved.push(new ClipboardItem(values));
    }
    // 原剪贴板仅保留在本次进程内，不写入日志或测试证据。
    globalThis.__smokeClipboardSaved = saved;
    globalThis.__smokeClipboardExpected = null;
  });
  clipboardSaved = true;
  await application.evaluate(({ dialog, Menu }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    Menu.getApplicationMenu().items[0].submenu.items[0].click();
  }, documentPath);
  await expect(page.locator('.preview h1')).toHaveText('剪贴板测试');
  await page.locator('#toggle-sourceview-btn').click();
  const editor = page.locator('.cm-content');
  await editor.click();
  await page.keyboard.press('Control+End');
  await page.keyboard.insertText('\n\n');
  await application.evaluate(async ({ clipboard, ClipboardItem, BrowserWindow }) => {
    await clipboard.write([new ClipboardItem({ 'text/plain': '原生富文本\n粗体',
      'text/html': '<h2>原生富文本</h2><p><b>粗体</b></p>' })]);
    globalThis.__smokeClipboardExpected = await clipboard.readText();
    BrowserWindow.getAllWindows()[0].webContents.paste();
  });
  await expect(editor).toContainText('## 原生富文本');
  await expect(editor).toContainText('**粗体**');
  checks.push('系统 HTML 剪贴板经原生粘贴事件转换为 Markdown');
  console.log('原生 HTML 粘贴通过');

  await editor.focus();
  await page.keyboard.press('Control+A');
  await application.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0].webContents.copy();
  });
  await expect.poll(() => application.evaluate(async ({ clipboard }) => {
    for (const item of await clipboard.read()) {
      if (item.types.includes('text/html')) return (await item.getType('text/html')).text();
    }
    return '';
  }))
    .toContain('<strong>粗体</strong>');
  const copied = await application.evaluate(async ({ clipboard }) => {
    globalThis.__smokeClipboardExpected = await clipboard.readText();
    return clipboard.readText();
  });
  assert.match(copied, /## 原生富文本/);
  checks.push('原生复制同时写入 Markdown 文本与 HTML 格式');
  console.log('原生双格式复制通过');

  await application.evaluate(async ({ clipboard, ClipboardItem, nativeImage }) => {
    const image = nativeImage.createFromBitmap(Buffer.from([0, 0, 0, 255]),
      { width: 1, height: 1 });
    if (image.isEmpty()) throw new Error('原生截图测试图片生成失败');
    await clipboard.write([new ClipboardItem({
      'image/png': new Blob([image.toPNG()], { type: 'image/png' }),
    })]);
    globalThis.__smokeClipboardExpected = await clipboard.readText();
  });
  const pasted = await page.evaluate((sourceFilePath) => {
    return window.nativeBridge.images.paste({ sourceFilePath, copyDirectory: 'assets' });
  }, documentPath);
  assert.equal(pasted.images.length, 1);
  const imagePath = path.resolve(directory, pasted.images[0].reference);
  assert.equal(path.dirname(imagePath), path.join(directory, 'assets'));
  assert.equal(await application.evaluate(({ nativeImage }, filePath) => {
    return nativeImage.createFromPath(filePath).getSize().width;
  }, imagePath), 1);
  checks.push('系统图片剪贴板通过真实 IPC 保存截图到文档资源目录');
} finally {
  if (application) {
    if (clipboardSaved) {
      clipboardRestored = await application.evaluate(async ({ clipboard }) => {
        // 用户若在验证期间换了剪贴板，保留用户的最新内容。
        if (await clipboard.readText() !== globalThis.__smokeClipboardExpected) return false;
        clipboard.clear();
        const saved = globalThis.__smokeClipboardSaved;
        if (saved.length) await clipboard.write(saved);
        return true;
      });
    }
    const child = application.process();
    const exited = child.exitCode !== null ? Promise.resolve() : once(child, 'exit');
    await application.evaluate(({ app }) => app.exit(0)).catch(() => {});
    await exited;
  }
  const temporaryRoot = await fs.realpath(os.tmpdir());
  const target = await fs.realpath(directory);
  assert.equal(path.dirname(target), temporaryRoot);
  await fs.rm(target, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
}

const result = { ok: true, checks, clipboardRestored };
await fs.writeFile(path.join(evidenceDirectory, 'native-clipboard-result.json'),
  JSON.stringify(result, null, 2) + '\r\n', 'utf8');
console.log(JSON.stringify(result, null, 2));
