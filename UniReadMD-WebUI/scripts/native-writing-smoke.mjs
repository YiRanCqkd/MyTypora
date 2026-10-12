import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { once } from 'node:events';
import { fileURLToPath } from 'node:url';
import { showWorkspaceControls } from './native-workspace-controls.mjs';

const require = createRequire(import.meta.url);
const { _electron: electron, expect } = require('../frontend/node_modules/@playwright/test');
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-writing-'));
const executablePath = process.env.UNIREADMD_SMOKE_EXECUTABLE || require('electron');
const evidence = process.env.UNIREADMD_SMOKE_EVIDENCE;
let application;
let result;
const checks = [];

async function stop() {
  if (!application) return;
  const child = application.process();
  const exited = child.exitCode !== null ? Promise.resolve() : once(child, 'exit');
  await application.evaluate(({ app }) => app.exit(0)).catch(() => {});
  await exited;
  application = undefined;
}

try {
  const environment = { ...process.env, UNIREADMD_TEST_MODE: '1',
    UNIREADMD_TEST_USER_DATA: path.join(directory, 'profile'),
  };
  delete environment.ELECTRON_RUN_AS_NODE;
  delete environment.ELECTRON_START_URL;
  if (process.env.UNIREADMD_SMOKE_PANDOC) {
    environment.PATH = path.dirname(process.env.UNIREADMD_SMOKE_PANDOC)
      + path.delimiter + environment.PATH;
  }
  const launch = () => electron.launch({ executablePath, env: environment, timeout: 60_000,
    args: process.env.UNIREADMD_SMOKE_EXECUTABLE ? [] : [path.join(root, 'electron/main.cjs')],
  });
  application = await launch();
  let page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await expect(page.locator('.preview')).toBeVisible();
  await showWorkspaceControls(application, page);
  const source = path.join(directory, '写作文档.md');
  const imagePath = path.join(directory, '图片.png');
  const png = await application.evaluate(({ nativeImage }) => {
    return nativeImage.createFromBitmap(Buffer.from([20, 40, 80, 255]),
      { width: 1, height: 1 }).toPNG().toString('base64');
  });
  const content = '# 写作回归\n\n前文 **加粗** 后文\n\n参见 $\\eqref{mass}$\n\n'
    + '$$\nE=mc^2\\label{mass}\n$$\n\n$\\ce{H2O}$\n\n![图片](图片.png)';
  await fs.writeFile(source, content);
  await fs.writeFile(imagePath, Buffer.from(png, 'base64'));
  await application.evaluate(({ dialog, Menu }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    dialog.showMessageBox = async () => ({ response: 1 });
    Menu.getApplicationMenu().items[0].submenu.items.find((item) => item.label === '打开...')
      .click();
  }, source);
  await expect(page.locator('.preview h1')).toHaveText('写作回归');
  let headingFont;
  // 公式异步渲染可能替换标题节点，只采样仍连接到文档的真实字号和行距。
  await expect.poll(async () => {
    headingFont = await page.locator('.preview h1').evaluate((node) => ({
      size: getComputedStyle(node).fontSize, height: getComputedStyle(node).lineHeight,
      connected: node.isConnected,
    }));
    return headingFont.connected && /^\d/.test(headingFont.size)
      && /^\d/.test(headingFont.height);
  }).toBe(true);
  await page.locator('.preview h1').click({ modifiers: ['Alt'] });
  await expect(page.locator('.preview-inline-editor-input'))
    .toHaveCSS('font-size', headingFont.size);
  await expect(page.locator('.preview-inline-editor-input'))
    .toHaveCSS('line-height', headingFont.height);
  await page.locator('.preview-inline-editor-input').press('Control+Enter');
  await page.locator('.preview strong').click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await expect(input).toHaveValue('**加粗**');
  await input.fill('**修改**');
  await input.press('Control+Enter');
  await expect(page.locator('.preview strong')).toHaveText('修改');
  checks.push('安装产物标题字号保持、格式片段编辑与提交');

  await page.locator('#footer-word-count').click();
  const settings = page.getByRole('dialog', { name: '写作设置' });
  await settings.getByLabel('公式自动编号').check();
  await settings.getByLabel('浅色写作主题').selectOption('sepia');
  await settings.getByLabel('深色写作主题').selectOption('midnight');
  await settings.getByLabel('阅读速度（字／分钟）').fill('300');
  await settings.getByRole('button', { name: '完成' }).click();
  await expect(page.locator('.math-reference')).toHaveText('(1)');
  await expect(page.locator('.katex-error')).toHaveCount(0);
  checks.push('安装产物自动编号、前向引用及化学公式');

  await application.evaluate(({ dialog }, destination) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [destination] });
  }, directory);
  await page.getByRole('button', { name: '整理图片', exact: true }).click();
  await expect(page.getByRole('status')).toContainText('已整理 1 张本地图片');
  await application.evaluate(({ Menu }) => {
    Menu.getApplicationMenu().items[0].submenu.items.find((item) => item.label === '保存').click();
  });
  await expect.poll(() => fs.readFile(source, 'utf8')).toContain('-images-');
  const saved = await fs.readFile(source, 'utf8');
  const reference = /!\[图片\]\(([^)]+)\)/.exec(saved)[1];
  assert.deepEqual(await fs.readFile(path.resolve(directory, decodeURI(reference))),
    Buffer.from(png, 'base64'));
  assert.deepEqual(await fs.readFile(imagePath), Buffer.from(png, 'base64'));
  await expect.poll(() => page.locator('.preview img').evaluate((image) => image.naturalWidth))
    .toBe(1);
  checks.push('真实图片整理 IPC、原图保留、保存后图片可读');

  const capabilities = await page.evaluate(() => window.nativeBridge.exporter.getCapabilities());
  for (const format of ['revealjs', 'mediawiki', 'rst', 'textile', 'opml']) {
    assert.ok(capabilities.pandocFormats.includes(format), `安装产物缺少 ${format}`);
  }
  checks.push('安装产物识别五种新增 Pandoc 格式');
  await page.getByText('偏好设置', { exact: true }).click();
  await page.getByRole('combobox', { name: '主题', exact: true }).selectOption('system');
  await page.emulateMedia({ colorScheme: 'dark' });
  await expect(page.locator('html')).toHaveAttribute('data-writing-theme', 'midnight');
  await expect(page.locator('.preview')).toHaveCSS('background-color', 'rgb(20, 30, 46)');
  await stop();
  application = await launch();
  page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  // 原生回归需要真实合成帧；隐藏窗口在重启后可能暂停截图和交互稳定性检查。
  await application.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0].showInactive();
  });
  await page.emulateMedia({ colorScheme: 'dark' });
  await expect(page.locator('html')).toHaveAttribute('data-writing-theme', 'midnight');
  await page.locator('#footer-word-count').click();
  await expect(page.getByLabel('阅读速度（字／分钟）')).toHaveValue('300');
  await expect(page.getByLabel('公式自动编号')).toBeChecked();
  checks.push('安装产物系统主题、写作设置跨进程恢复');
  if (evidence) {
    await fs.mkdir(evidence, { recursive: true });
    await page.screenshot({ path: path.join(evidence, 'native-writing-settings.png') });
  }
  await page.getByRole('button', { name: '完成', exact: true }).click();
  const large = '# 性能测试\n\n'
    + ('正文内容 '.repeat(100) + 'search-target\n\n').repeat(5000);
  const largePath = path.join(directory, '大文档.md');
  await fs.writeFile(largePath, large);
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    dialog.showMessageBox = async () => ({ response: 1 });
  }, largePath);
  const opened = Date.now();
  await application.evaluate(({ Menu }) => {
    Menu.getApplicationMenu().items[0].submenu.items.find((item) => item.label === '打开...')
      .click();
  });
  await expect(page.locator('.preview p')).toHaveCount(5000);
  const openMs = Date.now() - opened;
  await page.locator('#toggle-sourceview-btn').click();
  const searched = Date.now();
  await page.keyboard.press('Control+f');
  await page.getByTestId('doc-search-query').fill('search-target');
  await expect(page.locator('.doc-search-summary')).toHaveText('1/5000');
  const searchMs = Date.now() - searched;
  await page.keyboard.press('Escape');
  await page.locator('.cm-content').focus();
  await page.keyboard.press('Control+End');
  const edited = Date.now();
  await page.keyboard.insertText('尾部编辑验证');
  await expect(page.locator('.preview p').last()).toHaveText('尾部编辑验证');
  const editMs = Date.now() - edited;
  assert.ok(openMs < 30_000 && searchMs < 10_000 && editMs < 10_000);
  checks.push('安装产物真实磁盘读取六 MiB 文档、搜索及编辑');
  const performance = { bytes: Buffer.byteLength(large), openMs, searchMs, editMs };
  const blankPath = path.join(directory, '空白文档.md');
  await fs.writeFile(blankPath, '');
  await application.evaluate(({ dialog, Menu }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    Menu.getApplicationMenu().items[0].submenu.items.find((item) => item.label === '打开...')
      .click();
  }, blankPath);
  if (await page.locator('.cm-content').isVisible()) {
    await page.locator('#toggle-sourceview-btn').click();
  }
  await page.locator('.preview-empty-line').click({ modifiers: ['Alt'] });
  const blankInput = page.locator('.preview-inline-editor-input');
  await blankInput.fill('原生空白写作');
  await blankInput.press('Control+Enter');
  await page.keyboard.press('Control+s');
  await expect.poll(() => fs.readFile(blankPath, 'utf8')).toBe('原生空白写作');
  checks.push('安装产物打开真实空白文件、直接输入并保存到原路径');
  await page.locator('.preview p').click({ modifiers: ['Alt'] });
  const maximizedInput = page.locator('.preview-inline-editor-input');
  await maximizedInput.evaluate((node) => node.setSelectionRange(2, 4, 'backward'));
  await application.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0].maximize();
  });
  await expect.poll(() => application.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()[0].isMaximized())).toBe(true);
  await expect.poll(async () => {
    const host = (await page.getByTestId('editor-preview').boundingBox()).width;
    const preview = (await page.getByTestId('preview-host').boundingBox()).width;
    const anchor = await page.locator('.preview p').boundingBox();
    const editor = await maximizedInput.boundingBox();
    return Math.abs(host - preview) + Math.abs(anchor.width - editor.width)
      + Math.abs(anchor.x - editor.x) + Math.abs(anchor.y - editor.y);
  }).toBeLessThanOrEqual(2);
  await expect(maximizedInput).toBeFocused();
  assert.deepEqual(await maximizedInput.evaluate((node) => ({
    start: node.selectionStart, end: node.selectionEnd, direction: node.selectionDirection,
  })), { start: 2, end: 4, direction: 'backward' });
  if (evidence) {
    await page.screenshot({ path: path.join(evidence, 'native-maximized-editing.png') });
  }
  await maximizedInput.press('Control+Enter');
  checks.push('安装产物真实窗口最大化、编辑框对齐与反向选区保留');
  await application.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0].unmaximize();
    BrowserWindow.getAllWindows()[0].setSize(920, 800);
  });
  const more = page.getByRole('button', { name: '更多文档工具' });
  await expect(more).toBeVisible();
  await more.click();
  await expect(page.getByRole('button', { name: '导出设置', exact: true })).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(more).toBeFocused();
  await expect(more).toHaveAttribute('aria-expanded', 'false');
  if (evidence) {
    await page.screenshot({ path: path.join(evidence, 'native-compact-writing.png') });
  }
  checks.push('安装产物窄窗口紧凑工具栏、次要工具和键盘焦点恢复');
  result = { ok: true, executablePath, checks, performance };
} catch (error) {
  console.error('写作原生回归失败', JSON.stringify({ checks, error: String(error) }));
  process.exitCode = 1;
} finally {
  await stop();
  const resolved = await fs.realpath(directory);
  if (path.dirname(resolved) !== await fs.realpath(os.tmpdir())
    || !path.basename(resolved).startsWith('unireadmd-writing-')) {
    throw new Error('写作回归清理目录越界');
  }
  await fs.rm(resolved, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
}
if (result) {
  if (evidence) {
    await fs.writeFile(path.join(evidence, 'native-writing-result.json'),
      JSON.stringify(result, null, 2));
  }
  console.log(JSON.stringify(result, null, 2));
}
