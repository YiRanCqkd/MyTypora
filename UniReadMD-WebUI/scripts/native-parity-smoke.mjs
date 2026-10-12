import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { once } from 'node:events';

const require = createRequire(import.meta.url);
const { _electron: electron, expect } = require('../frontend/node_modules/@playwright/test');
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-native-'));
const first = path.join(directory, '文档一.md');
const second = path.join(directory, '文档二.md');
const exported = path.join(directory, '全文.png');
const image = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8'
  + '/x8AAwMCAO+a2ioAAAAASUVORK5CYII=', 'base64');
let application;
const checks = [];
const rendererErrors = [];
let processDiagnostics = '';
const executablePath = process.env.UNIREADMD_SMOKE_EXECUTABLE;
const evidenceDirectory = process.env.UNIREADMD_SMOKE_EVIDENCE;

async function stopApplication() {
  if (!application) return;
  const child = application.process();
  const exited = child.exitCode !== null ? Promise.resolve() : once(child, 'exit');
  await application.evaluate(({ app }) => app.exit(0)).catch(() => {});
  await exited;
  application = undefined;
}

async function invokeFileMenu(label) {
  await application.evaluate(({ Menu }, itemLabel) => {
    const item = Menu.getApplicationMenu().items[0].submenu.items
      .find((entry) => entry.label === itemLabel);
    if (!item) throw new Error(`找不到原生菜单：${itemLabel}`);
    item.click();
  }, label);
}

try {
  await fs.writeFile(first, '# 文档一\n\n公式 $x^2$\n\n![本地图片](图片.png)');
  await fs.writeFile(second, '# 文档二');
  await fs.writeFile(path.join(directory, '图片.png'), image);
  const environment = { ...process.env,
    UNIREADMD_TEST_MODE: '1', UNIREADMD_TEST_USER_DATA: path.join(directory, '用户数据'),
  };
  delete environment.ELECTRON_RUN_AS_NODE;
  delete environment.ELECTRON_START_URL;
  if (process.env.UNIREADMD_SMOKE_PANDOC) {
    environment.PATH = path.dirname(process.env.UNIREADMD_SMOKE_PANDOC)
      + path.delimiter + environment.PATH;
  }
  const launch = async () => {
    const launched = await electron.launch({
      executablePath: executablePath || require('electron'),
      args: executablePath ? [] : [path.join(root, 'electron/main.cjs')],
      env: environment,
      timeout: 60_000,
    });
    launched.process().stderr.on('data', (chunk) => { processDiagnostics += String(chunk); });
    launched.process().stdout.on('data', (chunk) => { processDiagnostics += String(chunk); });
    launched.on('window', (window) => {
      window.on('pageerror', (error) => rendererErrors.push(String(error)));
    });
    await launched.evaluate(() => {
      process.on('uncaughtExceptionMonitor', (error) => {
        console.error('[native-smoke-uncaught]', String(error));
      });
    });
    return launched;
  };
  application = await launch();
  let page = await application.firstWindow();
  page.setDefaultTimeout(20_000);
  await expect(page.getByTestId('preview-host')).toContainText('UniReadMD');
  assert.deepEqual(await application.evaluate(({ BrowserWindow }) => {
    const preferences = BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences();
    return { sandbox: preferences.sandbox, webSecurity: preferences.webSecurity,
      nodeIntegration: preferences.nodeIntegration, contextIsolation: preferences.contextIsolation,
    };
  }), { sandbox: true, webSecurity: true, nodeIntegration: false, contextIsolation: true });
  const denied = await page.evaluate(async () => {
    try { await window.nativeBridge.file.openFromPath({ filePath: 'C:/Windows/win.ini' }); }
    catch (error) { return String(error); }
    return '';
  });
  assert.match(denied, /ERR_PATH_NOT_AUTHORIZED/);
  checks.push('原生沙箱、IPC 和路径授权');
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
    dialog.showMessageBox = async () => ({ response: 2 });
  }, first);
  // CDP 键盘输入不会稳定触发 Windows 菜单加速键，直接调用真实菜单入口。
  await invokeFileMenu('打开...');
  await expect(page.locator('.preview h1')).toHaveText('文档一');
  await expect(page.locator('.preview .katex')).toHaveCount(1);
  await expect.poll(() => page.locator('.preview img').evaluate((node) => node.naturalWidth))
    .toBe(1);
  checks.push('首次公式渲染和中文相对图片');
  await expect.poll(() => page.evaluate(async () => {
    return (await window.nativeBridge.session.load())?.content || '';
  })).toContain('文档一');
  await expect(page.locator('.error-banner')).toHaveCount(0);
  checks.push('打开文档自动保存恢复快照，无克隆失败提示');
  await page.locator('#toggle-sourceview-btn').click();
  const editor = page.locator('.cm-content').first();
  await editor.click();
  await page.keyboard.press('Control+End');
  await page.keyboard.insertText('\n\n<img src="x" onerror="window.__auditProbe=123">');
  await expect(page.locator('.dirty-dot')).toHaveCount(1);
  assert.equal(await page.evaluate(() => window.__auditProbe), undefined);
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
  }, second);
  await invokeFileMenu('打开...');
  await expect(editor).toContainText('文档一');
  await fs.writeFile(first, '# 外部修改');
  await invokeFileMenu('保存');
  await expect(page.locator('.error-banner')).toContainText('外部修改');
  assert.equal(await fs.readFile(first, 'utf8'), '# 外部修改');
  checks.push('XSS、取消切换、外部修改冲突');
  await application.evaluate(({ dialog }) => {
    dialog.showMessageBox = async () => ({ response: 1 });
  });
  await invokeFileMenu('打开...');
  await expect(editor).toContainText('文档二');
  await editor.click();
  await page.keyboard.press('Control+Z');
  await expect(editor).toHaveText('# 文档二');
  checks.push('文档撤销历史隔离');
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath });
  }, exported);
  await page.evaluate(async () => window.nativeBridge.exporter.exportDocument({
    format: 'png', markdown: '',
    html: '<style>body{margin:0} article{height:2400px;background:#ccc}</style>'
      + '<article>长文档末尾</article>',
  }));
  assert.equal(await application.evaluate(({ nativeImage }, filePath) => {
    return nativeImage.createFromPath(filePath).getSize().height;
  }, exported), 2400);
  checks.push('真实 Chromium 全文图片导出');
  const jpeg = path.join(directory, '尺寸质量.jpg');
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath });
  }, jpeg);
  await page.evaluate(async () => window.nativeBridge.exporter.exportDocument({
    format: 'jpeg', markdown: '', html: '<article>原生 JPEG 导出</article>',
    options: { imageWidth: 864, jpegQuality: 75 },
  }));
  assert.equal(await application.evaluate(({ nativeImage }, filePath) => {
    return nativeImage.createFromPath(filePath).getSize().width;
  }, jpeg), 864);
  const pdf = path.join(directory, '自定义纸张.pdf');
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath });
  }, pdf);
  await page.evaluate(async () => window.nativeBridge.exporter.exportDocument({
    format: 'pdf', markdown: '', html: '<h1>原生 PDF 导出</h1><p>中文正文</p>',
    pdfOptions: { pageSize: 'Custom', widthMm: 148, heightMm: 210,
      marginsMm: { top: 10, bottom: 15, left: 12, right: 18 } },
  }));
  const pdfBytes = await fs.readFile(pdf);
  assert.equal(pdfBytes.subarray(0, 5).toString(), '%PDF-');
  const mediaBox = /\/MediaBox\s*\[0 0 ([\d.]+) ([\d.]+)\]/
    .exec(pdfBytes.toString('latin1'));
  assert.ok(mediaBox, 'PDF 必须包含实际纸张尺寸');
  assert.ok(Math.abs(Number(mediaBox[1]) - 148 / 25.4 * 72) < 1);
  assert.ok(Math.abs(Number(mediaBox[2]) - 210 / 25.4 * 72) < 1);
  checks.push('真实 JPEG 尺寸质量和 PDF 自定义纸张、独立边距导出');

  if (process.env.UNIREADMD_SMOKE_PANDOC) {
    const capabilities = await page.evaluate(() => window.nativeBridge.exporter.getCapabilities());
    assert.equal(capabilities.pandocAvailable, true);
    for (const format of ['docx', 'odt', 'rtf', 'epub', 'latex']) {
      const target = path.join(directory, `原生转换.${format === 'latex' ? 'tex' : format}`);
      await application.evaluate(({ dialog }, filePath) => {
        dialog.showSaveDialog = async () => ({ canceled: false, filePath });
      }, target);
      await page.evaluate(async ({ format, first }) => {
        return window.nativeBridge.exporter.exportDocument({
          format, sourceFilePath: first, html: '<h1>原生转换</h1>',
          markdown: '# 原生转换\n\n![本地图片](图片.png)',
          options: { title: '原生转换标题', author: '原生验证', language: 'zh-CN' },
        });
      }, { format, first });
      assert.ok((await fs.stat(target)).size > 0);
    }
    checks.push('安装程序通过真实 IPC 调用 Pandoc 五种格式并处理中文相对图片');
  }

  const migration = path.join(directory, '迁移图片.md');
  const movedDirectory = path.join(directory, '移动目标');
  const savedDirectory = path.join(directory, '另存目标');
  await fs.writeFile(migration, '# 迁移测试\n\n![本地图片](图片.png)');
  await fs.mkdir(movedDirectory);
  await fs.mkdir(savedDirectory);
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
  }, migration);
  await page.evaluate(() => window.nativeBridge.file.openMarkdown());
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [filePath] });
  }, movedDirectory);
  await page.evaluate(() => window.nativeBridge.file.chooseMoveDirectory());
  const moved = await page.evaluate(({ migration, movedDirectory }) => {
    return window.nativeBridge.file.operate({ action: 'move', filePath: migration,
      directoryPath: movedDirectory });
  }, { migration, movedDirectory });
  const movedDocument = await page.evaluate((filePath) => {
    return window.nativeBridge.file.openFromPath({ filePath });
  }, moved.filePath);
  const saved = path.join(savedDirectory, '另存副本.md');
  await application.evaluate(({ dialog }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath });
  }, saved);
  const savedDocument = await page.evaluate(({ movedDocument }) => {
    return window.nativeBridge.file.saveMarkdown({ filePath: null,
      sourceFilePath: movedDocument.filePath, content: movedDocument.content });
  }, { movedDocument });
  for (const filePath of [moved.filePath, saved]) {
    const content = await fs.readFile(filePath, 'utf8');
    const reference = /!\[本地图片\]\(([^)]+)\)/.exec(content)?.[1];
    assert.ok(reference);
    assert.deepEqual(await fs.readFile(path.resolve(path.dirname(filePath),
      decodeURIComponent(reference))), image);
  }
  assert.equal(savedDocument.content, await fs.readFile(saved, 'utf8'));
  assert.deepEqual(await fs.readFile(path.join(directory, '图片.png')), image);
  checks.push('真实 IPC 跨目录移动及另存为后图片可读，原图保留');

  if (evidenceDirectory) {
    const exportEvidence = path.join(evidenceDirectory, 'native-exports');
    await fs.mkdir(exportEvidence, { recursive: true });
    for (const name of await fs.readdir(directory)) {
      if (/\.(?:pdf|png|jpg|docx|odt|rtf|epub|tex)$/.test(name)) {
        await fs.copyFile(path.join(directory, name), path.join(exportEvidence, name));
      }
    }
  }
  const firstOrigin = new URL(page.url()).origin;
  await page.evaluate(async () => window.nativeBridge.session.save({
    filePath: '', content: '# 恢复草稿', isDirty: true, cursorHint: 1, sourceMode: true,
  }));
  await stopApplication();
  application = await launch();
  page = await application.firstWindow();
  // 关闭验收使用真实可见窗口，避免重启后的隐藏窗口进入后台节流状态。
  await application.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0].showInactive();
  });
  await expect(page.getByTestId('preview-host')).toContainText('恢复草稿');
  await expect(page.locator('.dirty-dot')).toHaveCount(1);
  assert.notEqual(new URL(page.url()).origin, firstOrigin);
  checks.push('跨随机端口的原生崩溃恢复');
  await application.evaluate(({ dialog }) => {
    dialog.showMessageBox = async () => ({ response: 1 });
  });
  if (evidenceDirectory) {
    await page.screenshot({ path: path.join(evidenceDirectory, 'native-recovery.png') });
  }
  const child = application.process();
  const exited = once(child, 'exit');
  await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close());
  let closeTimer;
  try {
    await Promise.race([exited, new Promise((_, reject) => {
      closeTimer = setTimeout(() => reject(new Error('正常关闭窗口超时')), 20_000);
    })]);
  } finally { clearTimeout(closeTimer); }
  application = undefined;
  assert.equal(child.exitCode, 0);
  assert.doesNotMatch(processDiagnostics, /native-smoke-uncaught|Object has been destroyed/);
  assert.deepEqual(rendererErrors, []);
  checks.push('正常关闭窗口无销毁对象异常，进程退出码为 0');
  console.log(JSON.stringify({ ok: true, checks }, null, 2));
  if (evidenceDirectory) {
    await fs.mkdir(evidenceDirectory, { recursive: true });
    await fs.writeFile(path.join(evidenceDirectory, 'native-result.json'),
      JSON.stringify({ ok: true, executablePath: executablePath || require('electron'),
        checks }, null, 2) + '\r\n', 'utf8');
  }
} catch (error) {
  console.error('原生验证失败', JSON.stringify({ checks, error: String(error) }));
  console.error('退出诊断', JSON.stringify({
    windows: application?.windows().map((page) => ({ closed: page.isClosed(), url: page.url() })),
    exitCode: application?.process().exitCode,
    processDiagnostics: processDiagnostics.slice(-12_000),
  }));
  if (evidenceDirectory && application) {
    await fs.mkdir(evidenceDirectory, { recursive: true });
    const page = application.windows()[0];
    if (page) {
      await page.screenshot({ path: path.join(evidenceDirectory, 'native-failure.png') })
        .catch(() => {});
      await fs.writeFile(path.join(evidenceDirectory, 'native-failure.html'),
        await page.content().catch(() => ''), 'utf8');
    }
  }
  throw error;
} finally {
  // 等待本次子进程退出，避免清理错误覆盖真正的功能断言失败。
  await stopApplication();
  const temporaryRoot = await fs.realpath(os.tmpdir());
  const resolvedDirectory = await fs.realpath(directory);
  assert.equal(path.dirname(resolvedDirectory), temporaryRoot);
  await fs.rm(resolvedDirectory, { recursive: true, force: true,
    maxRetries: 10, retryDelay: 200 });
}
