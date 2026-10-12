const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { findDocumentImages, relocateDocumentImages } = require('../lib/document-images.cjs');
const { createAssetService } = require('../lib/assets.cjs');
const { createFilesService } = require('../lib/files.cjs');
const { createAuthorizationService } = require('../lib/authorization.cjs');
const { createWorkspaceTools } = require('../lib/workspace-tools.cjs');
const { createPrintOptions, validatePdfOptions } = require('../lib/pdf-options.cjs');
const { validateExportOptions, pandocOptionArguments } = require('../lib/export-options.cjs');
const { createExportsService } = require('../lib/exports.cjs');
const { createExportEngines } = require('../lib/export-engines.cjs');
const { createImageService, readClipboardImage } = require('../lib/images.cjs');
const { createContextMenuService } = require('../lib/context-menus.cjs');
const { collectDocumentImages } = require('../lib/image-collection.cjs');
const { createSettingsService, validateSettingUpdate } = require('../lib/settings.cjs');

test('写作设置保存到原生配置，新服务恢复且拒绝非法配置', async (context) => {
  const f = await fixture(context);
  const getFilePath = () => path.join(f.root, 'settings.json');
  const options = { fs, getFilePath, authorization: f.authorization };
  const preferences = { statUnit: 'characters', readingSpeed: 300, equationNumbering: true,
    lightTheme: 'sepia', darkTheme: 'midnight' };
  const service = createSettingsService(options);
  await service.set(f.event, { key: 'theme', value: 'system' });
  await service.set(f.event, { key: 'writingPreferencesJson', value: JSON.stringify(preferences) });
  const restored = await createSettingsService(options).getAll(f.event);
  assert.deepEqual(JSON.parse(restored.settings.writingPreferencesJson), preferences);
  assert.equal(restored.settings.theme, 'system');
  for (const value of [null, {}, 'null', '[]', '{', 'x'.repeat(2049),
    ...[{ readingSpeed: 0 }, { equationNumbering: 'yes' }, { statUnit: 'invalid' },
      { lightTheme: 'unknown' }, { darkTheme: 'unknown' }, { arbitraryPath: 'C:/' }]
      .map((patch) => JSON.stringify({ ...preferences, ...patch }))]) {
    assert.throws(() => validateSettingUpdate('writingPreferencesJson', value),
      { code: 'ERR_INVALID_SETTING_VALUE' });
  }
  assert.equal(validateSettingUpdate('writingPreferencesJson', '').value, '');
});

test('整理含井号和括号的路径及 ICO 图片后，资源服务可读回副本', async (context) => {
  const f = await fixture(context);
  const source = path.join(path.dirname(f.source), '#文档(1).md');
  const original = path.join(path.dirname(source), '#图.ico');
  const content = '![图标](%23%E5%9B%BE.ico)';
  await fs.writeFile(source, content);
  await fs.writeFile(original, 'icon-bytes');
  const result = await collectDocumentImages({ ...f, fs, source, content,
    destination: f.targetDirectory });
  const reference = findDocumentImages(result.content)[0].reference;
  assert.ok(reference.includes('%23'));
  assert.ok(reference.includes('%281%29'));
  assert.equal(path.extname(reference), '.ico');
  const token = new URL(f.assets.issue(f.event, source, f.root)).pathname.split('/').at(-1);
  const resource = await f.assets.readForOwner(f.event, token, decodeURIComponent(reference));
  assert.equal(resource.contentType, 'image/x-icon');
  assert.equal(resource.content.toString(), 'icon-bytes');
  assert.equal(await fs.readFile(original, 'utf8'), 'icon-bytes');
});

test('批量整理引用式与重复图片，副本可读，代码示例和原图保留', async (context) => {
  const f = await fixture(context);
  const content = f.content + '\n\n![重复][pic]\n\n[pic]: %E5%9B%BE.png';
  const result = await collectDocumentImages({ ...f, fs, source: f.source, content,
    destination: f.targetDirectory });
  assert.equal(result.count, 1);
  assert.ok(result.content.includes('`![示例](图.png)`'));
  const images = findDocumentImages(result.content);
  assert.equal(images[0].reference, images[1].reference);
  assert.equal(await fs.readFile(path.resolve(path.dirname(f.source),
    decodeURI(images[0].reference)), 'utf8'), 'image-bytes');
  assert.equal(await fs.readFile(path.join(path.dirname(f.source), '图.png'), 'utf8'), 'image-bytes');
  assert.equal(await fs.readFile(f.source, 'utf8'), f.content);
});

test('图片整理拒绝工作区之外及缺失资源，失败不留下新目录', async (context) => {
  const f = await fixture(context);
  await assert.rejects(collectDocumentImages({ ...f, fs, destination: os.tmpdir() }), /工作区/);
  await assert.rejects(collectDocumentImages({ ...f, fs, content: '![缺失](missing.png)',
    destination: f.targetDirectory }));
  assert.deepEqual(await fs.readdir(f.targetDirectory), []);
});

test('图片整理先验证窗口授权，再打开目录对话框', async (context) => {
  const f = await fixture(context);
  const service = createImageService({ ...f, fs, getEventWindow: () => null,
    dialog: { showOpenDialog: () => assert.fail('未授权调用不能打开目录对话框') } });
  await assert.rejects(service.collect({ sender: { id: 123456 } }, {
    sourceFilePath: f.source, content: f.content,
  }), { code: 'ERR_PATH_NOT_AUTHORIZED' });
});

test('导出目录与章节配置有范围约束，新格式使用固定参数', () => {
  const options = validateExportOptions({ tableOfContents: true, tocDepth: 4, epubChapterLevel: 2 });
  const epub = pandocOptionArguments(options, 'epub');
  assert.ok(epub.includes('--toc'));
  assert.deepEqual(epub.slice(-2), ['--split-level', '2']);
  assert.ok(pandocOptionArguments(options, 'revealjs').includes('--standalone'));
  assert.ok(pandocOptionArguments(options, 'opml').includes('--standalone'));
  assert.throws(() => validateExportOptions({ tocDepth: 7 }));
  assert.throws(() => validateExportOptions({ epubChapterLevel: '--lua-filter' }));
  assert.throws(() => validateExportOptions({ tableOfContents: 'yes' }));
});

test('旧版截图接口继续可用，异步文本剪贴板不创建图片', async () => {
  const image = { isEmpty: () => false };
  assert.equal(await readClipboardImage({ readImage: () => image }), image);
  assert.equal(await readClipboardImage({ read: async () => [{ types: ['text/plain'] }] }), null);
});

test('异步 ClipboardItem 图片通过截图入口写入资源目录，不依赖旧 readImage', async (context) => {
  const f = await fixture(context);
  const bytes = Buffer.from('真实图片字节');
  const service = createImageService({ ...f, fs,
    getEventWindow: () => null,
    clipboard: { read: async () => [{ types: ['image/png'],
      getType: async () => new Blob([bytes], { type: 'image/png' }) }] },
    nativeImage: { createFromBuffer: (buffer) => {
      assert.deepEqual(buffer, bytes);
      return { isEmpty: () => false, toPNG: () => buffer };
    } },
  });
  const result = await service.paste(f.event, { sourceFilePath: f.source });
  assert.equal(result.images.length, 1);
  assert.deepEqual(await fs.readFile(path.resolve(path.dirname(f.source),
    result.images[0].reference)), bytes);
});

test('异步截图拒绝超限载荷，空图片可以继续尝试下一种格式', async () => {
  let decoded = 0;
  const nativeImage = { createFromBuffer: () => {
    decoded += 1;
    return { isEmpty: () => true };
  } };
  const large = { read: async () => [{ types: ['image/png'], getType: async () => ({
    size: 21 * 1024 * 1024, arrayBuffer: () => { throw new Error('不应读取超限字节'); },
  }) }] };
  await assert.rejects(readClipboardImage(large, nativeImage), /超过 20 MiB/);
  assert.equal(decoded, 0);
  const empty = { read: async () => [{ types: ['image/png', 'image/jpeg'],
    getType: async () => new Blob(['empty']) }] };
  assert.equal(await readClipboardImage(empty, nativeImage), null);
  assert.equal(decoded, 2);
});

test('文件菜单等待异步剪贴板写入，失败记录日志而不产生未处理拒绝', async () => {
  let template;
  const writes = [];
  const logs = [];
  const clipboard = { writeText: async (text) => { writes.push(text); } };
  const service = createContextMenuService({ clipboard,
    Menu: { buildFromTemplate: (items) => { template = items; return { popup: () => {} }; } },
    getEventWindow: () => ({ isDestroyed: () => false }),
    log: (...args) => logs.push(args),
  });
  const filePath = path.resolve('中文文件.md');
  await service.showFileTreeMenu({}, { filePath }, 'clipboard-smoke');
  await template[0].click();
  await template[1].click();
  assert.deepEqual(writes, ['中文文件.md', filePath]);
  clipboard.writeText = async () => { throw new Error('系统剪贴板暂不可写'); };
  await template[0].click();
  assert.equal(logs.at(-1)[0], 'error');
  assert.match(logs.at(-1)[3].message, /系统剪贴板暂不可写/);
});

async function fixture(context) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-editing-'));
  context.after(() => fs.rm(root, { recursive: true, force: true }));
  const sourceDirectory = path.join(root, 'old');
  const targetDirectory = path.join(root, 'new');
  await fs.mkdir(sourceDirectory);
  await fs.mkdir(targetDirectory);
  const source = path.join(sourceDirectory, '文档.md');
  await fs.writeFile(path.join(sourceDirectory, '图.png'), 'image-bytes');
  const content = '# 文档\n\n![图片](%E5%9B%BE.png)\n\n`![示例](图.png)`';
  await fs.writeFile(source, content);
  const authorization = createAuthorizationService();
  const event = { sender: { id: 950 } };
  authorization.grantDirectory(event, root);
  const assets = createAssetService({ fs, getServerUrl: () => 'http://127.0.0.1:4173/' });
  assets.issue(event, source, root);
  const dialog = { showSaveDialog: async () => ({ filePath: path.join(targetDirectory, '另存为.md') }) };
  const files = createFilesService({ fs, assets, authorization, dialog, getEventWindow: () => null });
  const tools = createWorkspaceTools({ fs, files, assets, authorization, dialog,
    getEventWindow: () => null });
  return { root, source, targetDirectory, content, event, assets, authorization, files, tools };
}

test('图片源映射只选实际图片，支持引用式、列表和引用，代码示例保持原文', () => {
  const content = '> ![甲](a.png)\n\n- ![乙][id]\n\n[id]: b.png\n\n'
    + '`![假](x.png)`\n\n```md\n![假](y.png)\n```';
  const images = findDocumentImages(content);
  assert.deepEqual(images.map((item) => item.reference), ['a.png', 'b.png']);
  assert.equal(content.slice(images[0].from, images[0].to), '![甲](a.png)');
  assert.equal(content.slice(images[1].from, images[1].to), '![乙][id]');
});

test('跨目录移动复制已授权图片并更新引用，不删除原图与代码示例', async (context) => {
  const f = await fixture(context);
  const result = await f.tools.operate(f.event, {
    action: 'move', filePath: f.source, directoryPath: f.targetDirectory,
  });
  const content = await fs.readFile(result.filePath, 'utf8');
  assert.match(decodeURI(content), /文档-assets-/);
  assert.ok(content.includes('`![示例](图.png)`'));
  const reference = findDocumentImages(content)[0].reference;
  assert.equal(await fs.readFile(path.resolve(f.targetDirectory, decodeURI(reference)), 'utf8'),
    'image-bytes');
  assert.equal(await fs.readFile(path.join(path.dirname(f.source), '图.png'), 'utf8'), 'image-bytes');
  await assert.rejects(fs.stat(f.source), { code: 'ENOENT' });
});

test('另存为返回迁移后的正文，原文保留，新图在新文档能力下可读取', async (context) => {
  const f = await fixture(context);
  const result = await f.files.saveMarkdown(f.event, {
    sourceFilePath: f.source, content: f.content,
  });
  assert.ok(decodeURI(result.content).includes('另存为-assets-'));
  assert.equal(await fs.readFile(f.source, 'utf8'), f.content);
  const token = new URL(result.resourceBaseUrl).pathname.split('/').at(-1);
  assert.equal((await f.assets.readForOwner(f.event, token,
    decodeURI(findDocumentImages(result.content)[0].reference))).content.toString(), 'image-bytes');
});

test('图片缺失或越界时停止迁移，源文档不丢失且不留下目标文件', async (context) => {
  const f = await fixture(context);
  await fs.writeFile(f.source, '![缺失](missing.png)');
  await assert.rejects(f.tools.operate(f.event, {
    action: 'move', filePath: f.source, directoryPath: f.targetDirectory,
  }));
  assert.equal(await fs.readFile(f.source, 'utf8'), '![缺失](missing.png)');
  assert.deepEqual(await fs.readdir(f.targetDirectory), []);
  await assert.rejects(f.files.saveMarkdown({ sender: { id: 951 } }, {
    sourceFilePath: f.source, content: f.content,
  }), { code: 'ERR_PATH_NOT_AUTHORIZED' });
});

test('自定义 PDF 纸张及独立页边距转换成英寸，旧配置继续采用统一页边距', () => {
  const options = createPrintOptions({ pageSize: 'Custom', widthMm: 254, heightMm: 508,
    marginsMm: { top: 25.4, bottom: 12.7, left: 0, right: 5.08 } });
  assert.deepEqual(options.pageSize, { width: 10, height: 20 });
  assert.deepEqual(options.margins, { top: 1, bottom: 0.5, left: 0, right: 0.2 });
  assert.equal(createPrintOptions({ marginMm: 25.4 }).margins.left, 1);
  assert.throws(() => validatePdfOptions({ pageSize: 'Custom', widthMm: 0, heightMm: 210 }));
  assert.throws(() => validatePdfOptions({ marginsMm: { top: 1, bottom: 2, left: NaN, right: 0 } }));
});

test('导出设置采用固定 Pandoc 参数，不接受越界值或任意命令参数', () => {
  const settings = validateExportOptions({ title: '中文标题', author: '作者', language: 'zh-CN',
    referenceDoc: 'D:/参考.docx', epubCover: 'D:/封面.png', args: ['--lua-filter=evil'] });
  assert.deepEqual(pandocOptionArguments(settings, 'docx'), [
    '--metadata', 'title=中文标题', '--metadata', 'author=作者', '--metadata', 'lang=zh-CN',
    '--reference-doc', 'D:/参考.docx',
  ]);
  assert.ok(pandocOptionArguments(settings, 'epub').includes('--epub-cover-image'));
  assert.throws(() => validateExportOptions({ imageWidth: 100000 }));
  assert.throws(() => validateExportOptions({ jpegQuality: 101 }));
  assert.throws(() => validateExportOptions({ title: 'a\0b' }));
  assert.throws(() => validateExportOptions({ language: '--lua-filter' }));
});

test('未经原生授权的 Word 参考文件不能参与导出，验证失败前不选择输出目标', async (context) => {
  const f = await fixture(context);
  let saveDialog = false;
  const service = createExportsService({ fs, path, authorization: f.authorization,
    dialog: { showSaveDialog: async () => { saveDialog = true; return { canceled: true }; } },
    getEventWindow: () => null, engines: {} });
  await assert.rejects(service.exportDocument(f.event, { format: 'docx', html: '', markdown: '',
    options: { referenceDoc: path.join(os.tmpdir(), '未授权.docx') } }),
  { code: 'ERR_PATH_NOT_AUTHORIZED' });
  assert.equal(saveDialog, false);
});

test('图片宽度和 JPEG 质量传递给实际导出引擎，PDF 保持纸张设置', async (context) => {
  const f = await fixture(context);
  const seen = {};
  class Window {
    constructor(options) {
      seen.width = options.width;
      this.webContents = {
        executeJavaScript: async (script) => script.includes('Math.ceil')
          ? { width: seen.width, height: 200 } : true,
        capturePage: async (size) => {
          seen.size = size;
          return { toJPEG: (quality) => { seen.quality = quality; return Buffer.from('jpeg'); } };
        },
        printToPDF: async (options) => { seen.pdf = options; return Buffer.from('pdf'); },
      };
    }
    async loadURL() {}
    setContentSize() {}
    isDestroyed() { return false; }
    destroy() { seen.destroyed = true; }
  }
  const engines = createExportEngines({ fs, path, BrowserWindow: Window });
  await engines.exportChromium('<html></html>', 'jpeg', path.join(f.root, 'image.jpg'), undefined,
    { imageWidth: 2048, jpegQuality: 75 });
  assert.equal(seen.width, 2048);
  assert.equal(seen.size.width, 2048);
  assert.equal(seen.quality, 75);
  assert.equal(seen.destroyed, true);
  await engines.exportChromium('<html></html>', 'pdf', path.join(f.root, 'doc.pdf'),
    { pageSize: 'Custom', widthMm: 254, heightMm: 508 });
  assert.deepEqual(seen.pdf.pageSize, { width: 10, height: 20 });
});

test('Pandoc 引擎传递 Word 样式参考及元数据，原子写出并清理临时文件', async (context) => {
  const f = await fixture(context);
  const calls = [];
  const engines = createExportEngines({ fs, path, app: { getPath: () => f.root },
    runCommand: async (_command, args) => {
      calls.push(args);
      if (args[0] === '--version') return { stdout: 'pandoc 3.10' };
      if (args[0] === '--list-output-formats') return { stdout: 'docx epub' };
      await fs.writeFile(args[args.indexOf('--output') + 1], 'converted');
      return { stdout: '' };
    },
  });
  const target = path.join(f.root, 'result.docx');
  await engines.exportPandoc('# 正文', 'docx', target, 'test', '', undefined,
    { title: '中文', referenceDoc: 'D:/参考.docx' });
  assert.ok(calls.at(-1).includes('--reference-doc'));
  assert.ok(calls.at(-1).includes('title=中文'));
  assert.equal(await fs.readFile(target, 'utf8'), 'converted');
  assert.ok(!(await fs.readdir(f.root)).some((name) => name.startsWith('unireadmd-export-')));
});
