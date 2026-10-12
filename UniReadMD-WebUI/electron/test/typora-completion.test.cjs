const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const vm = require('node:vm');
const { createAuthorizationService } = require('../lib/authorization.cjs');
const { createFilesService } = require('../lib/files.cjs');
const { createAssetService } = require('../lib/assets.cjs');
const { createWorkspaceTools } = require('../lib/workspace-tools.cjs');
const { createImageService } = require('../lib/images.cjs');
const { createSessionService } = require('../lib/sessions.cjs');
const { createPrintOptions, validatePdfOptions } = require('../lib/pdf-options.cjs');
const { searchLines } = require('../lib/search-lines.cjs');
const { createExportsService } = require('../lib/exports.cjs');
const { createExportEngines } = require('../lib/export-engines.cjs');

async function fixture(context) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-parity-'));
  context.after(() => fs.rm(root, { recursive: true, force: true }));
  const event = { sender: { id: 901 } };
  const authorization = createAuthorizationService();
  const assets = createAssetService({ fs, getServerUrl: () => 'http://127.0.0.1:4173/' });
  let chosen = root;
  const dialog = { showOpenDialog: async () => ({ canceled: !chosen, filePaths: [chosen] }) };
  const files = createFilesService({ fs, authorization, assets, dialog,
    getEventWindow: () => null });
  await fs.mkdir(path.join(root, 'docs'));
  await fs.mkdir(path.join(root, 'images'));
  const document = path.join(root, 'docs', 'a.md');
  await fs.writeFile(document, '# 文档\n\n![图](../images/a.png)', 'utf8');
  await fs.writeFile(path.join(root, 'images', 'a.png'), Buffer.from('测试图片'));
  await files.openDirectory(event, {});
  const trash = [];
  const tools = createWorkspaceTools({ fs, files, assets, authorization, dialog,
    shell: { trashItem: async (target) => { trash.push(target); } }, getEventWindow: () => null });
  const images = createImageService({ fs, assets, files, authorization, dialog,
    clipboard: { readImage: () => ({ isEmpty: () => false, toPNG: () => Buffer.from('截图') }) },
    getEventWindow: () => null });
  return { root, event, authorization, assets, files, tools, images, document, trash,
    choose: (target) => { chosen = target; } };
}

test('预览与导出共用图片能力，不缩小工作区授权且拒绝其他窗口和越界', async (context) => {
  const f = await fixture(context);
  const document = await f.files.openFromPath(f.event, { filePath: f.document });
  const token = new URL(document.resourceBaseUrl).pathname.split('/').at(-1);
  assert.equal(f.assets.issue(f.event, f.document), document.resourceBaseUrl);
  assert.equal((await f.assets.readForOwner(f.event, token, '../images/a.png')).contentType,
    'image/png');
  await assert.rejects(f.assets.readForOwner({ sender: { id: 902 } }, token, '../images/a.png'));
  await assert.rejects(f.assets.read(token, '../../outside.png'));
  await assert.rejects(f.assets.read(token, '../docs/a.md'));
});

test('链接导航只授权工作区文档，首次系统拖入可打开文档', async (context) => {
  const f = await fixture(context);
  const target = path.join(f.root, '中文.md');
  await fs.writeFile(target, '# 新文档');
  const result = await f.tools.openLinkedFile(f.event, {
    sourceFilePath: f.document, href: '../%E4%B8%AD%E6%96%87.md#标题',
  });
  assert.equal(result.content, '# 新文档');
  await assert.rejects(f.tools.openLinkedFile(f.event, {
    sourceFilePath: f.document, href: 'https://example.com/a.md',
  }));
  const other = path.join(f.root, '临时', '拖入.md');
  await fs.mkdir(path.dirname(other));
  await fs.writeFile(other, '首次拖入');
  assert.equal((await f.tools.openDroppedFile({ sender: { id: 902 } }, {
    filePath: other,
  })).content, '首次拖入');
});

test('新建、复制、重命名、移动保留内容并拒绝覆盖、非法名称和移动到自身', async (context) => {
  const f = await fixture(context);
  const copy = await f.tools.operate(f.event, {
    action: 'duplicate', filePath: f.document, name: '副本.md',
  });
  assert.equal(await fs.readFile(copy.filePath, 'utf8'), await fs.readFile(f.document, 'utf8'));
  await assert.rejects(f.tools.operate(f.event, {
    action: 'rename', filePath: copy.filePath, name: 'a.md',
  }), { code: 'ERR_TARGET_EXISTS' });
  await assert.rejects(f.tools.operate(f.event, {
    action: 'rename', filePath: copy.filePath, name: '../逃逸.md',
  }), { code: 'ERR_INVALID_FILE_NAME' });
  const renamed = await f.tools.operate(f.event, {
    action: 'rename', filePath: copy.filePath, name: '改名.md',
  });
  await assert.rejects(fs.stat(copy.filePath), { code: 'ENOENT' });
  const folder = await f.tools.operate(f.event, {
    action: 'mkdir', filePath: f.root, name: '空目录',
  });
  assert.ok((await f.files.listMarkdownInDirectory(f.event, { filePath: f.document }))
    .directories.includes(folder.filePath));
  const moved = await f.tools.operate(f.event, {
    action: 'move', filePath: renamed.filePath, directoryPath: folder.filePath,
  });
  assert.ok((await fs.stat(moved.filePath)).isFile());
  await assert.rejects(f.tools.operate(f.event, {
    action: 'move', filePath: folder.filePath, directoryPath: folder.filePath,
  }));
  await assert.rejects(f.tools.operate(f.event, { action: 'trash', filePath: f.root }));
  const empty = await f.tools.operate(f.event, {
    action: 'create', filePath: folder.filePath, name: '新文档.md',
  });
  assert.equal(await fs.readFile(empty.filePath, 'utf8'), '');
  await f.tools.operate(f.event, { action: 'trash', filePath: empty.filePath });
  assert.deepEqual(f.trash, [empty.filePath]);
  f.choose(null);
  assert.equal(await f.tools.chooseMoveDirectory(f.event), null);
});

test('插图和截图复制到资源目录，取消无写入，拒绝越界目录', async (context) => {
  const f = await fixture(context);
  f.choose(path.join(f.root, 'images', 'a.png'));
  const inserted = await f.images.select(f.event, { sourceFilePath: f.document });
  assert.equal(inserted.images.length, 1);
  const copied = path.resolve(path.dirname(f.document), inserted.images[0].reference);
  assert.equal(await fs.readFile(copied, 'utf8'), '测试图片');
  const pasted = await f.images.paste(f.event, { sourceFilePath: f.document });
  assert.equal(await fs.readFile(path.resolve(path.dirname(f.document),
    pasted.images[0].reference), 'utf8'), '截图');
  await assert.rejects(f.images.paste(f.event, {
    sourceFilePath: f.document, copyDirectory: '../../越界',
  }), { code: 'ERR_RESOURCE_NOT_AUTHORIZED' });
  f.choose(null);
  assert.deepEqual(await f.images.select(f.event, { sourceFilePath: f.document }), { images: [] });
  assert.equal(await f.images.setRoot(f.event, { sourceFilePath: f.document }), null);
});

test('图片根目录仅通过已授权路径配置，导出后仍可读取根路径图片', async (context) => {
  const f = await fixture(context);
  const document = await f.files.openFromPath(f.event, { filePath: f.document });
  const token = new URL(document.resourceBaseUrl).pathname.split('/').at(-1);
  await f.images.applyRoot(f.event, {
    sourceFilePath: f.document, frontMatter: 'typora-root-url: ../images',
  });
  f.assets.issue(f.event, f.document);
  assert.equal((await f.assets.read(token, '/a.png')).content.toString(), '测试图片');
  await f.images.applyRoot(f.event, { sourceFilePath: f.document, frontMatter: '' });
  await assert.rejects(f.assets.read(token, '/a.png'));
});

test('历史草稿保留保存前内容，损坏记录不阻断列表，恢复副本不覆盖原文', async (context) => {
  const f = await fixture(context);
  const directory = path.join(f.root, 'sessions');
  const sessions = createSessionService({ fs, directory, authorization: f.authorization });
  await sessions.initialize();
  sessions.register({ webContents: { id: f.event.sender.id } }, false);
  await sessions.save(f.event, { snapshot: {
    filePath: f.document, content: '历史内容', isDirty: true,
  } });
  await sessions.save(f.event, { snapshot: {
    filePath: f.document, content: '新内容', isDirty: false,
  } });
  await fs.writeFile(path.join(directory, 'history', 'bad.json'), '{损坏');
  const { drafts } = await sessions.listDrafts();
  assert.ok(drafts.length >= 1);
  const restored = await sessions.restoreDraft(f.event, { id: drafts[0].id });
  assert.equal(restored.content, '历史内容');
  assert.equal(restored.filePath, '');
  assert.equal(restored.isDirty, true);
  assert.ok((await fs.readFile(f.document, 'utf8')).startsWith('# 文档'));
  await assert.rejects(sessions.restoreDraft(f.event, { id: '../../outside' }));
});

test('PDF 纸张、页边距和书签传递到打印，页眉脚 HTML 安全转义', () => {
  const result = createPrintOptions({ pageSize: 'Letter', marginMm: 25.4,
    landscape: true, bookmarks: false, header: '<img> {{page}} / {{pages}}' });
  assert.equal(result.pageSize, 'Letter');
  assert.equal(result.margins.left, 1);
  assert.equal(result.generateDocumentOutline, false);
  assert.match(result.headerTemplate, /&lt;img&gt;/);
  assert.match(result.headerTemplate, /class="pageNumber"/);
  assert.match(result.headerTemplate, /class="totalPages"/);
  for (const invalid of [null, [], { marginMm: -1 }, { landscape: 1 }, { pageSize: 'A0' }]) {
    assert.throws(() => validatePdfOptions(invalid));
  }
});

test('跨文件搜索支持中文全词、大小写和正则，危险回溯在限时线程终止', async () => {
  assert.deepEqual(await searchLines(['测试', '测试文档', 'x 测试 y'], '测试',
    { wholeWord: true }), [0, 2]);
  assert.deepEqual(await searchLines(['Abc', 'abc'], 'Abc', { caseSensitive: true }), [0]);
  assert.deepEqual(await searchLines(['foo1', 'foo2'], '^foo[12]$',
    { regularExpression: true }), [0, 1]);
  await assert.rejects(searchLines(['a'.repeat(40) + '!'], '(a+)+$',
    { regularExpression: true }), /超时|终止/);
});

test('隔离 preload 拒绝合成拖放、字符串路径和复用的拖放能力', async () => {
  const code = await fs.readFile(path.join(__dirname, '..', 'preload.cjs'), 'utf8');
  const systemFile = {};
  let bridge;
  let dropListener;
  const calls = [];
  const electron = {
    contextBridge: { exposeInMainWorld: (_name, value) => { bridge = value; } },
    ipcRenderer: { invoke: async (...args) => { calls.push(args); return {}; } },
    webUtils: { getPathForFile: (file) => {
      if (file !== systemFile) throw new TypeError('非法 File');
      return 'C:\\文档\\系统.md';
    } },
  };
  vm.runInNewContext(code, { require: () => electron, document: {
    addEventListener: (_name, handler) => { dropListener = handler; },
  }, Date, Map, console });
  dropListener({ isTrusted: false, dataTransfer: { files: [systemFile] } });
  await assert.rejects(bridge.file.openDroppedFile(systemFile));
  await assert.rejects(bridge.file.openDroppedFile('C:\\任意.md'));
  dropListener({ isTrusted: true, dataTransfer: { files: [systemFile] } });
  await bridge.file.openDroppedFile(systemFile);
  assert.equal(calls[0][0], 'file:openDroppedFile');
  await assert.rejects(bridge.file.openDroppedFile(systemFile));
});

test('Word 导出验证并传递工作区相对图片，结束后预览仍可读取原能力', async (context) => {
  const f = await fixture(context);
  const document = await f.files.openFromPath(f.event, { filePath: f.document });
  let resources;
  const exports = createExportsService({ fs, path, assets: f.assets, authorization: f.authorization,
    getEventWindow: () => null, dialog: {}, engines: {
      exportPandoc: async (...args) => { resources = args[5]; },
    } });
  const result = await exports.exportDocument(f.event, {
    format: 'docx', html: '<p>文档</p>', markdown: document.content,
    sourceFilePath: f.document, targetPath: path.join(f.root, 'Word.docx'),
  });
  assert.equal(result.engine, 'pandoc');
  assert.equal(resources.get('../images/a.png').content.toString(), '测试图片');
  const token = new URL(document.resourceBaseUrl).pathname.split('/').at(-1);
  assert.equal((await f.assets.read(token, '../images/a.png')).content.toString(), '测试图片');
});

test('Pandoc AST 转换重定向根目录图片，保留代码中的同名路径', async (context) => {
  const f = await fixture(context);
  let source;
  const engines = createExportEngines({ fs, path, app: { getPath: () => f.root },
    runCommand: async (_command, args) => {
      if (args[0] === '--version') return { stdout: 'pandoc 3.10' };
      if (args[0] === '--list-output-formats') return { stdout: 'docx epub latex' };
      if (args[args.indexOf('--to') + 1] === 'json') return { stdout: JSON.stringify({
        blocks: [{ t: 'Para', c: [{ t: 'Image', c: [[], [], ['/a.png', '']] }] },
          { t: 'CodeBlock', c: [[], '/a.png'] }],
      }) };
      source = JSON.parse(await fs.readFile(args[0], 'utf8'));
      assert.equal(await fs.readFile(source.blocks[0].c[0].c[2][0], 'utf8'), '已授权图片');
      await fs.writeFile(args[args.indexOf('--output') + 1], '转换产物');
      return { stdout: '' };
    },
  });
  await engines.exportPandoc('![图](/a.png)\n\n`/a.png`', 'docx',
    path.join(f.root, '结果.docx'), '测试', f.document, new Map([
      ['/a.png', { content: Buffer.from('已授权图片'), contentType: 'image/png' }],
    ]));
  assert.equal(source.blocks[1].c[1], '/a.png');
  assert.ok((await engines.getCapabilities()).pandocFormats.includes('epub'));
  assert.ok((await engines.getCapabilities()).pandocFormats.includes('latex'));
});

test('多格式导入提取图片并生成相对引用，失败保留已有文档', async (context) => {
  const f = await fixture(context);
  let fail = false;
  const engines = createExportEngines({ fs, path, runCommand: async (_command, args) => {
    if (args[0] === '--version') return { stdout: 'pandoc 3.10' };
    if (args[0] === '--list-output-formats') return { stdout: 'docx epub latex' };
    if (fail) throw new Error('模拟损坏文档');
    const directory = args[args.indexOf('--extract-media') + 1];
    await fs.writeFile(path.join(directory, 'picture.png'), '提取图片');
    return { stdout: `![图片](${directory.replace(/\\/g, '/')}/picture.png)` };
  } });
  const target = path.join(f.root, '导入.md');
  await engines.importPandoc(path.join(f.root, '源.docx'), target);
  const markdown = await fs.readFile(target, 'utf8');
  assert.match(markdown, /^!\[图片\]\(import-assets-/);
  const reference = /\((.+)\)/.exec(markdown)[1];
  assert.equal(await fs.readFile(path.resolve(f.root, reference), 'utf8'), '提取图片');
  fail = true;
  await assert.rejects(engines.importPandoc(path.join(f.root, '损坏.docx'), target));
  assert.equal(await fs.readFile(target, 'utf8'), markdown);
});

test('经原生选择的图片根目录可跨重启恢复，其他文档不能借用权限', async (context) => {
  const f = await fixture(context);
  const outside = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-image-root-'));
  context.after(() => fs.rm(outside, { recursive: true, force: true }));
  await fs.writeFile(path.join(outside, 'a.png'), '根目录图片');
  const approvalsPath = path.join(f.root, 'approvals.json');
  const selected = createImageService({ fs, assets: f.assets, files: f.files,
    authorization: f.authorization, approvalsPath, getEventWindow: () => null,
    dialog: { showOpenDialog: async () => ({ canceled: false, filePaths: [outside] }) },
  });
  await selected.setRoot(f.event, { sourceFilePath: f.document });
  const authorization = createAuthorizationService();
  authorization.grant(f.event, f.document);
  const assets = createAssetService({ fs, getServerUrl: () => 'http://127.0.0.1/' });
  const token = new URL(assets.issue(f.event, f.document)).pathname.split('/').at(-1);
  const restored = createImageService({ fs, assets, authorization, approvalsPath,
    files: { getWorkspaceRoot: () => undefined },
  });
  await restored.applyRoot(f.event, {
    sourceFilePath: f.document, frontMatter: `typora-root-url: ${outside}`,
  });
  assert.equal((await assets.read(token, '/a.png')).content.toString(), '根目录图片');
  const other = path.join(f.root, 'docs', 'other.md');
  authorization.grant(f.event, other);
  await assert.rejects(restored.applyRoot(f.event, {
    sourceFilePath: other, frontMatter: `typora-root-url: ${outside}`,
  }), { code: 'ERR_RESOURCE_NOT_AUTHORIZED' });
});

test('LaTeX 图片保存到产物旁的相对目录，失败清理本次资源并保留旧文件', async (context) => {
  const f = await fixture(context);
  let fail = false;
  const engines = createExportEngines({ fs, path, app: { getPath: () => f.root },
    runCommand: async (_command, args) => {
      if (args[0] === '--version') return { stdout: 'pandoc 3.10' };
      if (args[0] === '--list-output-formats') return { stdout: 'latex' };
      if (args[args.indexOf('--to') + 1] === 'json') return { stdout: JSON.stringify({
        blocks: [{ t: 'Image', c: [[], [], ['../images/a.png', '']] }],
      }) };
      if (fail) throw new Error('模拟转换失败');
      const ast = JSON.parse(await fs.readFile(args[0], 'utf8'));
      const reference = ast.blocks[0].c[2][0];
      await fs.writeFile(args[args.indexOf('--output') + 1], reference);
      return { stdout: '' };
    },
  });
  const target = path.join(f.root, '文档.tex');
  const resources = new Map([['../images/a.png', {
    content: Buffer.from('有效图片'), contentType: 'image/png',
  }]]);
  await engines.exportPandoc('![图](../images/a.png)', 'latex', target, '', '', resources);
  const reference = await fs.readFile(target, 'utf8');
  assert.equal(path.isAbsolute(reference), false);
  assert.equal(await fs.readFile(path.resolve(f.root, reference), 'utf8'), '有效图片');
  const before = await fs.readdir(f.root);
  fail = true;
  await assert.rejects(engines.exportPandoc('图片', 'latex', target, '', '', resources));
  assert.equal(await fs.readFile(target, 'utf8'), reference);
  assert.deepEqual(await fs.readdir(f.root), before);
});
