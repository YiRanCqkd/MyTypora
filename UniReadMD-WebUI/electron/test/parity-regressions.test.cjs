const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const { EventEmitter } = require('node:events');
const { createAuthorizationService } = require('../lib/authorization.cjs');
const { createFilesService } = require('../lib/files.cjs');
const { createAssetService } = require('../lib/assets.cjs');
const { createSessionService } = require('../lib/sessions.cjs');
const { validateIpcSender } = require('../lib/security.cjs');
const { sanitizeCss } = require('../lib/safe-css.cjs');
const { prepareExportHtml } = require('../lib/export-html.cjs');
const { createExportEngines } = require('../lib/export-engines.cjs');
const { createExportsService } = require('../lib/exports.cjs');
const { buildApplicationMenuTemplate } = require('../lib/menus.cjs');
const { createWindowService } = require('../lib/windows.cjs');
const { createLaunchService } = require('../lib/launch.cjs');

test('窗口销毁后按稳定标识清理图片、授权、启动队列和会话，保留其他窗口', async (context) => {
  await withDirectory(async (directory) => {
    const originalResourcesPath = Object.getOwnPropertyDescriptor(process, 'resourcesPath');
    Object.defineProperty(process, 'resourcesPath', { configurable: true, value: directory });
    context.after(() => {
      if (originalResourcesPath) {
        Object.defineProperty(process, 'resourcesPath', originalResourcesPath);
      } else Reflect.deleteProperty(process, 'resourcesPath');
    });
    let nextId = 200;
    class TestWindow extends EventEmitter {
      constructor() {
        super();
        this.destroyed = false;
        this.contents = Object.assign(new EventEmitter(), {
          id: ++nextId,
          session: {
            setPermissionRequestHandler() {},
            setPermissionCheckHandler() {},
          },
          setWindowOpenHandler() {},
        });
      }
      get webContents() {
        if (this.destroyed) throw new TypeError('Object has been destroyed');
        return this.contents;
      }
      isDestroyed() { return this.destroyed; }
      loadURL() {}
    }
    const authorization = createAuthorizationService();
    const assets = createAssetService({ fs, getServerUrl: () => 'http://127.0.0.1:1/' });
    const sessions = createSessionService({ fs, directory, authorization });
    const launch = createLaunchService({ BrowserWindow: TestWindow, authorization, fsSync: {}, path });
    const logs = [];
    const windows = createWindowService({
      app: { isPackaged: false },
      BrowserWindow: TestWindow,
      fsSync: { existsSync: () => true },
      isDev: false,
      launch,
      path,
      rendererServer: { getUrl: () => 'http://127.0.0.1:1/' },
      rootDirectory: directory,
      runtime: { log: (...args) => logs.push(args), createTraceId: () => '关闭回归' },
      settings: { readSettings: async () => ({}) },
      shell: {},
      spell: { applyInitial() {} },
      sessions,
      assets,
    });
    const filePath = path.join(directory, '窗口一.md');
    const otherPath = path.join(directory, '窗口二.md');
    const first = windows.create(filePath);
    const second = windows.create(otherPath);
    const firstId = first.webContents.id;
    const secondId = second.webContents.id;
    const token = new URL(assets.issue(first, filePath)).pathname.split('/').at(-1);
    const otherToken = new URL(assets.issue(second, otherPath)).pathname.split('/').at(-1);
    const saving = sessions.save({ sender: { id: firstId } }, {
      snapshot: { filePath, content: '关闭前最后一次修改', isDirty: true },
    });
    first.destroyed = true;
    assert.doesNotThrow(() => first.emit('closed'));
    await saving;
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(await sessions.load({ sender: { id: firstId } }), null);
    assert.equal(authorization.isAuthorized(firstId, filePath), false);
    assert.equal(authorization.isAuthorized(secondId, otherPath), true);
    assert.equal(launch.consumeForWindow({
      webContents: { id: firstId }, isDestroyed: () => false,
    }), '');
    assert.equal(launch.consumeForWindow(second), otherPath);
    await assert.rejects(assets.read(token, '图片.png'),
      { code: 'ERR_RESOURCE_NOT_AUTHORIZED' });
    await fs.writeFile(path.join(directory, '图片.png'), Buffer.from([1]));
    assert.deepEqual((await assets.read(otherToken, '图片.png')).content, Buffer.from([1]));
    const snapshots = (await fs.readdir(directory)).filter((name) => name.endsWith('.json'));
    assert.equal(snapshots.length, 1);
    assert.equal(JSON.parse(await fs.readFile(path.join(directory, snapshots[0]), 'utf8')).content,
      '关闭前最后一次修改');
    assert.equal(logs.some((entry) => entry[1] === 'session release failed'), false);

    // 保持现有传入窗口对象的清理接口兼容。
    assets.revoke(second);
    await sessions.release(second);
    launch.removeWindow(second);
    assert.equal(authorization.isAuthorized(secondId, otherPath), false);
    await assert.rejects(assets.read(otherToken, '图片.png'),
      { code: 'ERR_RESOURCE_NOT_AUTHORIZED' });
    second.destroyed = true;
    assert.doesNotThrow(() => second.emit('closed'));
  });
});

async function withDirectory(executor) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-regression-'));
  try { await executor(directory); }
  finally { await fs.rm(directory, { recursive: true, force: true }); }
}

test('默认拒绝未授权路径，不能读取或覆盖其他文件', async () => {
  const calls = [];
  const service = createFilesService({
    authorization: createAuthorizationService(),
    dialog: {},
    getEventWindow: () => null,
    fs: {
      readFile: async () => calls.push('read'),
      writeFile: async () => calls.push('write'),
    },
  });
  const event = { sender: { id: 101 } };
  const filePath = path.join(os.tmpdir(), '未授权.md');
  await assert.rejects(
    service.openFromPath(event, { filePath }),
    { code: 'ERR_PATH_NOT_AUTHORIZED' },
  );
  await assert.rejects(
    service.saveMarkdown(event, { filePath, content: '内容' }),
    { code: 'ERR_PATH_NOT_AUTHORIZED' },
  );
  assert.deepEqual(calls, []);
});

test('图片能力隔离窗口，拒绝目录越界和目录联接越界', async () => {
  await withDirectory(async (directory) => {
    const root = path.join(directory, '文档');
    const outside = path.join(directory, '外部');
    await fs.mkdir(root);
    await fs.mkdir(outside);
    await fs.writeFile(path.join(root, '图片.png'), Buffer.from([1, 2, 3]));
    await fs.writeFile(path.join(outside, '秘密.png'), Buffer.from([4]));
    await fs.symlink(outside, path.join(root, '联接'), 'junction');
    const assets = createAssetService({ fs, getServerUrl: () => 'http://127.0.0.1:1/' });
    const owner = { sender: { id: 103 } };
    const other = { sender: { id: 104 } };
    const url = new URL(assets.issue(owner, path.join(root, '文章.md')));
    const token = url.pathname.split('/').at(-1);
    assert.deepEqual((await assets.readForOwner(owner, token, '图片.png')).content,
      Buffer.from([1, 2, 3]));
    for (const reference of ['../外部/秘密.png', '联接/秘密.png']) {
      await assert.rejects(assets.readForOwner(owner, token, reference),
        { code: 'ERR_RESOURCE_NOT_AUTHORIZED' });
    }
    await assert.rejects(assets.readForOwner(other, token, '图片.png'),
      { code: 'ERR_RESOURCE_NOT_AUTHORIZED' });
    assets.revoke(owner);
    await assert.rejects(assets.read(token, '图片.png'));
  });
});

test('会话按窗口独立恢复，损坏快照和超限写入不破坏有效快照', async () => {
  await withDirectory(async (directory) => {
    const authorization = createAuthorizationService();
    const options = { fs, directory, authorization, dialog: {}, getEventWindow: () => null };
    const sessions = createSessionService(options);
    await sessions.initialize();
    const first = { webContents: { id: 105 } };
    const second = { webContents: { id: 106 } };
    sessions.register(first, false);
    sessions.register(second, false);
    await sessions.save({ sender: first.webContents }, {
      snapshot: { filePath: '', content: '窗口一', isDirty: true },
    });
    await sessions.save({ sender: second.webContents }, {
      snapshot: { filePath: '', content: '窗口二', isDirty: true },
    });
    await assert.rejects(sessions.save({ sender: first.webContents }, {
      snapshot: { filePath: '', content: 'x'.repeat(25 * 1024 * 1024) },
    }), { code: 'ERR_SESSION_TOO_LARGE' });
    await fs.writeFile(path.join(directory, 'deadbeef.json'), '{损坏');
    await fs.writeFile(path.join(directory, 'abcdef.json'), 'x'.repeat(25 * 1024 * 1024 + 1));
    const restored = createSessionService(options);
    await restored.initialize();
    assert.equal(restored.getRecoveryCount(), 2);
    restored.register(first, false);
    restored.register(second, false);
    const snapshots = await Promise.all([first, second].map((win) => {
      return restored.load({ sender: win.webContents });
    }));
    assert.deepEqual(snapshots.map((item) => item.content).sort(), ['窗口一', '窗口二']);
    assert.ok(snapshots.every((item) => item.isDirty));
  });
});

test('递归工作区忽略目录联接环，正文搜索可定位嵌套文件', async () => {
  await withDirectory(async (directory) => {
    const nested = path.join(directory, '子目录');
    await fs.mkdir(nested);
    await fs.writeFile(path.join(nested, '文档.md'), '# 标题\n中文搜索词');
    await fs.writeFile(path.join(directory, '程序.cjs'), '禁止覆盖');
    await fs.symlink(directory, path.join(nested, '环'), 'junction');
    const event = { sender: { id: 107 } };
    const files = createFilesService({ fs, authorization: createAuthorizationService(),
      getEventWindow: () => null,
      dialog: { showOpenDialog: async () => ({ canceled: false, filePaths: [directory] }) },
    });
    const result = await files.openDirectory(event);
    assert.deepEqual(result.filePaths, [await fs.realpath(path.join(nested, '文档.md'))]);
    assert.equal((await files.searchDirectory(event, { query: '中文' })).matches[0].line, 2);
    await assert.rejects(files.saveMarkdown(event, {
      filePath: path.join(directory, '程序.cjs'), content: '覆盖程序',
    }), { code: 'ERR_UNSUPPORTED_DOCUMENT' });
    assert.equal(await fs.readFile(path.join(directory, '程序.cjs'), 'utf8'), '禁止覆盖');
    assert.equal((await files.openDirectory(event, { directoryPath: directory })).directoryPath,
      await fs.realpath(directory));
  });
});

test('IPC 拒绝子框架、远程页面和未知窗口', () => {
  const frame = { url: 'http://127.0.0.1:1234/', parent: null };
  const sender = { id: 108, mainFrame: frame };
  const win = { isDestroyed: () => false, webContents: sender };
  const windows = { fromWebContents: () => win };
  assert.doesNotThrow(() => validateIpcSender({ sender, senderFrame: frame }, windows,
    () => frame.url));
  assert.throws(() => validateIpcSender({ sender,
    senderFrame: { url: frame.url, parent: frame },
  }, windows, () => frame.url), { code: 'ERR_IPC_SENDER' });
  frame.url = 'https://example.com';
  assert.throws(() => validateIpcSender({ sender, senderFrame: frame }, windows,
    () => 'http://127.0.0.1:1234/'), { code: 'ERR_IPC_SENDER' });
  assert.throws(() => validateIpcSender({ sender, senderFrame: frame },
    { fromWebContents: () => null }, () => frame.url), { code: 'ERR_IPC_SENDER' });
});

test('用户 CSS 限于预览，导出过滤脚本并保留公式、样式和图表', async () => {
  const css = sanitizeCss('@import "https://example.com"; body { color: red; }'
    + '.test { background: url(file:///secret); }');
  assert.match(css, /\.preview/);
  assert.doesNotMatch(css, /@import|url\(/);
  const boundaries = sanitizeCss('.preview + .toolbar { color:red }'
    + '.preview-evil { color:blue } .preview p + p { margin:2px }'
    + '.text { background:u\\72l(https://example.com) }');
  assert.doesNotMatch(boundaries, /toolbar|u\\72l/);
  assert.match(boundaries, /\.preview \.preview-evil/);
  assert.match(boundaries, /\.preview p \+ p/);
  assert.throws(() => sanitizeCss('body {'), { code: 'ERR_INVALID_CSS' });
  const embeddedFonts = '@font-face{font-family:export-font;src:url(data:font/woff2;base64,'
    + 'A'.repeat(1024 * 1024) + ')}';
  assert.throws(() => sanitizeCss(embeddedFonts), { code: 'ERR_INVALID_CSS' });
  assert.match(sanitizeCss(embeddedFonts, { scoped: false }), /data:font\/woff2;base64/);
  const html = await prepareExportHtml('<title>中文标题</title><style>.preview{color:red}</style>'
    + '<article class="preview"><math><mi>x</mi></math><svg><path d="M0 0L1 1" /></svg>'
    + '<img src="x" onerror="alert(1)"><script>alert(2)</script></article>', {});
  assert.match(html, /中文标题/);
  assert.match(html, /color:red/);
  assert.match(html, /<math>/);
  assert.match(html, /<svg>/);
  assert.match(html, /Content-Security-Policy/);
  assert.doesNotMatch(html, /onerror|<script/);
});

test('导出拒绝格式不一致的目标，Pandoc 拒绝文档目录之外的资源', async () => {
  await withDirectory(async (directory) => {
    const documentDirectory = path.join(directory, '文档');
    await fs.mkdir(documentDirectory);
    const sourceFilePath = path.join(documentDirectory, '文章.md');
    const targetPath = path.join(directory, '结果.docx');
    const program = path.join(directory, '程序.cjs');
    await fs.writeFile(sourceFilePath, '# 文章');
    await fs.writeFile(program, '原程序');
    await fs.writeFile(path.join(directory, '秘密.png'), Buffer.from([1]));
    const event = { sender: { id: 111 } };
    const authorization = createAuthorizationService();
    authorization.grantDirectory(event, directory);
    const assets = createAssetService({ fs, getServerUrl: () => 'http://127.0.0.1:1/' });
    let calls = 0;
    const exports = createExportsService({ fs, path, authorization, assets, dialog: {},
      getEventWindow: () => null,
      engines: { exportPandoc: async () => { calls += 1; } },
    });
    await assert.rejects(exports.exportDocument(event, {
      format: 'html', html: '<p>覆盖</p>', markdown: '', targetPath: program,
    }), { code: 'ERR_EXPORT_TARGET_EXTENSION' });
    await assert.rejects(exports.exportDocument(event, {
      format: 'docx', html: '', markdown: '![秘密](../秘密.png)', sourceFilePath, targetPath,
    }), { code: 'ERR_RESOURCE_NOT_AUTHORIZED' });
    assert.equal(calls, 0);
    assert.equal(await fs.readFile(program, 'utf8'), '原程序');
  });
});

test('长文图片导出捕获全文，超尺寸拒绝并销毁临时窗口', async () => {
  await withDirectory(async (directory) => {
    const heights = [2400, 40000];
    const captures = [];
    let destroyed = 0;
    class ExportWindow {
      constructor() {
        const height = heights.shift();
        this.webContents = {
          executeJavaScript: async (script) => script.includes('scrollHeight')
            ? { width: 1280, height } : true,
          capturePage: async (size) => {
            captures.push(size);
            return { toPNG: () => Buffer.from('完整图片') };
          },
        };
      }
      async loadURL() {}
      setContentSize() {}
      isDestroyed() { return false; }
      destroy() { destroyed += 1; }
    }
    const engines = createExportEngines({ fs, path, BrowserWindow: ExportWindow });
    const target = path.join(directory, '全文.png');
    await engines.exportChromium('<p>全文</p>', 'png', target);
    assert.equal(captures[0].height, 2400);
    assert.equal(await fs.readFile(target, 'utf8'), '完整图片');
    await assert.rejects(engines.exportChromium('<p>超限</p>', 'png', target));
    assert.equal(destroyed, 2);
    assert.equal(await fs.readFile(target, 'utf8'), '完整图片');
  });
});

test('原子保存失败保留原文件，且拒绝外部修改冲突', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'unireadmd-parity-'));
  const filePath = path.join(directory, '原文.md');
  await fs.writeFile(filePath, '原始内容');
  const authorization = createAuthorizationService();
  const event = { sender: { id: 102 } };
  const dialog = { showOpenDialog: async () => ({ canceled: false, filePaths: [filePath] }) };
  try {
    const service = createFilesService({
      authorization, dialog, fs, getEventWindow: () => null,
    });
    await service.openMarkdown(event, {});
    await fs.writeFile(filePath, '外部修改');
    await assert.rejects(
      service.saveMarkdown(event, { filePath, content: '编辑内容' }),
      { code: 'ERR_DOCUMENT_CONFLICT' },
    );
    assert.equal(await fs.readFile(filePath, 'utf8'), '外部修改');
    await service.openMarkdown(event, {});
    const failing = createFilesService({
      authorization, dialog, getEventWindow: () => null,
      fs: { ...fs, rename: async () => { throw new Error('模拟磁盘错误'); } },
    });
    await failing.openMarkdown(event, {});
    await assert.rejects(failing.saveMarkdown(event, { filePath, content: '新内容' }));
    assert.equal(await fs.readFile(filePath, 'utf8'), '外部修改');
    assert.deepEqual(await fs.readdir(directory), ['原文.md']);
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test('Pandoc 缺失或版本不匹配时拒绝导出并禁用菜单', async () => {
  for (const version of [null, 'pandoc 3.9']) {
    const engines = createExportEngines({ runCommand: async () => {
      if (version === null) throw new Error('模拟未安装');
      return { stdout: version };
    } });
    assert.equal((await engines.getCapabilities()).pandocAvailable, false);
    await assert.rejects(engines.exportPandoc('', 'docx', ''),
      { code: 'ERR_PANDOC_UNAVAILABLE' });
  }
  const fileMenu = buildApplicationMenuTemplate({ pandocFormats: [] })[0];
  const entries = fileMenu.submenu.find((entry) => entry.label === '导出').submenu;
  assert.ok(entries.filter((entry) => /docx|odt|RTF/.test(entry.label || ''))
    .every((entry) => entry.enabled === false));
});

test('Pandoc 3.10 合同传入文档资源目录，三种格式均原子写入目标', async () => {
  await withDirectory(async (directory) => {
    const commands = [];
    const engines = createExportEngines({ fs, path,
      app: { getPath: () => directory },
      runCommand: async (_command, args) => {
        if (args[0] === '--version') return { stdout: 'pandoc 3.10\n' };
        if (args[0] === '--list-output-formats') return { stdout: 'docx\nodt\nrtf\n' };
        commands.push(args);
        await fs.writeFile(args[args.indexOf('--output') + 1], Buffer.from('模拟转换产物'));
        return { stdout: '' };
      },
    });
    assert.deepEqual((await engines.getCapabilities()).pandocFormats, ['docx', 'odt', 'rtf']);
    for (const format of ['docx', 'odt', 'rtf']) {
      const target = path.join(directory, `结果.${format}`);
      await engines.exportPandoc('# 内容', format, target, '测试', path.join(directory, '原文.md'));
      assert.equal(await fs.readFile(target, 'utf8'), '模拟转换产物');
    }
    assert.ok(commands.every((args) => args.includes('markdown-raw_html-raw_tex')
      && args[args.indexOf('--resource-path') + 1] === directory));
    assert.ok((await fs.readdir(directory)).every((name) => name.startsWith('结果.')));
  });
});
