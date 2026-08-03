const assert = require('node:assert/strict');
const fs = require('node:fs');
const fsPromises = require('node:fs/promises');
const path = require('node:path');
const test = require('node:test');

const { createAuthorizationService } = require('../lib/authorization.cjs');
const {
  createExportsService,
  deriveExportPath,
  getExportDialogFilter,
  validateExportDocumentPayload,
} = require('../lib/exports.cjs');
const { createFilesService } = require('../lib/files.cjs');
const {
  createSettingsService,
  mergeSettings,
  normalizeRecentFiles,
} = require('../lib/settings.cjs');
const { buildApplicationMenuTemplate } = require('../lib/menus.cjs');
const { normalizeCliMarkdownPath } = require('../lib/launch.cjs');

const event = { sender: { id: 7 } };
const fakeWindow = {
  isDestroyed: () => false,
  isMinimized: () => false,
  focus: () => {},
};

function createDirectory(prefix) {
  return fs.mkdtempSync(`D:\\DevTools\\${prefix}-`);
}

test('设置适配器保留默认值、有效覆盖、持久化和最近文件合同', async () => {
  const directory = createDirectory('unireadmd-task3-settings');
  const settingsPath = path.join(directory, 'settings.json');
  const authorization = createAuthorizationService();
  const service = createSettingsService({
    authorization,
    dialog: {},
    fs: fsPromises,
    getEventWindow: () => fakeWindow,
    getFilePath: () => settingsPath,
  });

  try {
    const merged = mergeSettings({ theme: 'dark', fontSize: 15.6, contentZoom: 125.4 });
    assert.equal(merged.theme, 'dark');
    assert.equal(merged.fontSize, 16);
    assert.equal(merged.contentZoom, 125);
    assert.deepEqual(
      normalizeRecentFiles(['A.md', 'a.md', 'note.txt', ' B.markdown ']),
      ['A.md', 'B.markdown'],
    );
    assert.equal((await service.getAll(event, {}, 'trace')).settings.theme, 'light');
    const updated = await service.set(event, { key: 'theme', value: 'dark' }, 'trace');
    assert.equal(updated.value, 'dark');
    assert.equal(updated.settings.theme, 'dark');
    assert.deepEqual(
      await service.get(event, { key: 'theme' }, 'trace'),
      { key: 'theme', value: 'dark' },
    );
    assert.equal(JSON.parse(await fsPromises.readFile(settingsPath, 'utf8')).theme, 'dark');
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

test('文件适配器保留打开、保存、目录列表和取消返回合同', async () => {
  const directory = createDirectory('unireadmd-task3-file');
  const sourcePath = path.join(directory, 'source.md');
  const siblingPath = path.join(directory, 'sibling.markdown');
  fs.writeFileSync(sourcePath, '# 当前合同\r\n', 'utf8');
  fs.writeFileSync(siblingPath, '兄弟文件', 'utf8');
  const authorization = createAuthorizationService();
  const service = createFilesService({
    authorization,
    dialog: {
      showOpenDialog: async () => ({ canceled: false, filePaths: [sourcePath] }),
      showSaveDialog: async () => ({ canceled: true }),
    },
    enforceAuthorization: true,
    fs: fsPromises,
    getEventWindow: () => fakeWindow,
  });

  try {
    assert.deepEqual(await service.openMarkdown(event, {}, 'trace-open'), {
      filePath: sourcePath,
      content: '# 当前合同\r\n',
    });
    assert.deepEqual(
      await service.saveMarkdown(
        event,
        { filePath: sourcePath, content: '已保存' },
        'trace-save',
      ),
      { filePath: sourcePath },
    );
    assert.equal(await fsPromises.readFile(sourcePath, 'utf8'), '已保存');
    const listing = await service.listMarkdownInDirectory(
      event,
      { filePath: sourcePath },
      'trace-list',
    );
    assert.equal(listing.directoryPath, directory);
    assert.deepEqual(new Set(listing.filePaths), new Set([sourcePath, siblingPath]));

    const canceled = createFilesService({
      authorization,
      dialog: { showOpenDialog: async () => ({ canceled: true, filePaths: [] }) },
      fs: fsPromises,
      getEventWindow: () => fakeWindow,
    });
    assert.equal(await canceled.openMarkdown(event, {}, 'trace-cancel'), null);
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

test('授权输入保留仅接受真实 Markdown 启动文件的合同', () => {
  const directory = createDirectory('unireadmd-task3-launch');
  const markdownPath = path.join(directory, 'launch.md');
  const textPath = path.join(directory, 'launch.txt');
  fs.writeFileSync(markdownPath, 'launch', 'utf8');
  fs.writeFileSync(textPath, 'text', 'utf8');

  try {
    assert.equal(normalizeCliMarkdownPath(markdownPath, { fsSync: fs, path }), markdownPath);
    assert.equal(normalizeCliMarkdownPath(textPath, { fsSync: fs, path }), '');
    assert.equal(
      normalizeCliMarkdownPath(path.join(directory, 'missing.md'), { fsSync: fs, path }),
      '',
    );
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

test('导出适配器保留参数转换、HTML 落盘和返回结构', async () => {
  const directory = createDirectory('unireadmd-task3-export');
  const targetPath = path.join(directory, 'document.html');
  const authorization = createAuthorizationService();
  authorization.grant(event, targetPath);
  const service = createExportsService({
    authorization,
    dialog: {},
    enforceAuthorization: true,
    engines: {},
    fs: fsPromises,
    getEventWindow: () => fakeWindow,
    path,
  });

  try {
    assert.equal(deriveExportPath('D:\\docs\\readme.md', 'jpeg', path), 'readme.jpg');
    assert.deepEqual(getExportDialogFilter('pdf'), { name: 'PDF', extensions: ['pdf'] });
    const payload = validateExportDocumentPayload({
      format: 'HTML',
      html: '<h1>标题</h1>',
      markdown: '# 标题',
      targetPath,
    });
    assert.equal(payload.format, 'html');
    assert.deepEqual(await service.exportDocument(event, payload, 'trace-export'), {
      filePath: targetPath,
      format: 'html',
      engine: 'builtin',
    });
    assert.deepEqual(await service.exportHtml(event, {
      html: '<h1>标题</h1>',
      targetPath,
    }, 'trace-html'), {
      filePath: targetPath,
      format: 'html',
      engine: 'builtin',
    });
    assert.equal(await fsPromises.readFile(targetPath, 'utf8'), '<h1>标题</h1>');
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

test('菜单模型保留标签、快捷键和渲染器命令合同', () => {
  const commands = [];
  const template = buildApplicationMenuTemplate({
    isMac: false,
    sendCommand: (command) => commands.push(command),
  });
  const fileMenu = template.find((item) => item.label === '文件');
  const editMenu = template.find((item) => item.label === '编辑');
  const exportMenu = fileMenu.submenu.find((item) => item.label === '导出');

  assert.equal(fileMenu.submenu[0].accelerator, 'CmdOrCtrl+O');
  assert.deepEqual(
    exportMenu.submenu.filter((item) => item.label).map((item) => item.label),
    ['导出 HTML...', '导出 PDF...', '导出 PNG...', '导出 JPEG...', '导出 Word (docx)...'],
  );
  fileMenu.submenu[0].click();
  editMenu.submenu.find((item) => item.label === '查找').click();
  assert.deepEqual(commands, ['file-open', 'edit-search']);
});
