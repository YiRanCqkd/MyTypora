const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const { createIpcController } = require('../lib/ipc.cjs');

const expectedChannels = [
  'app:consumeLaunchFilePath',
  'app:getVersion',
  'app:logEvent',
  'context:showFileTreeMenu',
  'context:showSourceEditorMenu',
  'export:document',
  'export:getCapabilities',
  'export:html',
  'file:listMarkdownInDirectory',
  'file:openFromPath',
  'file:openMarkdown',
  'file:saveMarkdown',
  'settings:get',
  'settings:getAll',
  'settings:getRecentFiles',
  'settings:selectUserCssFile',
  'settings:set',
  'settings:setRecentFiles',
  'spell:getState',
  'spell:setEnabled',
  'spell:setLanguage',
  'window:close',
  'window:isMaximized',
  'window:minimize',
  'window:toggleDevTools',
  'window:toggleMaximize',
];

function createController() {
  const handlers = new Map();
  const passthrough = async (_event, payload) => payload;
  const win = {
    close: () => {},
    isMaximized: () => false,
    minimize: () => {},
    webContents: {
      closeDevTools: () => {},
      isDevToolsOpened: () => false,
      openDevTools: () => {},
    },
  };
  const controller = createIpcController({
    app: { getVersion: () => '1.0.0' },
    BrowserWindow: {
      fromWebContents: () => win,
      getFocusedWindow: () => win,
    },
    contextMenus: {
      showFileTreeMenu: passthrough,
      showSourceEditorMenu: passthrough,
    },
    exporters: {
      exportDocument: passthrough,
      exportHtml: passthrough,
      getCapabilities: passthrough,
    },
    files: {
      listMarkdownInDirectory: passthrough,
      openFromPath: passthrough,
      openMarkdown: passthrough,
      saveMarkdown: passthrough,
    },
    ipcMain: { handle: (channel, handler) => handlers.set(channel, handler) },
    launch: { consumeForWindow: () => 'D:\\docs\\launch.md' },
    runtime: {
      log: () => {},
      normalizeTraceId: (traceId) => traceId || 'trace-default',
    },
    settings: {
      get: passthrough,
      getAll: passthrough,
      getRecentFiles: passthrough,
      selectUserCssFile: passthrough,
      set: passthrough,
      setRecentFiles: passthrough,
    },
    spell: {
      getState: passthrough,
      setEnabled: passthrough,
      setLanguage: passthrough,
    },
  });
  controller.registerAll();
  return handlers;
}

test('IPC 装配保留 preload 使用的全部 channel 名', () => {
  const handlers = createController();
  const preload = fs.readFileSync(path.resolve(__dirname, '..', 'preload.cjs'), 'utf8');
  const preloadChannels = [...preload.matchAll(/ipcRenderer\.invoke\('([^']+)'/gu)]
    .map((match) => match[1]);

  assert.deepEqual([...handlers.keys()].sort(), expectedChannels);
  assert.deepEqual([...new Set(preloadChannels)].sort(), expectedChannels);
});

test('IPC 装配保留 payload、traceId 和返回结构', async () => {
  const handlers = createController();
  const event = { sender: { id: 7 } };

  assert.deepEqual(
    await handlers.get('settings:set')(event, {
      key: 'theme',
      value: 'dark',
      traceId: 'trace-settings',
    }),
    { key: 'theme', value: 'dark', traceId: 'trace-settings' },
  );
  assert.deepEqual(await handlers.get('app:getVersion')(event, {}), { version: '1.0.0' });
  assert.deepEqual(await handlers.get('app:consumeLaunchFilePath')(event, {}), {
    filePath: 'D:\\docs\\launch.md',
  });
});
