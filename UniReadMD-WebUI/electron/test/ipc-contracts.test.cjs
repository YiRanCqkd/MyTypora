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
  'document:confirmDiscard',
  'export:chooseResource',
  'export:document',
  'export:getCapabilities',
  'export:html',
  'export:importDocument',
  'file:chooseMoveDirectory',
  'file:inspect',
  'file:listMarkdownInDirectory',
  'file:openDirectory',
  'file:openDroppedFile',
  'file:openFromPath',
  'file:openLinkedFile',
  'file:openMarkdown',
  'file:operate',
  'file:saveMarkdown',
  'file:searchDirectory',
  'image:applyRoot',
  'image:collect',
  'image:drop',
  'image:paste',
  'image:select',
  'image:setRoot',
  'session:listDrafts',
  'session:load',
  'session:restoreDraft',
  'session:save',
  'settings:get',
  'settings:getAll',
  'settings:getRecentFiles',
  'settings:getUserCss',
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
      chooseResource: passthrough,
      importDocument: passthrough,
      exportDocument: passthrough,
      exportHtml: passthrough,
      getCapabilities: passthrough,
    },
    files: {
      openDirectory: passthrough,
      searchDirectory: passthrough,
      listMarkdownInDirectory: passthrough,
      openFromPath: passthrough,
      openMarkdown: passthrough,
      saveMarkdown: passthrough,
    },
    ipcMain: { handle: (channel, handler) => handlers.set(channel, handler) },
    workspaceTools: {
      openDroppedFile: passthrough, openLinkedFile: passthrough, operate: passthrough,
      inspect: passthrough, chooseMoveDirectory: passthrough,
    },
    images: {
      select: passthrough, paste: passthrough, drop: passthrough,
      setRoot: passthrough, applyRoot: passthrough, collect: passthrough,
    },
    launch: { consumeForWindow: () => 'D:\\docs\\launch.md' },
    runtime: {
      log: () => {},
      normalizeTraceId: (traceId) => traceId || 'trace-default',
    },
    settings: {
      getUserCss: passthrough,
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
    sessions: {
      listDrafts: passthrough,
      restoreDraft: passthrough,
      load: passthrough,
      save: passthrough,
      confirmDiscard: passthrough,
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
  const imagePayload = {
    sourceFilePath: 'D:\\docs\\source.md', content: '![图](图.png)', traceId: 'trace-images',
  };
  assert.deepEqual(await handlers.get('image:collect')(event, imagePayload), imagePayload);
  assert.deepEqual(await handlers.get('app:consumeLaunchFilePath')(event, {}), {
    filePath: 'D:\\docs\\launch.md',
  });
});
