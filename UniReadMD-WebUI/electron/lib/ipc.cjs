const { safeErrorMessage } = require('./errors.cjs');

function createIpcController(options) {
  const {
    app,
    BrowserWindow,
    contextMenus,
    exporters,
    files,
    ipcMain,
    launch,
    runtime,
    settings,
    spell,
  } = options;

  function getEventWindow(event) {
    return BrowserWindow.fromWebContents(event.sender) || BrowserWindow.getFocusedWindow();
  }

  function register(channel, handler) {
    ipcMain.handle(channel, async (event, payload = {}) => {
      const traceId = runtime.normalizeTraceId(payload?.traceId);
      try {
        return await handler(event, payload, traceId);
      } catch (error) {
        runtime.log('error', `ipc failed: ${channel}`, traceId, {
          code: error?.code || 'ERR_IPC_FAILED',
          error: safeErrorMessage(error),
        });
        throw error;
      }
    });
  }

  function windowAction(event, payload, traceId, action) {
    const win = getEventWindow(event);
    if (!win) throw new Error('window not found');
    action(win);
    runtime.log('info', 'window action executed', traceId, payload ?? null);
  }

  function registerAll() {
    register('app:getVersion', async (_event, _payload, traceId) => {
      runtime.log('info', 'app getVersion', traceId);
      return { version: app.getVersion() };
    });
    register('app:logEvent', async (_event, payload, traceId) => {
      runtime.log(
        payload?.level === 'error' ? 'error' : 'info',
        typeof payload?.message === 'string' ? payload.message : 'renderer event',
        traceId,
        payload?.detail ?? null,
      );
      return { ok: true };
    });
    register('app:consumeLaunchFilePath', async (event, _payload, traceId) => {
      const filePath = launch.consumeForWindow(getEventWindow(event));
      runtime.log('info', 'consume launch file path', traceId, { filePath });
      return filePath ? { filePath } : null;
    });
    register('file:openMarkdown', files.openMarkdown);
    register('file:openFromPath', files.openFromPath);
    register('file:listMarkdownInDirectory', files.listMarkdownInDirectory);
    register('file:saveMarkdown', files.saveMarkdown);
    register('export:html', exporters.exportHtml);
    register('export:document', exporters.exportDocument);
    register('export:getCapabilities', exporters.getCapabilities);
    register('window:minimize', async (event, payload, traceId) => {
      windowAction(event, payload, traceId, (win) => win.minimize());
      return { ok: true };
    });
    register('window:toggleMaximize', async (event, payload, traceId) => {
      windowAction(event, payload, traceId, (win) => {
        if (win.isMaximized()) win.unmaximize();
        else win.maximize();
      });
      return { ok: true };
    });
    register('window:isMaximized', async (event, _payload, traceId) => {
      const isMaximized = Boolean(getEventWindow(event)?.isMaximized());
      runtime.log('info', 'window isMaximized', traceId, { isMaximized });
      return { isMaximized };
    });
    register('window:close', async (event, payload, traceId) => {
      windowAction(event, payload, traceId, (win) => win.close());
      return { ok: true };
    });
    register('window:toggleDevTools', async (event, payload, traceId) => {
      windowAction(event, payload, traceId, (win) => {
        if (win.webContents.isDevToolsOpened()) win.webContents.closeDevTools();
        else win.webContents.openDevTools({ mode: 'detach' });
      });
      return { ok: true };
    });
    register('settings:getAll', settings.getAll);
    register('settings:getRecentFiles', settings.getRecentFiles);
    register('settings:get', settings.get);
    register('settings:set', settings.set);
    register('settings:setRecentFiles', settings.setRecentFiles);
    register('settings:selectUserCssFile', settings.selectUserCssFile);
    register('context:showFileTreeMenu', contextMenus.showFileTreeMenu);
    register('context:showSourceEditorMenu', contextMenus.showSourceEditorMenu);
    register('spell:getState', spell.getState);
    register('spell:setEnabled', spell.setEnabled);
    register('spell:setLanguage', spell.setLanguage);
  }

  return { getEventWindow, registerAll };
}

function sendMenuCommandToWindow(target, command) {
  if (target && !target.isDestroyed()) target.webContents.send('menu:command', command);
}

function sendMenuCommand(BrowserWindow, command) {
  const target = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
  sendMenuCommandToWindow(target, command);
}

module.exports = {
  createIpcController,
  sendMenuCommand,
  sendMenuCommandToWindow,
};
