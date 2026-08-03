const { safeErrorMessage } = require('./errors.cjs');

function isExternalNavigationUrl(rawUrl) {
  if (typeof rawUrl !== 'string' || !rawUrl.trim()) return false;
  try {
    const protocol = new URL(rawUrl).protocol;
    return ['http:', 'https:', 'mailto:', 'tel:'].includes(protocol);
  } catch {
    return false;
  }
}

function createWindowService(options) {
  const {
    app,
    BrowserWindow,
    fsSync,
    isDev,
    launch,
    path,
    pathToFileURL,
    rendererServer,
    rootDirectory,
    runtime,
    settings,
    shell,
    spell,
  } = options;

  function resolveIconPath() {
    const packaged = path.join(process.resourcesPath, 'res', 'unireadmd.ico');
    if (app.isPackaged && fsSync.existsSync(packaged)) return packaged;
    return path.resolve(rootDirectory, 'res', 'unireadmd.ico');
  }

  function getRendererUrl() {
    if (isDev) return process.env.ELECTRON_START_URL;
    const serverUrl = rendererServer.getUrl();
    if (serverUrl) return serverUrl;
    return pathToFileURL(path.resolve(rootDirectory, 'frontend', 'dist', 'index.html')).toString();
  }

  function create(initialFilePath = '') {
    const preloadPath = path.join(rootDirectory, 'electron', 'preload.cjs');
    runtime.log('info', 'window preload path resolved', runtime.createTraceId('boot'), {
      preloadPath,
      preloadExists: fsSync.existsSync(preloadPath),
      isDev,
    });
    const win = new BrowserWindow({
      width: 1320,
      height: 880,
      minWidth: 980,
      minHeight: 640,
      title: 'UniReadMD',
      icon: resolveIconPath(),
      webPreferences: {
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: false,
        webSecurity: false,
        preload: preloadPath,
      },
    });
    const session = win.webContents.session;
    if (initialFilePath) launch.enqueue(win, initialFilePath);
    win.webContents.setWindowOpenHandler(({ url }) => {
      if (isExternalNavigationUrl(url)) void shell.openExternal(url);
      return { action: 'deny' };
    });
    win.webContents.on('will-navigate', (event, url) => {
      if (!isExternalNavigationUrl(url)) return;
      event.preventDefault();
      void shell.openExternal(url);
    });
    win.webContents.on('did-finish-load', () => {
      const traceId = runtime.createTraceId('boot');
      runtime.log('info', 'renderer did-finish-load', traceId);
      const command = process.env.UNIREADMD_AUTO_MENU_COMMAND?.trim();
      if (command) setTimeout(() => {
        runtime.log('info', 'auto menu command dispatch', traceId, { command });
        win.webContents.send('menu:command', command);
      }, 800);
    });
    win.webContents.on('did-fail-load', (_event, code, description, url, isMainFrame) => {
      runtime.log('error', 'renderer did-fail-load', runtime.createTraceId('boot'), {
        code,
        description,
        url,
        isMainFrame,
      });
    });
    win.webContents.on('preload-error', (_event, value, error) => {
      runtime.log('error', 'renderer preload-error', runtime.createTraceId('boot'), {
        preloadPath: value,
        error: safeErrorMessage(error),
      });
    });
    win.webContents.on('console-message', (_event, level, message, line, sourceId) => {
      if (level < 2) return;
      runtime.log('error', 'renderer console error', runtime.createTraceId('boot'), {
        level,
        message,
        line,
        sourceId,
      });
    });
    win.on('closed', () => launch.removeWindow(win));
    win.loadURL(getRendererUrl());
    void settings.readSettings()
      .then((current) => spell.applyInitial(session, current, runtime.createTraceId('spell')))
      .catch((error) => runtime.log(
        'error',
        'init spell check failed',
        runtime.createTraceId('spell'),
        { error: safeErrorMessage(error) },
      ));
    if (isDev) win.webContents.openDevTools({ mode: 'detach' });
    return win;
  }

  return { create, getRendererUrl, resolveIconPath };
}

module.exports = {
  createWindowService,
  isExternalNavigationUrl,
};
