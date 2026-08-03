const MARKDOWN_FILE_RE = /\.(md|markdown)$/i;

function normalizeFileUrlToWindowsPath(rawUrl) {
  try {
    const parsed = new URL(rawUrl);
    if (parsed.protocol !== 'file:') return '';
    const decoded = decodeURIComponent(parsed.pathname || '');
    return /^\/[A-Za-z]:\//.test(decoded)
      ? decoded.slice(1).replace(/\//g, '\\')
      : decoded.replace(/\//g, '\\');
  } catch {
    return '';
  }
}

function normalizeCliMarkdownPath(input, options) {
  const { fsSync, path } = options;
  if (typeof input !== 'string') return '';
  const trimmed = input.trim();
  if (!trimmed || trimmed.startsWith('-')) return '';
  const dequoted = trimmed.replace(/^"+|"+$/g, '');
  const rawPath = /^file:\/\//i.test(dequoted)
    ? normalizeFileUrlToWindowsPath(dequoted)
    : dequoted;
  if (!rawPath) return '';
  const resolved = path.resolve(rawPath);
  if (!MARKDOWN_FILE_RE.test(resolved)) return '';
  try {
    return fsSync.statSync(resolved).isFile() ? resolved : '';
  } catch {
    return '';
  }
}

function resolveLaunchMarkdownPath(argv, options) {
  if (!Array.isArray(argv)) return '';
  for (const argument of argv.slice(1)) {
    const filePath = normalizeCliMarkdownPath(argument, options);
    if (filePath) return filePath;
  }
  return '';
}

function createLaunchService(options) {
  const { BrowserWindow, authorization, createWindow, fsSync, path } = options;
  const pending = [];
  const queues = new Map();

  function getPrimaryWindow() {
    return BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0] || null;
  }

  function focus(win) {
    if (!win || win.isDestroyed()) return;
    if (win.isMinimized()) win.restore();
    win.focus();
  }

  function enqueue(win, filePath) {
    if (!win || win.isDestroyed() || !filePath) return;
    const key = win.webContents.id;
    const queue = queues.get(key) || [];
    queue.push(filePath);
    queues.set(key, queue);
    authorization.grant(win, filePath);
  }

  function consume(win) {
    if (!win || win.isDestroyed()) return '';
    const key = win.webContents.id;
    const queue = queues.get(key) || [];
    const filePath = queue.shift() || '';
    if (queue.length === 0) queues.delete(key);
    return filePath;
  }

  function dispatch(win, filePath) {
    if (!win || win.isDestroyed() || !filePath || win.webContents.isLoadingMainFrame()) {
      return false;
    }
    win.webContents.send('app:launchFile', { filePath });
    return true;
  }

  function handle(filePath, handleOptions = {}) {
    if (!filePath) return null;
    if (handleOptions.openInNewWindow) {
      const win = createWindow(filePath);
      focus(win);
      return win;
    }
    const win = handleOptions.targetWindow || getPrimaryWindow();
    if (!win) {
      pending.push(filePath);
      return null;
    }
    enqueue(win, filePath);
    if (!dispatch(win, filePath)) focus(win);
    return win;
  }

  function resolve(argv) {
    return resolveLaunchMarkdownPath(argv, { fsSync, path });
  }

  return {
    addPending: (filePath) => pending.push(filePath),
    consumeForWindow: (win) => consume(win) || pending.shift() || '',
    consumePending: () => pending.shift() || '',
    enqueue,
    focus,
    getPending: () => [...pending],
    getPrimaryWindow,
    handle,
    removeWindow: (win) => {
      queues.delete(win.webContents.id);
      authorization.revoke(win);
    },
    resolve,
  };
}

module.exports = {
  createLaunchService,
  normalizeCliMarkdownPath,
  normalizeFileUrlToWindowsPath,
  resolveLaunchMarkdownPath,
};
