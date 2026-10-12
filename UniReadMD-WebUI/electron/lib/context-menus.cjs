const path = require('node:path');

const { requireString } = require('./errors.cjs');

function createContextMenuService(options) {
  const { Menu, clipboard, getEventWindow, log, sendCommandToWindow, shell,
    authorization } = options;

  function popupOptions(win, payload) {
    const x = Number.isFinite(payload?.x) ? Math.max(0, Math.round(payload.x)) : null;
    const y = Number.isFinite(payload?.y) ? Math.max(0, Math.round(payload.y)) : null;
    return {
      window: win,
      ...(x === null ? {} : { x }),
      ...(y === null ? {} : { y }),
    };
  }

  async function showFileTreeMenu(event, payload, traceId) {
    const filePath = requireString(payload?.filePath, 'filePath');
    authorization?.assertAuthorized(event, filePath);
    const win = getEventWindow(event);
    if (!win || win.isDestroyed()) throw new Error('window is not available');
    const fileName = path.basename(filePath);
    const menu = Menu.buildFromTemplate([
      {
        label: '复制文件名',
        click: async () => {
          try {
            await clipboard.writeText(fileName);
            log('info', 'context menu copied file name', traceId, { fileName });
          } catch (error) {
            log('error', 'context menu copy failed', traceId, { message: String(error) });
          }
        },
      },
      {
        label: '复制文件全路径',
        click: async () => {
          try {
            await clipboard.writeText(filePath);
            log('info', 'context menu copied file path', traceId, { filePath });
          } catch (error) {
            log('error', 'context menu copy failed', traceId, { message: String(error) });
          }
        },
      },
      { type: 'separator' },
      {
        label: '在资源管理器中显示',
        click: () => shell.showItemInFolder(path.normalize(filePath)),
      },
    ]);
    menu.popup(popupOptions(win, payload));
    log('info', 'context menu opened', traceId, { filePath });
    return { ok: true };
  }

  async function showSourceEditorMenu(event, payload, traceId) {
    const win = getEventWindow(event);
    if (!win || win.isDestroyed()) throw new Error('window is not available');
    const action = (name) => () => win.webContents[name]();
    const menu = Menu.buildFromTemplate([
      { label: '剪切', click: action('cut') },
      { label: '复制', click: action('copy') },
      { label: '粘贴', click: action('paste') },
      { type: 'separator' },
      { label: '全选', click: action('selectAll') },
      { label: '查找', click: () => sendCommandToWindow(win, 'edit-search') },
    ]);
    menu.popup(popupOptions(win, payload));
    log('info', 'source context menu opened', traceId);
    return { ok: true };
  }

  return { showFileTreeMenu, showSourceEditorMenu };
}

module.exports = { createContextMenuService };
