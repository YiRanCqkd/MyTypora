const MAX_CONTENT_BYTES = 20 * 1024 * 1024;

const path = require('node:path');

const { createContractError, requireString } = require('./errors.cjs');

function validateSavePayload(payload) {
  const content = typeof payload?.content === 'string' ? payload.content : null;
  if (content === null) {
    throw createContractError('ERR_INVALID_ARGUMENT', 'content 必须是字符串');
  }
  if (Buffer.byteLength(content, 'utf8') > MAX_CONTENT_BYTES) {
    throw createContractError('ERR_DOCUMENT_TOO_LARGE', '文档超过 20 MiB 限制');
  }

  const filePath = payload?.filePath === undefined || payload?.filePath === null
    ? null
    : requireString(payload.filePath, 'filePath');

  return { filePath, content };
}

function isMarkdownFileName(fileName) {
  return /\.(md|markdown)$/i.test(String(fileName || ''));
}

function createFilesService(options) {
  const {
    authorization,
    dialog,
    enforceAuthorization = false,
    fs,
    getEventWindow,
    log = () => {},
  } = options;

  function assertAuthorized(event, filePath) {
    if (enforceAuthorization) authorization.assertAuthorized(event, filePath);
  }

  function focusForDialog(win) {
    if (!win || win.isDestroyed()) return;
    if (win.isMinimized()) win.restore();
    win.focus();
  }

  async function openMarkdown(event, _payload, traceId) {
    const win = getEventWindow(event);
    focusForDialog(win);
    const dialogOptions = {
      properties: ['openFile'],
      filters: [{ name: 'Markdown', extensions: ['md', 'markdown', 'txt'] }],
    };
    const result = win && !win.isDestroyed()
      ? await dialog.showOpenDialog(win, dialogOptions)
      : await dialog.showOpenDialog(dialogOptions);
    if (result.canceled || result.filePaths.length === 0) {
      log('info', 'open markdown canceled', traceId);
      return null;
    }
    const filePath = result.filePaths[0];
    authorization.grant(event, filePath);
    const content = await fs.readFile(filePath, 'utf8');
    log('info', 'open markdown success', traceId, { filePath });
    return { filePath, content };
  }

  async function openFromPath(event, payload, traceId) {
    const filePath = requireString(payload?.filePath, 'filePath');
    assertAuthorized(event, filePath);
    const content = await fs.readFile(filePath, 'utf8');
    log('info', 'open markdown from path success', traceId, { filePath });
    return { filePath, content };
  }

  async function listMarkdownInDirectory(event, payload, traceId) {
    const filePath = requireString(payload?.filePath, 'filePath');
    assertAuthorized(event, filePath);
    const directoryPath = path.dirname(filePath);
    const entries = await fs.readdir(directoryPath, { withFileTypes: true });
    const filePaths = entries
      .filter((entry) => entry.isFile() && isMarkdownFileName(entry.name))
      .map((entry) => path.join(directoryPath, entry.name))
      .sort((left, right) => path.basename(left).localeCompare(path.basename(right), 'zh-CN'));
    authorization.grantMany(event, filePaths);
    log('info', 'list markdown in directory success', traceId, {
      directoryPath,
      count: filePaths.length,
    });
    return { directoryPath, filePaths };
  }

  async function saveMarkdown(event, payload, traceId) {
    const validated = validateSavePayload(payload);
    let targetPath = validated.filePath;
    if (targetPath) {
      assertAuthorized(event, targetPath);
    } else {
      const win = getEventWindow(event);
      focusForDialog(win);
      const dialogOptions = {
        filters: [{ name: 'Markdown', extensions: ['md'] }],
      };
      const result = win && !win.isDestroyed()
        ? await dialog.showSaveDialog(win, dialogOptions)
        : await dialog.showSaveDialog(dialogOptions);
      if (result.canceled || !result.filePath) {
        log('info', 'save markdown canceled', traceId);
        return null;
      }
      targetPath = result.filePath;
      authorization.grant(event, targetPath);
    }
    await fs.writeFile(targetPath, validated.content, 'utf8');
    log('info', 'save markdown success', traceId, { filePath: targetPath });
    return { filePath: targetPath };
  }

  return {
    listMarkdownInDirectory,
    openFromPath,
    openMarkdown,
    saveMarkdown,
  };
}

module.exports = {
  MAX_CONTENT_BYTES,
  createFilesService,
  isMarkdownFileName,
  validateSavePayload,
};
