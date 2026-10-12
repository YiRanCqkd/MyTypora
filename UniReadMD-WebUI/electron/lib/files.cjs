const MAX_CONTENT_BYTES = 20 * 1024 * 1024;

const path = require('node:path');
const { createHash } = require('node:crypto');
const { atomicWrite } = require('./atomic-write.cjs');
const { toAuthorizationKey } = require('./authorization.cjs');
const { relocateDocumentImages } = require('./document-images.cjs');

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
    enforceAuthorization = true,
    assets,
    fs,
    getEventWindow,
    log = () => {},
  } = options;
  const versions = new Map();
  const roots = new Map();
  const writes = new Map();
  const versionKey = (event, filePath) => `${event.sender.id}:${toAuthorizationKey(filePath)}`;
  const hashContent = (content) => createHash('sha256').update(content).digest('hex');

  async function readDocument(event, filePath) {
    assertDocumentPath(filePath);
    const stat = await fs.stat(filePath);
    if (!stat.isFile() || stat.size > MAX_CONTENT_BYTES) {
      throw createContractError('ERR_DOCUMENT_TOO_LARGE', '文档超过 20 MiB 限制');
    }
    const content = await fs.readFile(filePath, 'utf8');
    const version = hashContent(content);
    versions.set(versionKey(event, filePath), version);
    const resourceBaseUrl = assets?.issue(event, filePath, roots.get(event.sender.id));
    const frontMatter = /^---\r?\n([\s\S]*?)\r?\n---/.exec(content)?.[1] || '';
    const imageRoot = /^typora-root-url:\s*(.+)$/m.exec(frontMatter)?.[1]?.trim()
      .replace(/^['"]|['"]$/g, '');
    if (assets && imageRoot) {
      const requestedRoot = path.resolve(path.dirname(filePath), imageRoot);
      if (authorization.isAuthorized(event, requestedRoot)) {
        assets.setImageRoot(event, filePath, requestedRoot);
      }
    }
    return {
      filePath, content,
      ...(assets ? {
        version,
        resourceBaseUrl,
      } : {}),
    };
  }

  function assertAuthorized(event, filePath) {
    if (enforceAuthorization) authorization.assertAuthorized(event, filePath);
  }

  function focusForDialog(win) {
    if (!win || win.isDestroyed()) return;
    if (win.isMinimized()) win.restore();
    win.focus();
  }

  function assertDocumentPath(filePath) {
    if (enforceAuthorization && !/\.(md|markdown|txt)$/i.test(filePath)) {
      throw createContractError('ERR_UNSUPPORTED_DOCUMENT', '仅支持 Markdown 和纯文本文档');
    }
  }

  async function directoryResult(directoryPath, filePaths, directories = []) {
    if (!assets) return { directoryPath, filePaths };
    const modifiedTimes = {};
    for (let offset = 0; offset < filePaths.length; offset += 32) {
      await Promise.all(filePaths.slice(offset, offset + 32).map(async (filePath) => {
        try { modifiedTimes[filePath] = (await fs.stat(filePath)).mtimeMs; }
        catch { /* 文件在扫描中被移动时，在下次刷新时移除。 */ }
      }));
    }
    return { directoryPath, filePaths, modifiedTimes, directories };
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
    const document = await readDocument(event, filePath);
    log('info', 'open markdown success', traceId, { filePath });
    return document;
  }

  async function openFromPath(event, payload, traceId) {
    const filePath = requireString(payload?.filePath, 'filePath');
    assertAuthorized(event, filePath);
    const document = await readDocument(event, filePath);
    log('info', 'open markdown from path success', traceId, { filePath });
    return document;
  }

  async function listMarkdownInDirectory(event, payload, traceId) {
    const filePath = requireString(payload?.filePath, 'filePath');
    assertAuthorized(event, filePath);
    const directoryPath = roots.get(event.sender.id) || path.dirname(filePath);
    const directories = [];
    const filePaths = await scanDirectory(directoryPath, directories);
    authorization.grantMany(event, filePaths);
    log('info', 'list markdown in directory success', traceId, {
      directoryPath,
      count: filePaths.length,
    });
    return directoryResult(directoryPath, filePaths, directories);
  }

  async function scanDirectory(root, directories = []) {
    const filePaths = [];
    let visited = 0;
    async function visit(directory, depth) {
      if (depth > 12 || ++visited > 2000 || filePaths.length >= 10000) return;
      let entries;
      try {
        entries = await fs.readdir(directory, { withFileTypes: true });
      } catch (error) {
        if (directory === root) throw error;
        return;
      }
      for (const entry of entries) {
        if (entry.isSymbolicLink() || filePaths.length >= 10000) continue;
        const target = path.join(directory, entry.name);
        if (entry.isDirectory() && !['node_modules', '.git'].includes(entry.name)) {
          directories.push(target);
          await visit(target, depth + 1);
        } else if (entry.isFile() && isMarkdownFileName(entry.name)) {
          filePaths.push(target);
        }
      }
    }
    await visit(root, 0);
    return filePaths.sort((left, right) => left.localeCompare(right, 'zh-CN', { numeric: true }));
  }

  async function openDirectory(event, payload) {
    let result;
    if (payload?.directoryPath) {
      assertAuthorized(event, requireString(payload.directoryPath, 'directoryPath'));
      result = { canceled: false, filePaths: [payload.directoryPath] };
    } else result = await dialog.showOpenDialog(getEventWindow(event), {
      properties: ['openDirectory'],
    });
    if (result.canceled || !result.filePaths[0]) return null;
    const directoryPath = await fs.realpath(result.filePaths[0]);
    authorization.grantDirectory(event, directoryPath);
    roots.set(event.sender.id, directoryPath);
    const directories = [];
    const filePaths = await scanDirectory(directoryPath, directories);
    authorization.grantMany(event, filePaths);
    return directoryResult(directoryPath, filePaths, directories);
  }

  async function searchDirectory(event, payload) {
    const directory = roots.get(event.sender.id);
    if (!directory) return { matches: [] };
    const query = requireString(payload?.query, 'query', 200);
    const { searchLines, buildPattern } = require('./search-lines.cjs');
    const searchOptions = {
      caseSensitive: payload.caseSensitive === true,
      wholeWord: payload.wholeWord === true,
      regularExpression: payload.regularExpression === true,
    };
    try { buildPattern(query, searchOptions); }
    catch { throw new Error('正则表达式无效'); }
    const matches = [];
    let truncated = false;
    let scannedBytes = 0;
    const started = Date.now();
    for (const filePath of await scanDirectory(directory)) {
      if (matches.length >= 200 || scannedBytes >= 64 * 1024 * 1024
        || Date.now() - started > 10000) { truncated = true; break; }
      try {
        const stat = await fs.stat(filePath);
        if (stat.size > MAX_CONTENT_BYTES) continue;
        scannedBytes += stat.size;
        const lines = (await fs.readFile(filePath, 'utf8')).split(/\r?\n/);
        const indices = await searchLines(lines, query, searchOptions, 201 - matches.length);
        for (const index of indices) {
          if (matches.length >= 200) { truncated = true; break; }
          matches.push({ filePath, line: index + 1, text: lines[index].slice(0, 240) });
        }
      } catch (error) {
        if (!['ENOENT', 'EACCES', 'EPERM'].includes(error.code)) throw error;
        // 扫描中消失或无权限的文件不阻断其余结果。
      }
    }
    return { matches, truncated };
  }

  async function saveMarkdown(event, payload, traceId) {
    const validated = validateSavePayload(payload);
    const source = payload.sourceFilePath;
    if (source) assertAuthorized(event, requireString(source, 'sourceFilePath'));
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
    const key = versionKey(event, targetPath);
    assertDocumentPath(targetPath);
    const expected = payload.version || versions.get(key);
    const writeKey = toAuthorizationKey(targetPath);
    let savedContent = validated.content;
    const pending = (writes.get(writeKey) || Promise.resolve()).catch(() => {}).then(async () => {
      if (expected) {
        let current;
        try {
          current = hashContent(await fs.readFile(targetPath, 'utf8'));
        } catch {
          throw createContractError('ERR_DOCUMENT_CONFLICT', '原文已被移动或删除，请另存为');
        }
        if (current !== expected) {
          throw createContractError('ERR_DOCUMENT_CONFLICT', '原文已被外部修改，请另存为或重新打开');
        }
      }
      const relocated = source && !validated.filePath
        ? await relocateDocumentImages({ content: savedContent, source, target: targetPath,
          event, assets, fs, rootDirectory: roots.get(event.sender.id) })
        : { content: savedContent, cleanup: async () => {} };
      try {
        await atomicWrite(fs, targetPath, relocated.content);
        savedContent = relocated.content;
        versions.set(key, hashContent(savedContent));
      } catch (error) { await relocated.cleanup(); throw error; }
    });
    writes.set(writeKey, pending);
    try { await pending; }
    finally { if (writes.get(writeKey) === pending) writes.delete(writeKey); }
    log('info', 'save markdown success', traceId, { filePath: targetPath });
    return {
      filePath: targetPath,
      ...(savedContent !== validated.content ? { content: savedContent } : {}),
      ...(assets ? {
        version: versions.get(key),
        resourceBaseUrl: assets.issue(event, targetPath, roots.get(event.sender.id)),
      } : {}),
    };
  }

  return {
    listMarkdownInDirectory,
    openFromPath,
    openMarkdown,
    openDirectory,
    searchDirectory,
    saveMarkdown,
    getWorkspaceRoot: (event) => roots.get(event.sender.id),
  };
}

module.exports = {
  MAX_CONTENT_BYTES,
  createFilesService,
  isMarkdownFileName,
  validateSavePayload,
};
