const { app, BrowserWindow, Menu, clipboard, dialog, ipcMain, shell } = require('electron');
const crypto = require('node:crypto');
const { spawn } = require('node:child_process');
const http = require('node:http');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const isDev = Boolean(process.env.ELECTRON_START_URL);
const MAX_CONTENT_BYTES = 20 * 1024 * 1024;
const MARKDOWN_FILE_RE = /\.(md|markdown)$/i;
const DEFAULT_APP_NAME = 'UniReadMD';
const DEFAULT_APP_DESCRIPTION = 'UniReadMD Markdown 阅读器';
const DEFAULT_APP_PROG_ID = 'UniReadMD.AssocFile.Markdown';
const DEFAULT_APP_CAPABILITIES_PATH = 'Software\\UniReadMD\\Capabilities';
const DEFAULT_APP_ICON_FILE_NAME = 'UniReadMD.ico';

const DEFAULT_SETTINGS = {
  theme: 'light',
  fontSize: 14,
  contentZoom: 100,
  viewMode: 'split',
  autoSave: false,
  userCssEnabled: false,
  userCssPath: '',
  spellCheckEnabled: true,
  spellCheckLanguage: 'en-US',
};

const ALLOWED_SETTINGS = new Set(Object.keys(DEFAULT_SETTINGS));

let runtimeLogFilePath = '';
let settingsFilePath = '';
let settingsCache = null;
let recentFilesCache = null;
let rendererStaticServer = null;
let rendererStaticServerUrl = '';
let pandocProbeCache = null;
let pendingLaunchFilePaths = [];
const windowLaunchFileQueues = new Map();

const BUILTIN_EXPORT_FORMATS = ['html', 'pdf', 'png', 'jpeg'];
const PANDOC_EXPORT_FORMATS = ['docx', 'odt', 'rtf'];
const SUPPORTED_EXPORT_FORMATS = new Set([
  ...BUILTIN_EXPORT_FORMATS,
  ...PANDOC_EXPORT_FORMATS,
]);

function isMarkdownFilePath(filePath) {
  return MARKDOWN_FILE_RE.test(String(filePath || ''));
}

function normalizeFileUrlToWindowsPath(rawUrl) {
  try {
    const parsed = new URL(rawUrl);
    if (parsed.protocol !== 'file:') {
      return '';
    }

    const decodedPath = decodeURIComponent(parsed.pathname || '');
    if (/^\/[A-Za-z]:\//.test(decodedPath)) {
      return decodedPath.slice(1).replace(/\//g, '\\');
    }

    return decodedPath.replace(/\//g, '\\');
  } catch {
    return '';
  }
}

function normalizeCliMarkdownPath(inputArg) {
  if (typeof inputArg !== 'string') {
    return '';
  }

  const trimmed = inputArg.trim();
  if (!trimmed || trimmed.startsWith('-')) {
    return '';
  }

  const dequoted = trimmed.replace(/^"+|"+$/g, '');
  const rawPath = /^file:\/\//i.test(dequoted)
    ? normalizeFileUrlToWindowsPath(dequoted)
    : dequoted;
  if (!rawPath) {
    return '';
  }

  const resolvedPath = path.resolve(rawPath);
  if (!isMarkdownFilePath(resolvedPath)) {
    return '';
  }

  try {
    const stat = fsSync.statSync(resolvedPath);
    if (!stat.isFile()) {
      return '';
    }
  } catch {
    return '';
  }

  return resolvedPath;
}

function resolveLaunchMarkdownPath(argv) {
  if (!Array.isArray(argv) || argv.length === 0) {
    return '';
  }

  for (const arg of argv.slice(1)) {
    const normalizedPath = normalizeCliMarkdownPath(arg);
    if (normalizedPath) {
      return normalizedPath;
    }
  }

  return '';
}

const initialLaunchFilePath = resolveLaunchMarkdownPath(process.argv);
if (initialLaunchFilePath) {
  pendingLaunchFilePaths.push(initialLaunchFilePath);
}

const hasSingleInstanceLock = app.requestSingleInstanceLock();
if (!hasSingleInstanceLock) {
  app.quit();
}

function createTraceId(prefix = 'srv') {
  return `${prefix}-${Date.now()}-${crypto.randomBytes(4).toString('hex')}`;
}

function normalizeTraceId(value) {
  if (typeof value !== 'string') {
    return createTraceId();
  }

  const trimmed = value.trim();
  const isValid = /^[A-Za-z0-9:._-]{4,64}$/.test(trimmed);

  return isValid ? trimmed : createTraceId();
}

function safeErrorMessage(error) {
  if (error instanceof Error) {
    return `${error.name}: ${error.message}`;
  }

  return String(error);
}

async function ensureRuntimePaths() {
  const userDataPath = app.getPath('userData');
  const logsDir = path.join(userDataPath, 'logs');

  await fs.mkdir(logsDir, { recursive: true });

  runtimeLogFilePath = path.join(logsDir, 'runtime.log');
  settingsFilePath = path.join(userDataPath, 'settings.json');
}

async function appendRuntimeLog(level, message, traceId, detail) {
  const line = JSON.stringify({
    time: new Date().toISOString(),
    level,
    traceId,
    message,
    detail: detail ?? null,
  });

  if (!runtimeLogFilePath) {
    return;
  }

  await fs.appendFile(runtimeLogFilePath, `${line}\n`, 'utf8');
}

function logEvent(level, message, traceId, detail) {
  const logLine = `[${level}] [${traceId}] ${message}`;

  if (level === 'error') {
    console.error(logLine, detail ?? '');
  } else {
    console.log(logLine, detail ?? '');
  }

  void appendRuntimeLog(level, message, traceId, detail).catch((error) => {
    console.error('[logger] append failed', safeErrorMessage(error));
  });
}

function toUtf8Bytes(input) {
  return Buffer.byteLength(input, 'utf8');
}

function requireString(input, fieldName, maxLen = 4096) {
  if (typeof input !== 'string') {
    throw new Error(`${fieldName} must be a string`);
  }

  if (input.length === 0) {
    throw new Error(`${fieldName} cannot be empty`);
  }

  if (input.length > maxLen) {
    throw new Error(`${fieldName} is too long`);
  }

  return input;
}

function validateSavePayload(payload) {
  const content = typeof payload?.content === 'string' ? payload.content : null;
  if (content === null) {
    throw new Error('content must be string');
  }

  if (toUtf8Bytes(content) > MAX_CONTENT_BYTES) {
    throw new Error('content is too large');
  }

  let filePath = null;
  if (payload?.filePath !== undefined && payload?.filePath !== null) {
    filePath = requireString(payload.filePath, 'filePath');
  }

  return {
    filePath,
    content,
  };
}

function validateExportHtmlPayload(payload) {
  const html = typeof payload?.html === 'string' ? payload.html : null;
  if (html === null) {
    throw new Error('html must be string');
  }

  if (toUtf8Bytes(html) > MAX_CONTENT_BYTES) {
    throw new Error('html is too large');
  }

  let sourceFilePath = null;
  if (payload?.sourceFilePath !== undefined && payload?.sourceFilePath !== null) {
    sourceFilePath = requireString(payload.sourceFilePath, 'sourceFilePath');
  }

  let targetPath = null;
  if (payload?.targetPath !== undefined && payload?.targetPath !== null) {
    targetPath = requireString(payload.targetPath, 'targetPath');
  }

  return {
    html,
    sourceFilePath,
    targetPath,
  };
}

function validateExportDocumentPayload(payload) {
  const format = requireString(payload?.format, 'format', 32).toLowerCase();
  if (!SUPPORTED_EXPORT_FORMATS.has(format)) {
    throw new Error(`unsupported export format: ${format}`);
  }

  const html = typeof payload?.html === 'string' ? payload.html : null;
  if (html === null) {
    throw new Error('html must be string');
  }
  if (toUtf8Bytes(html) > MAX_CONTENT_BYTES) {
    throw new Error('html is too large');
  }

  const markdown = typeof payload?.markdown === 'string' ? payload.markdown : null;
  if (markdown === null) {
    throw new Error('markdown must be string');
  }
  if (toUtf8Bytes(markdown) > MAX_CONTENT_BYTES) {
    throw new Error('markdown is too large');
  }

  let sourceFilePath = null;
  if (payload?.sourceFilePath !== undefined && payload?.sourceFilePath !== null) {
    sourceFilePath = requireString(payload.sourceFilePath, 'sourceFilePath');
  }

  let targetPath = null;
  if (payload?.targetPath !== undefined && payload?.targetPath !== null) {
    targetPath = requireString(payload.targetPath, 'targetPath');
  }

  return {
    format,
    html,
    markdown,
    sourceFilePath,
    targetPath,
  };
}

function runCommand(command, args, timeoutMs = 5000) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    let settled = false;

    const timer = setTimeout(() => {
      if (settled) {
        return;
      }
      settled = true;
      child.kill('SIGKILL');
      reject(new Error(`${command} timeout after ${timeoutMs}ms`));
    }, timeoutMs);

    child.stdout.on('data', (chunk) => {
      stdout += String(chunk || '');
    });
    child.stderr.on('data', (chunk) => {
      stderr += String(chunk || '');
    });
    child.once('error', (error) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      reject(error);
    });
    child.once('close', (code) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      if (code === 0) {
        resolve({
          stdout,
          stderr,
        });
        return;
      }

      const stderrText = stderr.trim();
      const message = stderrText || `${command} exit code ${code}`;
      reject(new Error(message));
    });
  });
}

async function addRegistryValue(keyPath, valueName, valueType, valueData) {
  const args = ['add', keyPath];

  if (valueName === null) {
    args.push('/ve');
  } else {
    args.push('/v', valueName);
  }

  if (valueType) {
    args.push('/t', valueType);
  }

  if (valueData !== undefined && valueData !== null) {
    args.push('/d', valueData);
  }

  args.push('/f');
  await runCommand('reg.exe', args, 15000);
}

function getAssociationExecutablePath() {
  const portableExecutableFile = process.env.PORTABLE_EXECUTABLE_FILE || '';
  if (portableExecutableFile && fsSync.existsSync(portableExecutableFile)) {
    return portableExecutableFile;
  }

  if (app.isPackaged) {
    return process.execPath;
  }

  return path.resolve(__dirname, '..', 'release', 'electron', 'UniReadMD.exe');
}

function getAssociationIconPath() {
  const portableExecutableDir = process.env.PORTABLE_EXECUTABLE_DIR || '';
  const packagedIconPath = portableExecutableDir
    ? path.join(portableExecutableDir, DEFAULT_APP_ICON_FILE_NAME)
    : app.isPackaged
      ? path.join(path.dirname(process.execPath), DEFAULT_APP_ICON_FILE_NAME)
      : path.resolve(__dirname, '..', 'release', 'electron', DEFAULT_APP_ICON_FILE_NAME);

  if (fsSync.existsSync(packagedIconPath)) {
    return packagedIconPath;
  }

  return resolveAppIconPath();
}

async function refreshWindowsShellIconCache() {
  if (process.platform !== 'win32') {
    return;
  }

  const commands = [
    ['ie4uinit.exe', ['-show']],
    ['ie4uinit.exe', ['-ClearIconCache']],
  ];

  for (const [command, args] of commands) {
    try {
      await runCommand(command, args, 15000);
    } catch {
      // ignore shell refresh failures, the association itself is already written
    }
  }
}

async function registerMarkdownFileAssociation() {
  if (process.platform !== 'win32') {
    return {
      appName: DEFAULT_APP_NAME,
      settingsUri: '',
      requiresUserConfirmation: false,
    };
  }

  const executablePath = getAssociationExecutablePath();
  if (!fsSync.existsSync(executablePath)) {
    throw new Error(`executable not found: ${executablePath}`);
  }

  const iconPath = getAssociationIconPath();
  const command = `"${executablePath}" "%1"`;
  const appKey = 'HKCU\\Software\\Classes\\Applications\\UniReadMD.exe';
  const progIdKey = `HKCU\\Software\\Classes\\${DEFAULT_APP_PROG_ID}`;
  const capabilitiesKey = `HKCU\\${DEFAULT_APP_CAPABILITIES_PATH}`;

  await addRegistryValue(
    appKey,
    'FriendlyAppName',
    'REG_SZ',
    DEFAULT_APP_NAME,
  );
  await addRegistryValue(
    appKey,
    'DefaultIcon',
    'REG_SZ',
    iconPath,
  );
  await addRegistryValue(
    `${appKey}\\shell\\open\\command`,
    null,
    'REG_SZ',
    command,
  );
  await addRegistryValue(
    `${appKey}\\SupportedTypes`,
    '.md',
    'REG_SZ',
    '',
  );
  await addRegistryValue(
    `${appKey}\\SupportedTypes`,
    '.markdown',
    'REG_SZ',
    '',
  );

  await addRegistryValue(
    progIdKey,
    null,
    'REG_SZ',
    DEFAULT_APP_DESCRIPTION,
  );
  await addRegistryValue(
    progIdKey,
    'FriendlyTypeName',
    'REG_SZ',
    DEFAULT_APP_DESCRIPTION,
  );
  await addRegistryValue(
    `${progIdKey}\\DefaultIcon`,
    null,
    'REG_SZ',
    iconPath,
  );
  await addRegistryValue(
    `${progIdKey}\\shell\\open\\command`,
    null,
    'REG_SZ',
    command,
  );

  await addRegistryValue(
    capabilitiesKey,
    'ApplicationName',
    'REG_SZ',
    DEFAULT_APP_NAME,
  );
  await addRegistryValue(
    capabilitiesKey,
    'ApplicationIcon',
    'REG_SZ',
    iconPath,
  );
  await addRegistryValue(
    capabilitiesKey,
    'ApplicationDescription',
    'REG_SZ',
    DEFAULT_APP_DESCRIPTION,
  );
  await addRegistryValue(
    `${capabilitiesKey}\\FileAssociations`,
    '.md',
    'REG_SZ',
    DEFAULT_APP_PROG_ID,
  );
  await addRegistryValue(
    `${capabilitiesKey}\\FileAssociations`,
    '.markdown',
    'REG_SZ',
    DEFAULT_APP_PROG_ID,
  );
  await addRegistryValue(
    'HKCU\\Software\\RegisteredApplications',
    DEFAULT_APP_NAME,
    'REG_SZ',
    DEFAULT_APP_CAPABILITIES_PATH,
  );
  await addRegistryValue(
    'HKCU\\Software\\Classes\\.md\\OpenWithProgids',
    DEFAULT_APP_PROG_ID,
    'REG_SZ',
    '',
  );
  await addRegistryValue(
    'HKCU\\Software\\Classes\\.markdown\\OpenWithProgids',
    DEFAULT_APP_PROG_ID,
    'REG_SZ',
    '',
  );

  await refreshWindowsShellIconCache();

  const settingsUri = `ms-settings:defaultapps?registeredAppUser=${encodeURIComponent(
    DEFAULT_APP_NAME,
  )}`;

  return {
    appName: DEFAULT_APP_NAME,
    settingsUri,
    requiresUserConfirmation: true,
  };
}

async function probePandoc() {
  try {
    const versionResult = await runCommand('pandoc', ['--version']);
    const versionLine = String(versionResult.stdout || '')
      .split(/\r?\n/)
      .map((item) => item.trim())
      .find((item) => item.length > 0) || '';

    let formatText = '';
    try {
      const formatResult = await runCommand('pandoc', ['--list-output-formats']);
      formatText = String(formatResult.stdout || '');
    } catch {
      formatText = '';
    }

    const outputFormats = formatText
      .split(/\s+/)
      .map((item) => item.trim().toLowerCase())
      .filter((item) => item.length > 0);
    return {
      available: true,
      version: versionLine,
      outputFormats,
    };
  } catch {
    return {
      available: false,
      version: '',
      outputFormats: [],
    };
  }
}

async function getPandocProbe() {
  if (!pandocProbeCache) {
    pandocProbeCache = probePandoc();
  }
  return pandocProbeCache;
}

function mergeSettings(raw) {
  const merged = { ...DEFAULT_SETTINGS };

  if (!raw || typeof raw !== 'object') {
    return merged;
  }

  const input = raw;

  if (input.theme === 'light' || input.theme === 'dark' || input.theme === 'system') {
    merged.theme = input.theme;
  }

  if (typeof input.fontSize === 'number' && input.fontSize >= 10 && input.fontSize <= 28) {
    merged.fontSize = Math.round(input.fontSize);
  }

  if (typeof input.contentZoom === 'number' && input.contentZoom >= 50 && input.contentZoom <= 200) {
    merged.contentZoom = Math.round(input.contentZoom);
  }

  if (input.viewMode === 'edit' || input.viewMode === 'preview' || input.viewMode === 'split') {
    merged.viewMode = input.viewMode;
  }

  if (typeof input.autoSave === 'boolean') {
    merged.autoSave = input.autoSave;
  }

  if (typeof input.userCssEnabled === 'boolean') {
    merged.userCssEnabled = input.userCssEnabled;
  }

  if (typeof input.userCssPath === 'string' && input.userCssPath.length <= 1024) {
    merged.userCssPath = input.userCssPath;
  }

  if (typeof input.spellCheckEnabled === 'boolean') {
    merged.spellCheckEnabled = input.spellCheckEnabled;
  }

  if (typeof input.spellCheckLanguage === 'string' && input.spellCheckLanguage.length <= 64) {
    merged.spellCheckLanguage = input.spellCheckLanguage;
  }

  return merged;
}

function normalizeRecentFiles(raw) {
  if (!Array.isArray(raw)) {
    return [];
  }

  const seen = new Set();
  const output = [];

  raw.forEach((item) => {
    if (typeof item !== 'string') {
      return;
    }

    const normalizedPath = item.trim();
    if (!normalizedPath || !isMarkdownFilePath(normalizedPath)) {
      return;
    }

    const dedupeKey = normalizedPath.toLowerCase();
    if (seen.has(dedupeKey)) {
      return;
    }

    seen.add(dedupeKey);
    output.push(normalizedPath);
  });

  return output.slice(0, 12);
}

function toPersistedSettingsPayload(settings, recentFiles) {
  return {
    ...settings,
    recentFiles: normalizeRecentFiles(recentFiles),
  };
}

async function readSettingsStorage() {
  if (settingsCache && recentFilesCache) {
    return {
      settings: settingsCache,
      recentFiles: recentFilesCache,
    };
  }

  try {
    const text = await fs.readFile(settingsFilePath, 'utf8');
    const parsed = JSON.parse(text);
    settingsCache = mergeSettings(parsed);
    recentFilesCache = normalizeRecentFiles(parsed?.recentFiles);
  } catch {
    settingsCache = { ...DEFAULT_SETTINGS };
    recentFilesCache = [];
  }

  return {
    settings: settingsCache,
    recentFiles: recentFilesCache,
  };
}

async function readSettings() {
  const storage = await readSettingsStorage();
  return storage.settings;
}

async function readRecentFiles() {
  const storage = await readSettingsStorage();
  return storage.recentFiles;
}

async function persistSettings(nextSettings) {
  settingsCache = nextSettings;
  const recentFiles = await readRecentFiles();
  const payload = toPersistedSettingsPayload(nextSettings, recentFiles);
  await fs.writeFile(settingsFilePath, JSON.stringify(payload, null, 2), 'utf8');
}

async function persistRecentFiles(nextRecentFiles) {
  recentFilesCache = normalizeRecentFiles(nextRecentFiles);
  const settings = await readSettings();
  const payload = toPersistedSettingsPayload(settings, recentFilesCache);
  await fs.writeFile(settingsFilePath, JSON.stringify(payload, null, 2), 'utf8');
}

function getRendererUrl() {
  if (isDev) {
    return process.env.ELECTRON_START_URL;
  }

  if (rendererStaticServerUrl) {
    return rendererStaticServerUrl;
  }

  const indexPath = path.resolve(__dirname, '..', 'frontend', 'dist', 'index.html');
  return pathToFileURL(indexPath).toString();
}

function isExternalNavigationUrl(rawUrl) {
  if (typeof rawUrl !== 'string' || rawUrl.trim().length === 0) {
    return false;
  }

  try {
    const parsed = new URL(rawUrl);
    if (parsed.protocol === 'file:') {
      return false;
    }

    if (parsed.protocol === 'http:' || parsed.protocol === 'https:') {
      return true;
    }

    return parsed.protocol === 'mailto:' || parsed.protocol === 'tel:';
  } catch {
    return false;
  }
}

function getContentTypeByExtension(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const contentTypeMap = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.mjs': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.svg': 'image/svg+xml',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.ico': 'image/x-icon',
    '.map': 'application/json; charset=utf-8',
    '.woff': 'font/woff',
    '.woff2': 'font/woff2',
    '.ttf': 'font/ttf',
  };

  return contentTypeMap[ext] || 'application/octet-stream';
}

function resolveRendererFilePath(distRoot, requestUrl) {
  const requestPath = decodeURIComponent((requestUrl || '/').split('?')[0] || '/');
  const normalizedPath = requestPath === '/' ? '/index.html' : requestPath;
  const sanitizedRelativePath = normalizedPath.replace(/^\/+/, '');
  const absolutePath = path.resolve(distRoot, sanitizedRelativePath);

  if (!absolutePath.startsWith(distRoot)) {
    return '';
  }

  return absolutePath;
}

function patchRendererIndexHtml(htmlText) {
  const scriptMatch = htmlText.match(
    /<script\s+type="module"\s+crossorigin\s+src="([^"]+)"><\/script>/i,
  );
  if (!scriptMatch || !scriptMatch[1]) {
    return htmlText;
  }

  const entrySrc = scriptMatch[1];
  const bootScriptSrc = `/__unireadmd_bootstrap__.js?entry=${encodeURIComponent(entrySrc)}`;
  const bootScript = `<script src="${bootScriptSrc}"></script>`;

  return htmlText
    .replace(/<link\s+rel="modulepreload"[^>]*>\s*/gi, '')
    .replace(scriptMatch[0], bootScript);
}

async function startRendererStaticServer() {
  if (isDev || rendererStaticServer) {
    return;
  }

  const distRoot = path.resolve(__dirname, '..', 'frontend', 'dist');

  rendererStaticServer = http.createServer((request, response) => {
    const requestUrl = request.url || '/';
    const requestPath = decodeURIComponent((requestUrl || '/').split('?')[0] || '/');

    if (requestPath === '/__unireadmd_bootstrap__.js') {
      const entry = new URL(`http://127.0.0.1${requestUrl}`).searchParams.get('entry') || '';
      const bootstrapJs = [
        '(async () => {',
        `  const entry = ${JSON.stringify(entry)};`,
        '  try {',
        '    await import(entry);',
        '  } catch (error) {',
        "    console.error('[boot-import-failed]', error);",
        '  }',
        '})();',
      ].join('\n');
      response.writeHead(200, {
        'Content-Type': 'text/javascript; charset=utf-8',
        'Cache-Control': 'no-cache',
      });
      response.end(bootstrapJs);
      return;
    }

    const candidatePath = resolveRendererFilePath(distRoot, requestUrl);
    const fallbackPath = path.join(distRoot, 'index.html');
    const targetPath = candidatePath && fsSync.existsSync(candidatePath)
      ? candidatePath
      : fallbackPath;

    try {
      const stat = fsSync.statSync(targetPath);
      if (!stat.isFile()) {
        response.writeHead(404);
        response.end('Not Found');
        return;
      }
    } catch {
      response.writeHead(404);
      response.end('Not Found');
      return;
    }

    const isIndexHtml = targetPath.toLowerCase().endsWith(`${path.sep}index.html`);
    if (isIndexHtml) {
      try {
        const htmlText = fsSync.readFileSync(targetPath, 'utf8');
        const patchedHtml = patchRendererIndexHtml(htmlText);
        response.writeHead(200, {
          'Content-Type': 'text/html; charset=utf-8',
          'Cache-Control': 'no-cache',
        });
        response.end(patchedHtml);
      } catch {
        response.writeHead(500);
        response.end('Internal Server Error');
      }
      return;
    }

    response.writeHead(200, {
      'Content-Type': getContentTypeByExtension(targetPath),
      'Cache-Control': 'no-cache',
    });

    fsSync.createReadStream(targetPath).pipe(response);
  });

  await new Promise((resolve, reject) => {
    rendererStaticServer.once('error', reject);
    rendererStaticServer.listen(0, '127.0.0.1', () => resolve());
  });

  const address = rendererStaticServer.address();
  const port = typeof address === 'object' && address ? address.port : 0;
  rendererStaticServerUrl = `http://127.0.0.1:${port}/`;

  logEvent('info', 'renderer static server started', createTraceId('boot'), {
    url: rendererStaticServerUrl,
    distRoot,
  });
}

function getEventWindow(event) {
  return BrowserWindow.fromWebContents(event.sender) || BrowserWindow.getFocusedWindow();
}

function getAvailableSpellLanguages(sessionRef) {
  if (!sessionRef || !Array.isArray(sessionRef.availableSpellCheckerLanguages)) {
    return [];
  }

  return sessionRef.availableSpellCheckerLanguages.filter((item) => typeof item === 'string');
}

function toSpellCheckState(settings, sessionRef) {
  const availableLanguages = getAvailableSpellLanguages(sessionRef);
  const preferred = typeof settings.spellCheckLanguage === 'string'
    ? settings.spellCheckLanguage
    : '';
  const language = preferred || availableLanguages[0] || 'en-US';
  const dictionaryStatus = availableLanguages.includes(language) ? 'ready' : 'missing';

  return {
    enabled: Boolean(settings.spellCheckEnabled),
    language,
    availableLanguages,
    dictionaryStatus,
  };
}

function applySpellCheckState(sessionRef, state, traceId) {
  if (!sessionRef) {
    return;
  }

  try {
    sessionRef.setSpellCheckerEnabled(Boolean(state.enabled));

    if (state.dictionaryStatus === 'ready' && state.language) {
      sessionRef.setSpellCheckerLanguages([state.language]);
    }
  } catch (error) {
    logEvent('error', 'apply spell check state failed', traceId, {
      error: safeErrorMessage(error),
    });
  }
}

async function spellGetState(event, _payload, traceId) {
  const settings = await readSettings();
  const win = getEventWindow(event);
  const state = toSpellCheckState(settings, win?.webContents.session);
  applySpellCheckState(win?.webContents.session, state, traceId);
  logEvent('info', 'spell getState', traceId, state);
  return state;
}

async function spellSetEnabled(event, payload, traceId) {
  if (typeof payload?.enabled !== 'boolean') {
    throw new Error('invalid spell enabled flag');
  }

  const settings = await readSettings();
  const nextSettings = {
    ...settings,
    spellCheckEnabled: payload.enabled,
  };

  await persistSettings(nextSettings);
  const win = getEventWindow(event);
  const state = toSpellCheckState(nextSettings, win?.webContents.session);
  applySpellCheckState(win?.webContents.session, state, traceId);
  logEvent('info', 'spell setEnabled', traceId, state);
  return state;
}

async function spellSetLanguage(event, payload, traceId) {
  const language = requireString(payload?.language, 'language', 64);
  const settings = await readSettings();
  const nextSettings = {
    ...settings,
    spellCheckLanguage: language,
  };

  await persistSettings(nextSettings);
  const win = getEventWindow(event);
  const state = toSpellCheckState(nextSettings, win?.webContents.session);
  applySpellCheckState(win?.webContents.session, state, traceId);
  logEvent('info', 'spell setLanguage', traceId, state);
  return state;
}

async function openMarkdownFile(_event, payload, traceId) {
  const win = getEventWindow(_event);
  if (win && !win.isDestroyed()) {
    if (win.isMinimized()) {
      win.restore();
    }
    win.focus();
  }

  logEvent('info', 'open markdown dialog requested', traceId, {
    hasWindow: Boolean(win && !win.isDestroyed()),
  });

  const dialogOptions = {
    properties: ['openFile'],
    filters: [
      {
        name: 'Markdown',
        extensions: ['md', 'markdown', 'txt'],
      },
    ],
  };
  const result = win && !win.isDestroyed()
    ? await dialog.showOpenDialog(win, dialogOptions)
    : await dialog.showOpenDialog(dialogOptions);

  if (result.canceled || result.filePaths.length === 0) {
    logEvent('info', 'open markdown canceled', traceId);
    return null;
  }

  const selectedPath = result.filePaths[0];
  const content = await fs.readFile(selectedPath, 'utf8');

  logEvent('info', 'open markdown success', traceId, { filePath: selectedPath });

  return {
    filePath: selectedPath,
    content,
  };
}

async function openMarkdownFromPath(_event, payload, traceId) {
  const filePath = requireString(payload?.filePath, 'filePath');
  const content = await fs.readFile(filePath, 'utf8');

  logEvent('info', 'open markdown from path success', traceId, { filePath });

  return {
    filePath,
    content,
  };
}

function isMarkdownFileName(fileName) {
  const normalized = String(fileName || '').toLowerCase();
  return normalized.endsWith('.md') || normalized.endsWith('.markdown');
}

async function listMarkdownInDirectory(_event, payload, traceId) {
  const filePath = requireString(payload?.filePath, 'filePath');
  const directoryPath = path.dirname(filePath);
  const entries = await fs.readdir(directoryPath, { withFileTypes: true });

  const filePaths = entries
    .filter((entry) => entry.isFile() && isMarkdownFileName(entry.name))
    .map((entry) => path.join(directoryPath, entry.name))
    .sort((left, right) => {
      return path.basename(left).localeCompare(path.basename(right), 'zh-CN');
    });

  logEvent('info', 'list markdown in directory success', traceId, {
    directoryPath,
    count: filePaths.length,
  });

  return {
    directoryPath,
    filePaths,
  };
}

async function saveMarkdownFile(_event, payload, traceId) {
  const validated = validateSavePayload(payload);
  let targetPath = validated.filePath;

  if (!targetPath) {
    const win = getEventWindow(_event);
    if (win && !win.isDestroyed()) {
      if (win.isMinimized()) {
        win.restore();
      }
      win.focus();
    }

    const dialogOptions = {
      filters: [
        {
          name: 'Markdown',
          extensions: ['md'],
        },
      ],
    };
    const saveResult = win && !win.isDestroyed()
      ? await dialog.showSaveDialog(win, dialogOptions)
      : await dialog.showSaveDialog(dialogOptions);

    if (saveResult.canceled || !saveResult.filePath) {
      logEvent('info', 'save markdown canceled', traceId);
      return null;
    }

    targetPath = saveResult.filePath;
  }

  await fs.writeFile(targetPath, validated.content, 'utf8');

  logEvent('info', 'save markdown success', traceId, { filePath: targetPath });

  return {
    filePath: targetPath,
  };
}

function deriveExportPath(sourceFilePath, format) {
  const extensionMap = {
    html: 'html',
    pdf: 'pdf',
    png: 'png',
    jpeg: 'jpg',
    docx: 'docx',
    odt: 'odt',
    rtf: 'rtf',
  };
  const extension = extensionMap[format] || 'html';
  const fallbackName = `document.${extension}`;

  if (!sourceFilePath) {
    return fallbackName;
  }

  const baseName = path.basename(sourceFilePath, path.extname(sourceFilePath));
  return `${baseName || 'document'}.${extension}`;
}

function getExportDialogFilter(format) {
  if (format === 'pdf') {
    return {
      name: 'PDF',
      extensions: ['pdf'],
    };
  }
  if (format === 'png') {
    return {
      name: 'PNG Image',
      extensions: ['png'],
    };
  }
  if (format === 'jpeg') {
    return {
      name: 'JPEG Image',
      extensions: ['jpg', 'jpeg'],
    };
  }
  if (format === 'docx') {
    return {
      name: 'Word',
      extensions: ['docx'],
    };
  }
  if (format === 'odt') {
    return {
      name: 'OpenDocument',
      extensions: ['odt'],
    };
  }
  if (format === 'rtf') {
    return {
      name: 'RTF',
      extensions: ['rtf'],
    };
  }

  return {
    name: 'HTML',
    extensions: ['html'],
  };
}

async function resolveExportTargetPath(event, sourceFilePath, format, targetPath, traceId) {
  if (targetPath) {
    return targetPath;
  }

  const win = getEventWindow(event);
  if (win && !win.isDestroyed()) {
    if (win.isMinimized()) {
      win.restore();
    }
    win.focus();
  }

  const dialogOptions = {
    defaultPath: deriveExportPath(sourceFilePath, format),
    filters: [getExportDialogFilter(format)],
  };
  const saveResult = win && !win.isDestroyed()
    ? await dialog.showSaveDialog(win, dialogOptions)
    : await dialog.showSaveDialog(dialogOptions);

  if (saveResult.canceled || !saveResult.filePath) {
    logEvent('info', 'export canceled', traceId, { format });
    return '';
  }

  return saveResult.filePath;
}

async function exportViaChromium(html, format, targetPath) {
  const exportWindow = new BrowserWindow({
    show: false,
    width: 1280,
    height: 960,
    webPreferences: {
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
    },
  });

  try {
    const encodedHtml = `data:text/html;charset=utf-8,${encodeURIComponent(html)}`;
    await exportWindow.loadURL(encodedHtml);
    await exportWindow.webContents.executeJavaScript(
      `
        Promise.resolve(document.fonts?.ready)
          .catch(() => null)
          .then(() => true);
      `,
      true,
    );

    if (format === 'pdf') {
      const pdfData = await exportWindow.webContents.printToPDF({
        printBackground: true,
      });
      await fs.writeFile(targetPath, pdfData);
      return;
    }

    const image = await exportWindow.webContents.capturePage();
    const imageBuffer = format === 'jpeg'
      ? image.toJPEG(92)
      : image.toPNG();
    await fs.writeFile(targetPath, imageBuffer);
  } finally {
    if (!exportWindow.isDestroyed()) {
      exportWindow.destroy();
    }
  }
}

async function exportViaPandoc(markdown, format, targetPath, traceId) {
  const pandoc = await getPandocProbe();
  if (!pandoc.available) {
    throw new Error(
      '未检测到 Pandoc。请先安装 Pandoc（https://pandoc.org），或改用 HTML/PDF/Image 导出。',
    );
  }

  if (!pandoc.outputFormats.includes(format)) {
    throw new Error(
      `Pandoc 当前不支持 ${format} 导出。请检查 pandoc --list-output-formats 输出。`,
    );
  }

  const tempDir = await fs.mkdtemp(path.join(app.getPath('temp'), 'unireadmd-export-'));
  const sourceMarkdownPath = path.join(tempDir, 'source.md');

  try {
    await fs.writeFile(sourceMarkdownPath, markdown, 'utf8');
    await runCommand(
      'pandoc',
      [
        sourceMarkdownPath,
        '--from',
        'markdown',
        '--to',
        format,
        '--output',
        targetPath,
      ],
      30_000,
    );
    logEvent('info', 'pandoc export success', traceId, { format, targetPath });
  } catch (error) {
    throw new Error(`Pandoc 导出失败：${safeErrorMessage(error)}`);
  } finally {
    await fs.rm(tempDir, {
      recursive: true,
      force: true,
    });
  }
}

async function getExportCapabilities(_event, _payload, traceId) {
  const pandoc = await getPandocProbe();
  const pandocFormats = pandoc.available
    ? PANDOC_EXPORT_FORMATS.filter((format) => pandoc.outputFormats.includes(format))
    : [];

  const result = {
    builtinFormats: [...BUILTIN_EXPORT_FORMATS],
    pandocAvailable: pandoc.available,
    pandocVersion: pandoc.version,
    pandocFormats,
  };
  logEvent('info', 'export capabilities', traceId, result);
  return result;
}

async function exportDocumentFile(event, payload, traceId) {
  const validated = validateExportDocumentPayload(payload);
  const targetPath = await resolveExportTargetPath(
    event,
    validated.sourceFilePath,
    validated.format,
    validated.targetPath,
    traceId,
  );
  if (!targetPath) {
    return null;
  }

  logEvent('info', 'export document start', traceId, {
    format: validated.format,
    targetPath,
  });

  let engine = 'builtin';
  if (validated.format === 'html') {
    await fs.writeFile(targetPath, validated.html, 'utf8');
  } else if (validated.format === 'pdf' || validated.format === 'png' || validated.format === 'jpeg') {
    await exportViaChromium(validated.html, validated.format, targetPath);
  } else {
    engine = 'pandoc';
    await exportViaPandoc(validated.markdown, validated.format, targetPath, traceId);
  }

  logEvent('info', 'export document success', traceId, {
    format: validated.format,
    targetPath,
    engine,
  });

  return {
    filePath: targetPath,
    format: validated.format,
    engine,
  };
}

async function exportHtmlFile(event, payload, traceId) {
  const validated = validateExportHtmlPayload(payload);
  return exportDocumentFile(
    event,
    {
      format: 'html',
      html: validated.html,
      markdown: '',
      sourceFilePath: validated.sourceFilePath || undefined,
      targetPath: validated.targetPath || undefined,
      traceId,
    },
    traceId,
  );
}

function windowAction(event, payload, traceId, action) {
  const win = getEventWindow(event);
  if (!win) {
    throw new Error('window not found');
  }

  action(win);
  logEvent('info', 'window action executed', traceId, payload ?? null);
}

async function settingsGetAll(_event, _payload, traceId) {
  const settings = await readSettings();
  logEvent('info', 'settings getAll', traceId);

  return {
    settings,
  };
}

async function settingsGetRecentFiles(_event, _payload, traceId) {
  const filePaths = await readRecentFiles();
  logEvent('info', 'settings getRecentFiles', traceId, {
    count: filePaths.length,
  });

  return {
    filePaths,
  };
}

async function settingsGet(_event, payload, traceId) {
  const key = payload?.key;
  if (!ALLOWED_SETTINGS.has(key)) {
    throw new Error('invalid setting key');
  }

  const settings = await readSettings();

  return {
    key,
    value: settings[key],
  };
}

async function settingsSet(_event, payload, traceId) {
  const key = payload?.key;
  if (!ALLOWED_SETTINGS.has(key)) {
    throw new Error('invalid setting key');
  }

  const settings = await readSettings();
  const nextSettings = { ...settings };

  if (key === 'theme') {
    if (payload.value !== 'light' && payload.value !== 'dark' && payload.value !== 'system') {
      throw new Error('invalid theme');
    }
    nextSettings.theme = payload.value;
  }

  if (key === 'fontSize') {
    if (typeof payload.value !== 'number' || payload.value < 10 || payload.value > 28) {
      throw new Error('invalid fontSize');
    }
    nextSettings.fontSize = Math.round(payload.value);
  }

  if (key === 'contentZoom') {
    if (typeof payload.value !== 'number' || payload.value < 50 || payload.value > 200) {
      throw new Error('invalid contentZoom');
    }
    nextSettings.contentZoom = Math.round(payload.value);
  }

  if (key === 'viewMode') {
    if (payload.value !== 'edit' && payload.value !== 'preview' && payload.value !== 'split') {
      throw new Error('invalid viewMode');
    }
    nextSettings.viewMode = payload.value;
  }

  if (key === 'autoSave') {
    if (typeof payload.value !== 'boolean') {
      throw new Error('invalid autoSave');
    }
    nextSettings.autoSave = payload.value;
  }

  if (key === 'userCssEnabled') {
    if (typeof payload.value !== 'boolean') {
      throw new Error('invalid userCssEnabled');
    }
    nextSettings.userCssEnabled = payload.value;
  }

  if (key === 'userCssPath') {
    if (typeof payload.value !== 'string' || payload.value.length > 1024) {
      throw new Error('invalid userCssPath');
    }
    nextSettings.userCssPath = payload.value;
  }

  if (key === 'spellCheckEnabled') {
    if (typeof payload.value !== 'boolean') {
      throw new Error('invalid spellCheckEnabled');
    }
    nextSettings.spellCheckEnabled = payload.value;
  }

  if (key === 'spellCheckLanguage') {
    if (typeof payload.value !== 'string' || payload.value.length > 64) {
      throw new Error('invalid spellCheckLanguage');
    }
    nextSettings.spellCheckLanguage = payload.value;
  }

  await persistSettings(nextSettings);
  logEvent('info', 'settings set', traceId, { key, value: nextSettings[key] });

  return {
    key,
    value: nextSettings[key],
    settings: nextSettings,
  };
}

async function settingsSetRecentFiles(_event, payload, traceId) {
  const filePaths = normalizeRecentFiles(payload?.filePaths);
  await persistRecentFiles(filePaths);
  logEvent('info', 'settings setRecentFiles', traceId, {
    count: filePaths.length,
  });

  return {
    filePaths,
  };
}

async function settingsSelectUserCssFile(_event, _payload, traceId) {
  const win = getEventWindow(_event);
  if (win && !win.isDestroyed()) {
    if (win.isMinimized()) {
      win.restore();
    }
    win.focus();
  }

  const dialogOptions = {
    properties: ['openFile'],
    filters: [
      {
        name: 'CSS',
        extensions: ['css'],
      },
    ],
  };
  const result = win && !win.isDestroyed()
    ? await dialog.showOpenDialog(win, dialogOptions)
    : await dialog.showOpenDialog(dialogOptions);

  if (result.canceled || result.filePaths.length === 0) {
    logEvent('info', 'select user css canceled', traceId);
    return null;
  }

  return {
    filePath: result.filePaths[0],
  };
}

async function showFileTreeContextMenu(event, payload, traceId) {
  const filePath = requireString(payload?.filePath, 'filePath');
  const win = getEventWindow(event);

  if (!win || win.isDestroyed()) {
    throw new Error('window is not available');
  }

  const fileName = path.basename(filePath);
  const contextMenu = Menu.buildFromTemplate([
    {
      label: '复制文件名',
      click: () => {
        clipboard.writeText(fileName);
        logEvent('info', 'context menu copied file name', traceId, { fileName });
      },
    },
    {
      label: '复制文件全路径',
      click: () => {
        clipboard.writeText(filePath);
        logEvent('info', 'context menu copied file path', traceId, { filePath });
      },
    },
    {
      type: 'separator',
    },
    {
      label: '在资源管理器中显示',
      click: () => {
        shell.showItemInFolder(path.normalize(filePath));
        logEvent('info', 'context menu reveal in explorer', traceId, { filePath });
      },
    },
  ]);
  const hasX = Number.isFinite(payload?.x);
  const hasY = Number.isFinite(payload?.y);
  const popupOptions = {
    window: win,
    ...(hasX ? { x: Math.max(0, Math.round(payload.x)) } : {}),
    ...(hasY ? { y: Math.max(0, Math.round(payload.y)) } : {}),
  };

  contextMenu.popup(popupOptions);
  logEvent('info', 'context menu opened', traceId, { filePath, hasX, hasY });

  return {
    ok: true,
  };
}

async function showSourceEditorContextMenu(event, payload, traceId) {
  const win = getEventWindow(event);
  if (!win || win.isDestroyed()) {
    throw new Error('window is not available');
  }

  const contextMenu = Menu.buildFromTemplate([
    {
      label: '剪切',
      click: () => {
        win.webContents.cut();
        logEvent('info', 'source context menu cut', traceId);
      },
    },
    {
      label: '复制',
      click: () => {
        win.webContents.copy();
        logEvent('info', 'source context menu copy', traceId);
      },
    },
    {
      label: '粘贴',
      click: () => {
        win.webContents.paste();
        logEvent('info', 'source context menu paste', traceId);
      },
    },
    {
      type: 'separator',
    },
    {
      label: '全选',
      click: () => {
        win.webContents.selectAll();
        logEvent('info', 'source context menu select all', traceId);
      },
    },
    {
      label: '查找',
      click: () => {
        sendMenuCommandToWindow(win, 'edit-search');
        logEvent('info', 'source context menu search', traceId);
      },
    },
  ]);

  const hasX = Number.isFinite(payload?.x);
  const hasY = Number.isFinite(payload?.y);
  const popupOptions = {
    window: win,
    ...(hasX ? { x: Math.max(0, Math.round(payload.x)) } : {}),
    ...(hasY ? { y: Math.max(0, Math.round(payload.y)) } : {}),
  };

  contextMenu.popup(popupOptions);
  logEvent('info', 'source context menu opened', traceId, { hasX, hasY });

  return {
    ok: true,
  };
}

function registerIpcHandler(channel, handler) {
  ipcMain.handle(channel, async (event, payload = {}) => {
    const traceId = normalizeTraceId(payload?.traceId);

    try {
      return await handler(event, payload, traceId);
    } catch (error) {
      logEvent('error', `ipc failed: ${channel}`, traceId, {
        error: safeErrorMessage(error),
      });
      throw error;
    }
  });
}

function getPrimaryWindow() {
  return BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0] || null;
}

function getWindowLaunchQueue(win) {
  if (!win || win.isDestroyed()) {
    return [];
  }

  const queueKey = win.webContents.id;
  const existingQueue = windowLaunchFileQueues.get(queueKey);
  if (existingQueue) {
    return existingQueue;
  }

  const queue = [];
  windowLaunchFileQueues.set(queueKey, queue);
  return queue;
}

function enqueueLaunchFileForWindow(win, filePath) {
  if (!win || win.isDestroyed() || !filePath) {
    return;
  }

  const queue = getWindowLaunchQueue(win);
  queue.push(filePath);
}

function consumeLaunchFileForWindow(win) {
  if (!win || win.isDestroyed()) {
    return '';
  }

  const queueKey = win.webContents.id;
  const queue = windowLaunchFileQueues.get(queueKey);
  if (!queue || queue.length === 0) {
    return '';
  }

  const filePath = queue.shift() || '';
  if (queue.length === 0) {
    windowLaunchFileQueues.delete(queueKey);
  }

  return filePath;
}

function focusWindow(win) {
  if (!win || win.isDestroyed()) {
    return;
  }

  if (win.isMinimized()) {
    win.restore();
  }

  win.focus();
}

function dispatchLaunchFilePathToWindow(win, filePath) {
  if (!win || win.isDestroyed() || !filePath) {
    return false;
  }

  if (win.webContents.isLoadingMainFrame()) {
    return false;
  }

  win.webContents.send('app:launchFile', {
    filePath,
  });
  return true;
}

function handleLaunchFilePath(filePath, options = {}) {
  if (!filePath) {
    return null;
  }

  if (options.openInNewWindow) {
    const win = createMainWindow(filePath);
    focusWindow(win);
    return win;
  }

  const targetWindow = options.targetWindow || getPrimaryWindow();
  if (!targetWindow) {
    pendingLaunchFilePaths.push(filePath);
    return null;
  }

  enqueueLaunchFileForWindow(targetWindow, filePath);
  const dispatched = dispatchLaunchFilePathToWindow(targetWindow, filePath);
  if (!dispatched) {
    focusWindow(targetWindow);
  }

  return targetWindow;
}

function registerIpcHandlers() {
  registerIpcHandler('app:getVersion', async (_event, _payload, traceId) => {
    logEvent('info', 'app getVersion', traceId);

    return {
      version: app.getVersion(),
    };
  });

  registerIpcHandler('app:logEvent', async (_event, payload, traceId) => {
    const level = payload?.level === 'error' ? 'error' : 'info';
    const message = typeof payload?.message === 'string' ? payload.message : 'renderer event';
    const detail = payload?.detail ?? null;

    logEvent(level, message, traceId, detail);

    return {
      ok: true,
    };
  });

  registerIpcHandler('app:consumeLaunchFilePath', async (_event, _payload, traceId) => {
    const win = getEventWindow(_event);
    const filePath = consumeLaunchFileForWindow(win) || pendingLaunchFilePaths.shift() || '';
    logEvent('info', 'consume launch file path', traceId, {
      hasLaunchFile: Boolean(filePath),
      filePath: filePath || '',
    });

    if (!filePath) {
      return null;
    }

    return {
      filePath,
    };
  });

  registerIpcHandler('file:openMarkdown', openMarkdownFile);
  registerIpcHandler('file:openFromPath', openMarkdownFromPath);
  registerIpcHandler('file:listMarkdownInDirectory', listMarkdownInDirectory);
  registerIpcHandler('file:saveMarkdown', saveMarkdownFile);
  registerIpcHandler('export:html', exportHtmlFile);
  registerIpcHandler('export:document', exportDocumentFile);
  registerIpcHandler('export:getCapabilities', getExportCapabilities);

  registerIpcHandler('window:minimize', async (event, payload, traceId) => {
    windowAction(event, payload, traceId, (win) => win.minimize());
    return { ok: true };
  });

  registerIpcHandler('window:toggleMaximize', async (event, payload, traceId) => {
    windowAction(event, payload, traceId, (win) => {
      if (win.isMaximized()) {
        win.unmaximize();
      } else {
        win.maximize();
      }
    });

    return {
      ok: true,
    };
  });

  registerIpcHandler('window:isMaximized', async (event, _payload, traceId) => {
    const win = getEventWindow(event);
    const isMaximized = Boolean(win?.isMaximized());

    logEvent('info', 'window isMaximized', traceId, { isMaximized });

    return {
      isMaximized,
    };
  });

  registerIpcHandler('window:close', async (event, payload, traceId) => {
    windowAction(event, payload, traceId, (win) => win.close());
    return { ok: true };
  });

  registerIpcHandler('window:toggleDevTools', async (event, payload, traceId) => {
    windowAction(event, payload, traceId, (win) => {
      if (win.webContents.isDevToolsOpened()) {
        win.webContents.closeDevTools();
      } else {
        win.webContents.openDevTools({ mode: 'detach' });
      }
    });

    return { ok: true };
  });

  registerIpcHandler('settings:getAll', settingsGetAll);
  registerIpcHandler('settings:getRecentFiles', settingsGetRecentFiles);
  registerIpcHandler('settings:get', settingsGet);
  registerIpcHandler('settings:set', settingsSet);
  registerIpcHandler('settings:setRecentFiles', settingsSetRecentFiles);
  registerIpcHandler('settings:selectUserCssFile', settingsSelectUserCssFile);
  registerIpcHandler('context:showFileTreeMenu', showFileTreeContextMenu);
  registerIpcHandler('context:showSourceEditorMenu', showSourceEditorContextMenu);
  registerIpcHandler('spell:getState', spellGetState);
  registerIpcHandler('spell:setEnabled', spellSetEnabled);
  registerIpcHandler('spell:setLanguage', spellSetLanguage);
}

function sendMenuCommandToWindow(target, command) {
  if (!target || target.isDestroyed()) {
    return;
  }

  target.webContents.send('menu:command', command);
}

function sendMenuCommand(command) {
  const target = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
  sendMenuCommandToWindow(target, command);
}

function resolveAppIconPath() {
  const packagedIconPath = path.join(process.resourcesPath, 'res', 'unireadmd.ico');
  if (app.isPackaged && fsSync.existsSync(packagedIconPath)) {
    return packagedIconPath;
  }

  return path.resolve(__dirname, '..', 'res', 'unireadmd.ico');
}

async function showAboutDialog(targetWindow) {
  const options = {
    type: 'info',
    title: '关于 UniReadMD',
    message: 'UniReadMD',
    detail: [
      `版本：${app.getVersion()}`,
      '问题请联系：terminator@leagsoft.com',
    ].join('\n'),
    buttons: ['确定'],
    noLink: true,
  };

  if (targetWindow && !targetWindow.isDestroyed()) {
    await dialog.showMessageBox(targetWindow, options);
    return;
  }

  await dialog.showMessageBox(options);
}

async function showSetDefaultMdReaderDialog(targetWindow) {
  const showDialog = async (options) => {
    if (targetWindow && !targetWindow.isDestroyed()) {
      await dialog.showMessageBox(targetWindow, options);
      return;
    }

    await dialog.showMessageBox(options);
  };

  if (process.platform !== 'win32') {
    await showDialog({
      type: 'info',
      title: '设为默认 md 阅读器',
      message: '当前功能仅支持 Windows 系统。',
      buttons: ['确定'],
      noLink: true,
    });
    return;
  }

  try {
    const result = await registerMarkdownFileAssociation();
    if (result.settingsUri) {
      await shell.openExternal(result.settingsUri);
    } else {
      await shell.openExternal('ms-settings:defaultapps');
    }

    await showDialog({
      type: 'info',
      title: '设为默认 md 阅读器',
      message: '已将 UniReadMD 注册为 .md 候选程序。',
      detail: [
        '系统默认应用设置页已打开。',
        '请在系统页面中将 UniReadMD 设为 .md 文档的默认打开方式。',
        '设置完成后，双击 .md 文档将默认使用 UniReadMD 打开。',
      ].join('\n'),
      buttons: ['确定'],
      noLink: true,
    });
  } catch (error) {
    await showDialog({
      type: 'error',
      title: '设为默认 md 阅读器',
      message: '设置默认 md 阅读器失败。',
      detail: safeErrorMessage(error),
      buttons: ['确定'],
      noLink: true,
    });
  }
}

function buildApplicationMenuTemplate() {
  const isMac = process.platform === 'darwin';
  const template = [];

  if (isMac) {
    template.push({
      role: 'appMenu',
    });
  }

  template.push({
    label: '文件',
    submenu: [
      {
        label: '打开...',
        accelerator: 'CmdOrCtrl+O',
        click: () => sendMenuCommand('file-open'),
      },
      {
        label: '保存',
        accelerator: 'CmdOrCtrl+S',
        click: () => sendMenuCommand('file-save'),
      },
      {
        label: '另存为...',
        accelerator: 'CmdOrCtrl+Shift+S',
        click: () => sendMenuCommand('file-save-as'),
      },
      {
        type: 'separator',
      },
      {
        label: '导出',
        submenu: [
          {
            label: '导出 HTML...',
            accelerator: 'CmdOrCtrl+E',
            click: () => sendMenuCommand('file-export-html'),
          },
          {
            label: '导出 PDF...',
            click: () => sendMenuCommand('file-export-pdf'),
          },
          {
            label: '导出 PNG...',
            click: () => sendMenuCommand('file-export-png'),
          },
          {
            label: '导出 JPEG...',
            click: () => sendMenuCommand('file-export-jpeg'),
          },
          {
            type: 'separator',
          },
          {
            label: '导出 Word (docx)...',
            click: () => sendMenuCommand('file-export-docx'),
          },
        ],
      },
      {
        type: 'separator',
      },
      isMac ? { role: 'close' } : { role: 'quit' },
    ],
  });

  template.push({
    label: '编辑',
    submenu: [
      { role: 'undo' },
      { role: 'redo' },
      { type: 'separator' },
      { role: 'cut' },
      { role: 'copy' },
      { role: 'paste' },
      { type: 'separator' },
      {
        label: '查找',
        accelerator: 'CmdOrCtrl+F',
        click: () => sendMenuCommand('edit-search'),
      },
      {
        label: '专注模式',
        click: () => sendMenuCommand('edit-focus'),
      },
      {
        label: '打字机模式',
        click: () => sendMenuCommand('edit-typewriter'),
      },
      { type: 'separator' },
      ...(isMac
        ? [
            { role: 'pasteAndMatchStyle' },
            { role: 'delete' },
            { type: 'separator' },
            { role: 'selectAll' },
          ]
        : [
            { role: 'delete' },
            { type: 'separator' },
            { role: 'selectAll' },
          ]),
    ],
  });

  template.push({
    label: '视图',
    submenu: [
      {
        label: '主题',
        submenu: [
          {
            label: '浅色',
            type: 'radio',
            click: () => sendMenuCommand('theme-light'),
          },
          {
            label: '深色',
            type: 'radio',
            click: () => sendMenuCommand('theme-dark'),
          },
          {
            label: '跟随系统',
            type: 'radio',
            checked: true,
            click: () => sendMenuCommand('theme-system'),
          },
        ],
      },
      { type: 'separator' },
      { role: 'reload' },
      { role: 'forceReload' },
      { role: 'toggleDevTools' },
      { type: 'separator' },
      { role: 'resetZoom' },
      { role: 'zoomIn' },
      { role: 'zoomOut' },
      { type: 'separator' },
      { role: 'togglefullscreen' },
    ],
  });

  template.push({
    label: '窗口',
    submenu: isMac
      ? [{ role: 'minimize' }, { role: 'zoom' }, { type: 'separator' }, { role: 'front' }]
      : [{ role: 'minimize' }, { role: 'close' }],
  });

  template.push({
    label: '帮助',
    submenu: [
      {
        label: '设为默认 md 阅读器',
        click: () => {
          const target = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
          void showSetDefaultMdReaderDialog(target);
        },
      },
      {
        type: 'separator',
      },
      {
        label: '关于 UniReadMD',
        click: () => {
          const target = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
          void showAboutDialog(target);
        },
      },
      {
        type: 'separator',
      },
      {
        label: '拼写检查',
        click: () => sendMenuCommand('help-spell'),
      },
    ],
  });

  return template;
}

function setupApplicationMenu() {
  const menu = Menu.buildFromTemplate(buildApplicationMenuTemplate());
  Menu.setApplicationMenu(menu);
}

function createMainWindow(initialFilePath = '') {
  const preloadPath = path.join(__dirname, 'preload.cjs');
  const preloadExists = fsSync.existsSync(preloadPath);
  logEvent('info', 'window preload path resolved', createTraceId('boot'), {
    preloadPath,
    preloadExists,
    isDev,
  });

  const appIconPath = resolveAppIconPath();
  const win = new BrowserWindow({
    width: 1320,
    height: 880,
    minWidth: 980,
    minHeight: 640,
    title: 'UniReadMD',
    icon: appIconPath,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      webSecurity: false,
      preload: preloadPath,
    },
  });
  const windowWebContentsId = win.webContents.id;
  const windowSession = win.webContents.session;

  if (initialFilePath) {
    enqueueLaunchFileForWindow(win, initialFilePath);
  }

  win.webContents.setWindowOpenHandler(({ url }) => {
    if (isExternalNavigationUrl(url)) {
      void shell.openExternal(url);
      return {
        action: 'deny',
      };
    }

    return {
      action: 'deny',
    };
  });

  win.webContents.on('will-navigate', (event, url) => {
    if (!isExternalNavigationUrl(url)) {
      return;
    }

    event.preventDefault();
    void shell.openExternal(url);
  });

  win.webContents.on('did-finish-load', () => {
    const traceId = createTraceId('boot');
    logEvent('info', 'renderer did-finish-load', traceId);

    const autoMenuCommand = process.env.UNIREADMD_AUTO_MENU_COMMAND;
    if (typeof autoMenuCommand === 'string' && autoMenuCommand.trim()) {
      const command = autoMenuCommand.trim();
      setTimeout(() => {
        logEvent('info', 'auto menu command dispatch', traceId, { command });
        win.webContents.send('menu:command', command);
      }, 800);
    }
  });

  win.webContents.on('did-fail-load', (_event, code, description, url, isMainFrame) => {
    logEvent('error', 'renderer did-fail-load', createTraceId('boot'), {
      code,
      description,
      url,
      isMainFrame,
    });
  });

  win.webContents.on('preload-error', (_event, preloadPathValue, error) => {
    logEvent('error', 'renderer preload-error', createTraceId('boot'), {
      preloadPath: preloadPathValue,
      error: safeErrorMessage(error),
    });
  });

  win.webContents.on('console-message', (_event, level, message, line, sourceId) => {
    if (level < 2) {
      return;
    }
    logEvent('error', 'renderer console error', createTraceId('boot'), {
      level,
      message,
      line,
      sourceId,
    });
  });

  win.on('closed', () => {
    windowLaunchFileQueues.delete(windowWebContentsId);
  });

  win.loadURL(getRendererUrl());
  void readSettings()
    .then((settings) => {
      const state = toSpellCheckState(settings, windowSession);
      applySpellCheckState(windowSession, state, createTraceId('spell'));
    })
    .catch((error) => {
      logEvent('error', 'init spell check failed', createTraceId('spell'), {
        error: safeErrorMessage(error),
      });
    });

  if (isDev) {
    win.webContents.openDevTools({ mode: 'detach' });
  }

  return win;
}

app.whenReady().then(async () => {
  await ensureRuntimePaths();
  await startRendererStaticServer();
  registerIpcHandlers();
  setupApplicationMenu();

  logEvent('info', 'app ready', createTraceId('boot'));
  if (pendingLaunchFilePaths.length > 0) {
    logEvent('info', 'launch markdown detected', createTraceId('boot'), {
      filePaths: pendingLaunchFilePaths,
    });
  }
  createMainWindow(pendingLaunchFilePaths.shift() || '');

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    }
  });
});

app.on('second-instance', (_event, argv) => {
  const filePath = resolveLaunchMarkdownPath(argv);
  if (!filePath) {
    const win = getPrimaryWindow() || createMainWindow();
    focusWindow(win);
    return;
  }

  handleLaunchFilePath(filePath, {
    openInNewWindow: true,
  });
});

app.on('open-file', (event, filePath) => {
  event.preventDefault();

  const normalizedPath = normalizeCliMarkdownPath(filePath);
  if (!normalizedPath) {
    return;
  }

  if (!app.isReady()) {
    pendingLaunchFilePaths.push(normalizedPath);
    return;
  }

  handleLaunchFilePath(normalizedPath, {
    openInNewWindow: true,
  });
});

app.on('window-all-closed', () => {
  if (rendererStaticServer) {
    rendererStaticServer.close();
    rendererStaticServer = null;
    rendererStaticServerUrl = '';
  }

  if (process.platform !== 'darwin') {
    app.quit();
  }
});
