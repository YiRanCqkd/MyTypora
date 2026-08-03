const DEFAULT_SETTINGS = Object.freeze({
  theme: 'light',
  fontSize: 14,
  contentZoom: 100,
  viewMode: 'split',
  autoSave: false,
  userCssEnabled: false,
  userCssPath: '',
  spellCheckEnabled: true,
  spellCheckLanguage: 'en-US',
});

const ALLOWED_SETTINGS = new Set(Object.keys(DEFAULT_SETTINGS));

const { createContractError } = require('./errors.cjs');

function validateSettingUpdate(key, value) {
  if (!ALLOWED_SETTINGS.has(key)) {
    throw createContractError('ERR_INVALID_SETTING_KEY', '设置 key 不在白名单中');
  }

  let normalizedValue = value;
  const invalid = () => {
    throw createContractError('ERR_INVALID_SETTING_VALUE', `设置 ${key} 的值无效`);
  };
  if (key === 'theme' && !['light', 'dark', 'system'].includes(value)) invalid();
  if (key === 'fontSize') {
    if (typeof value !== 'number' || value < 10 || value > 28) invalid();
    normalizedValue = Math.round(value);
  }
  if (key === 'contentZoom') {
    if (typeof value !== 'number' || value < 50 || value > 200) invalid();
    normalizedValue = Math.round(value);
  }
  if (key === 'viewMode' && !['edit', 'preview', 'split'].includes(value)) invalid();
  if (['autoSave', 'userCssEnabled', 'spellCheckEnabled'].includes(key)) {
    if (typeof value !== 'boolean') invalid();
  }
  if (key === 'userCssPath' && (typeof value !== 'string' || value.length > 1024)) invalid();
  if (
    key === 'spellCheckLanguage'
    && (typeof value !== 'string' || value.length > 64)
  ) invalid();

  return { key, value: normalizedValue };
}

function mergeSettings(raw) {
  const settings = { ...DEFAULT_SETTINGS };
  if (!raw || typeof raw !== 'object') return settings;
  for (const key of ALLOWED_SETTINGS) {
    try {
      settings[key] = validateSettingUpdate(key, raw[key]).value;
    } catch {
      // 旧设置中的单个非法值不会影响其他有效设置。
    }
  }
  return settings;
}

function normalizeRecentFiles(raw) {
  if (!Array.isArray(raw)) return [];
  const seen = new Set();
  const output = [];
  for (const item of raw) {
    if (typeof item !== 'string') continue;
    const filePath = item.trim();
    const key = filePath.toLowerCase();
    if (!filePath || !/\.(md|markdown)$/i.test(filePath) || seen.has(key)) continue;
    seen.add(key);
    output.push(filePath);
  }
  return output.slice(0, 12);
}

function createSettingsService(options) {
  const { authorization, dialog, fs, getEventWindow, getFilePath, log = () => {} } = options;
  let settingsCache = null;
  let recentFilesCache = null;

  async function readStorage() {
    if (settingsCache && recentFilesCache) {
      return { settings: settingsCache, recentFiles: recentFilesCache };
    }
    try {
      const parsed = JSON.parse(await fs.readFile(getFilePath(), 'utf8'));
      settingsCache = mergeSettings(parsed);
      recentFilesCache = normalizeRecentFiles(parsed?.recentFiles);
    } catch {
      settingsCache = { ...DEFAULT_SETTINGS };
      recentFilesCache = [];
    }
    return { settings: settingsCache, recentFiles: recentFilesCache };
  }

  async function persist() {
    const storage = await readStorage();
    const payload = { ...storage.settings, recentFiles: storage.recentFiles };
    await fs.writeFile(getFilePath(), JSON.stringify(payload, null, 2), 'utf8');
  }

  async function getAll(_event, _payload, traceId) {
    const { settings } = await readStorage();
    log('info', 'settings getAll', traceId);
    return { settings };
  }

  async function getRecentFiles(event, _payload, traceId) {
    const { recentFiles } = await readStorage();
    authorization.grantMany(event, recentFiles);
    log('info', 'settings getRecentFiles', traceId, { count: recentFiles.length });
    return { filePaths: recentFiles };
  }

  async function get(_event, payload) {
    const { key } = validateSettingUpdate(payload?.key, DEFAULT_SETTINGS[payload?.key]);
    const { settings } = await readStorage();
    return { key, value: settings[key] };
  }

  async function set(_event, payload, traceId) {
    const update = validateSettingUpdate(payload?.key, payload?.value);
    const storage = await readStorage();
    storage.settings = { ...storage.settings, [update.key]: update.value };
    settingsCache = storage.settings;
    await persist();
    log('info', 'settings set', traceId, update);
    return { ...update, settings: settingsCache };
  }

  async function setRecentFiles(_event, payload, traceId) {
    recentFilesCache = normalizeRecentFiles(payload?.filePaths);
    await persist();
    log('info', 'settings setRecentFiles', traceId, { count: recentFilesCache.length });
    return { filePaths: recentFilesCache };
  }

  async function selectUserCssFile(event, _payload, traceId) {
    const win = getEventWindow(event);
    if (win && !win.isDestroyed()) {
      if (win.isMinimized()) win.restore();
      win.focus();
    }
    const options_ = { properties: ['openFile'], filters: [{ name: 'CSS', extensions: ['css'] }] };
    const result = win && !win.isDestroyed()
      ? await dialog.showOpenDialog(win, options_)
      : await dialog.showOpenDialog(options_);
    if (result.canceled || result.filePaths.length === 0) {
      log('info', 'select user css canceled', traceId);
      return null;
    }
    return { filePath: result.filePaths[0] };
  }

  return {
    get,
    getAll,
    getRecentFiles,
    readSettings: async () => (await readStorage()).settings,
    selectUserCssFile,
    set,
    setRecentFiles,
  };
}

module.exports = {
  ALLOWED_SETTINGS,
  DEFAULT_SETTINGS,
  createSettingsService,
  mergeSettings,
  normalizeRecentFiles,
  validateSettingUpdate,
};
