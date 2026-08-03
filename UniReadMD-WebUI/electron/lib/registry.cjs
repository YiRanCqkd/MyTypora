const APP_NAME = 'UniReadMD';
const DESCRIPTION = 'UniReadMD Markdown 阅读器';
const PROG_ID = 'UniReadMD.AssocFile.Markdown';
const CAPABILITIES_PATH = 'Software\\UniReadMD\\Capabilities';
const ICON_FILE_NAME = 'UniReadMD.ico';

function createRegistryService(options) {
  const { app, fsSync, path, resolveFallbackIcon, runCommand } = options;

  async function addValue(keyPath, name, type, data) {
    const arguments_ = ['add', keyPath, name === null ? '/ve' : '/v'];
    if (name !== null) arguments_.push(name);
    if (type) arguments_.push('/t', type);
    if (data !== undefined && data !== null) arguments_.push('/d', data);
    arguments_.push('/f');
    await runCommand('reg.exe', arguments_, 15_000);
  }

  function getExecutablePath() {
    const portable = process.env.PORTABLE_EXECUTABLE_FILE || '';
    if (portable && fsSync.existsSync(portable)) return portable;
    if (app.isPackaged) return process.execPath;
    return path.resolve(__dirname, '..', '..', 'release', 'electron', 'UniReadMD.exe');
  }

  function getIconPath() {
    const portableDirectory = process.env.PORTABLE_EXECUTABLE_DIR || '';
    const candidate = portableDirectory
      ? path.join(portableDirectory, ICON_FILE_NAME)
      : app.isPackaged
        ? path.join(path.dirname(process.execPath), ICON_FILE_NAME)
        : path.resolve(__dirname, '..', '..', 'release', 'electron', ICON_FILE_NAME);
    return fsSync.existsSync(candidate) ? candidate : resolveFallbackIcon();
  }

  async function refreshIconCache() {
    if (process.platform !== 'win32') return;
    for (const arguments_ of [['-show'], ['-ClearIconCache']]) {
      try {
        await runCommand('ie4uinit.exe', arguments_, 15_000);
      } catch {
        // 文件关联已经写入，刷新图标失败不影响主流程。
      }
    }
  }

  async function register() {
    if (process.platform !== 'win32') {
      return { appName: APP_NAME, settingsUri: '', requiresUserConfirmation: false };
    }
    const executable = getExecutablePath();
    if (!fsSync.existsSync(executable)) throw new Error(`executable not found: ${executable}`);
    const icon = getIconPath();
    const command = `"${executable}" "%1"`;
    const appKey = 'HKCU\\Software\\Classes\\Applications\\UniReadMD.exe';
    const progIdKey = `HKCU\\Software\\Classes\\${PROG_ID}`;
    const capabilitiesKey = `HKCU\\${CAPABILITIES_PATH}`;
    const values = [
      [appKey, 'FriendlyAppName', 'REG_SZ', APP_NAME],
      [appKey, 'DefaultIcon', 'REG_SZ', icon],
      [`${appKey}\\shell\\open\\command`, null, 'REG_SZ', command],
      [`${appKey}\\SupportedTypes`, '.md', 'REG_SZ', ''],
      [`${appKey}\\SupportedTypes`, '.markdown', 'REG_SZ', ''],
      [progIdKey, null, 'REG_SZ', DESCRIPTION],
      [progIdKey, 'FriendlyTypeName', 'REG_SZ', DESCRIPTION],
      [`${progIdKey}\\DefaultIcon`, null, 'REG_SZ', icon],
      [`${progIdKey}\\shell\\open\\command`, null, 'REG_SZ', command],
      [capabilitiesKey, 'ApplicationName', 'REG_SZ', APP_NAME],
      [capabilitiesKey, 'ApplicationIcon', 'REG_SZ', icon],
      [capabilitiesKey, 'ApplicationDescription', 'REG_SZ', DESCRIPTION],
      [`${capabilitiesKey}\\FileAssociations`, '.md', 'REG_SZ', PROG_ID],
      [`${capabilitiesKey}\\FileAssociations`, '.markdown', 'REG_SZ', PROG_ID],
      ['HKCU\\Software\\RegisteredApplications', APP_NAME, 'REG_SZ', CAPABILITIES_PATH],
      ['HKCU\\Software\\Classes\\.md\\OpenWithProgids', PROG_ID, 'REG_SZ', ''],
      ['HKCU\\Software\\Classes\\.markdown\\OpenWithProgids', PROG_ID, 'REG_SZ', ''],
    ];
    for (const value of values) await addValue(...value);
    await refreshIconCache();
    return {
      appName: APP_NAME,
      settingsUri: `ms-settings:defaultapps?registeredAppUser=${encodeURIComponent(APP_NAME)}`,
      requiresUserConfirmation: true,
    };
  }

  return { register };
}

module.exports = { createRegistryService };
