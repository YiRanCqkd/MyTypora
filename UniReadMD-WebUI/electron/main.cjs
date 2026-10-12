const {
  app, BrowserWindow, Menu, clipboard, dialog, ipcMain, nativeImage, shell,
} = require('electron');
const { spawn } = require('node:child_process');
const http = require('node:http');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const { createAppDialogs } = require('./lib/app-dialogs.cjs');
const { createAssetService } = require('./lib/assets.cjs');
const { createSessionService } = require('./lib/sessions.cjs');
const { validateIpcSender } = require('./lib/security.cjs');
const { createAuthorizationService } = require('./lib/authorization.cjs');
const { createContextMenuService } = require('./lib/context-menus.cjs');
const { createExportEngines } = require('./lib/export-engines.cjs');
const { createExportsService } = require('./lib/exports.cjs');
const { createFilesService } = require('./lib/files.cjs');
const { createWorkspaceTools } = require('./lib/workspace-tools.cjs');
const { createImageService } = require('./lib/images.cjs');
const {
  createIpcController,
  sendMenuCommand,
  sendMenuCommandToWindow,
} = require('./lib/ipc.cjs');
const { createLaunchService } = require('./lib/launch.cjs');
const { setupApplicationMenu } = require('./lib/menus.cjs');
const { createRegistryService } = require('./lib/registry.cjs');
const { createRendererServer } = require('./lib/renderer-server.cjs');
const { createRuntime } = require('./lib/runtime.cjs');
const { createSettingsService } = require('./lib/settings.cjs');
const { createSpellService } = require('./lib/spell.cjs');
const { createWindowService } = require('./lib/windows.cjs');

if (process.env.UNIREADMD_TEST_MODE === '1' && process.env.UNIREADMD_TEST_USER_DATA) {
  app.setPath('userData', process.env.UNIREADMD_TEST_USER_DATA);
}

const rootDirectory = path.resolve(__dirname, '..');
const isDev = Boolean(process.env.ELECTRON_START_URL);
const runtime = createRuntime({ app, fs, path, spawn });
const authorization = createAuthorizationService();
const getEventWindow = (event) => (
  BrowserWindow.fromWebContents(event.sender) || BrowserWindow.getFocusedWindow()
);
const assets = createAssetService({ fs, getServerUrl: () => rendererServer.getUrl() });
const sessions = createSessionService({
  fs, authorization, dialog, getEventWindow,
  directory: path.join(app.getPath('userData'), 'sessions'),
});
const rendererServer = createRendererServer({
  fsSync,
  http,
  isDev,
  log: runtime.log,
  path,
  rootDirectory,
  readAsset: assets.read,
});
const settings = createSettingsService({
  authorization,
  dialog,
  fs,
  getEventWindow,
  getFilePath: runtime.getSettingsFilePath,
  log: runtime.log,
  enforceAuthorization: true,
  onChanged: ({ key, value }) => {
    const name = { documentToolsVisible: 'document-tools', preferencesVisible: 'preferences' }[key];
    if (!name) return;
    const item = Menu.getApplicationMenu()?.getMenuItemById(`view-${name}`);
    if (item) item.checked = value;
    for (const win of BrowserWindow.getAllWindows()) {
      sendMenuCommandToWindow(win, `view-${name}-${value ? 'show' : 'hide'}`);
    }
  },
});
const spell = createSpellService({ getEventWindow, log: runtime.log, settings });
const files = createFilesService({
  assets,
  authorization,
  dialog,
  fs,
  getEventWindow,
  log: runtime.log,
});
const workspaceTools = createWorkspaceTools({
  fs, files, assets, authorization, shell, dialog, getEventWindow,
});
const images = createImageService({
  fs, files, assets, authorization, dialog, clipboard, nativeImage, getEventWindow,
  approvalsPath: path.join(app.getPath('userData'), 'image-root-approvals.json'),
});
const engines = createExportEngines({
  app,
  BrowserWindow,
  fs,
  log: runtime.log,
  path,
  runCommand: runtime.runCommand,
});
const exporters = createExportsService({
  files,
  assets,
  authorization,
  dialog,
  engines,
  fs,
  getEventWindow,
  log: runtime.log,
  path,
});
let windows;
const launch = createLaunchService({
  BrowserWindow,
  authorization,
  createWindow: (filePath) => windows.create(filePath),
  fsSync,
  path,
});
const contextMenus = createContextMenuService({
  authorization,
  Menu,
  clipboard,
  getEventWindow,
  log: runtime.log,
  sendCommandToWindow: (win, command) => sendMenuCommandToWindow(win, command),
  shell,
});
const ipc = createIpcController({
  workspaceTools,
  images,
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
  sessions,
  validateSender: (event) => validateIpcSender(event, BrowserWindow, windows.getRendererUrl),
});
windows = createWindowService({
  assets,
  sessions,
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
});
const registry = createRegistryService({
  app,
  fsSync,
  path,
  resolveFallbackIcon: windows.resolveIconPath,
  runCommand: runtime.runCommand,
});
const appDialogs = createAppDialogs({ app, dialog, registry, shell });

function getPrimaryWindow() {
  return BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0] || null;
}

async function setupMenu(pandocFormats = []) {
  const getTarget = () => getPrimaryWindow();
  const { settings: preferences } = await settings.getAll();
  setupApplicationMenu(Menu, {
    documentToolsVisible: preferences.documentToolsVisible,
    preferencesVisible: preferences.preferencesVisible,
    onAbout: () => void appDialogs.showAbout(getTarget()),
    onSetDefaultReader: () => void appDialogs.showSetDefaultReader(getTarget()),
    sendCommand: (command) => sendMenuCommand(BrowserWindow, command),
    pandocFormats,
  });
}

const initialLaunchFilePath = launch.resolve(process.argv);
if (initialLaunchFilePath) launch.addPending(initialLaunchFilePath);
const hasSingleInstanceLock = app.requestSingleInstanceLock();
if (!hasSingleInstanceLock) app.quit();

app.whenReady().then(async () => {
  await runtime.initializePaths();
  await sessions.initialize();
  await rendererServer.start();
  ipc.registerAll();
  await setupMenu();
  runtime.log('info', 'app ready', runtime.createTraceId('boot'));
  if (launch.getPending().length > 0) {
    runtime.log('info', 'launch markdown detected', runtime.createTraceId('boot'), {
      filePaths: launch.getPending(),
    });
  }
  windows.create(launch.consumePending());
  while (sessions.getRecoveryCount() > 0) windows.create();
  void engines.getCapabilities().then((capabilities) => setupMenu(capabilities.pandocFormats));
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) windows.create();
  });
});

app.on('second-instance', (_event, argv) => {
  const filePath = launch.resolve(argv);
  if (filePath) launch.handle(filePath, { openInNewWindow: true });
  else launch.focus(getPrimaryWindow() || windows.create());
});

app.on('open-file', (event, filePath) => {
  event.preventDefault();
  const normalizedPath = launch.resolve(['app', filePath]);
  if (!normalizedPath) return;
  if (!app.isReady()) launch.addPending(normalizedPath);
  else launch.handle(normalizedPath, { openInNewWindow: true });
});

app.on('window-all-closed', () => {
  rendererServer.close();
  if (process.platform !== 'darwin') app.quit();
});
