const { app, BrowserWindow, Menu, clipboard, dialog, ipcMain, shell } = require('electron');
const { spawn } = require('node:child_process');
const http = require('node:http');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const { createAppDialogs } = require('./lib/app-dialogs.cjs');
const { createAuthorizationService } = require('./lib/authorization.cjs');
const { createContextMenuService } = require('./lib/context-menus.cjs');
const { createExportEngines } = require('./lib/export-engines.cjs');
const { createExportsService } = require('./lib/exports.cjs');
const { createFilesService } = require('./lib/files.cjs');
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

const rootDirectory = path.resolve(__dirname, '..');
const isDev = Boolean(process.env.ELECTRON_START_URL);
const runtime = createRuntime({ app, fs, path, spawn });
const authorization = createAuthorizationService();
const getEventWindow = (event) => (
  BrowserWindow.fromWebContents(event.sender) || BrowserWindow.getFocusedWindow()
);
const rendererServer = createRendererServer({
  fsSync,
  http,
  isDev,
  log: runtime.log,
  path,
  rootDirectory,
});
const settings = createSettingsService({
  authorization,
  dialog,
  fs,
  getEventWindow,
  getFilePath: runtime.getSettingsFilePath,
  log: runtime.log,
});
const spell = createSpellService({ getEventWindow, log: runtime.log, settings });
const files = createFilesService({
  authorization,
  dialog,
  fs,
  getEventWindow,
  log: runtime.log,
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
  Menu,
  clipboard,
  getEventWindow,
  log: runtime.log,
  sendCommandToWindow: (win, command) => sendMenuCommandToWindow(win, command),
  shell,
});
const ipc = createIpcController({
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
});
windows = createWindowService({
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

function setupMenu() {
  const getTarget = () => getPrimaryWindow();
  setupApplicationMenu(Menu, {
    onAbout: () => void appDialogs.showAbout(getTarget()),
    onSetDefaultReader: () => void appDialogs.showSetDefaultReader(getTarget()),
    sendCommand: (command) => sendMenuCommand(BrowserWindow, command),
  });
}

const initialLaunchFilePath = launch.resolve(process.argv);
if (initialLaunchFilePath) launch.addPending(initialLaunchFilePath);
const hasSingleInstanceLock = app.requestSingleInstanceLock();
if (!hasSingleInstanceLock) app.quit();

app.whenReady().then(async () => {
  await runtime.initializePaths();
  await rendererServer.start();
  ipc.registerAll();
  setupMenu();
  runtime.log('info', 'app ready', runtime.createTraceId('boot'));
  if (launch.getPending().length > 0) {
    runtime.log('info', 'launch markdown detected', runtime.createTraceId('boot'), {
      filePaths: launch.getPending(),
    });
  }
  windows.create(launch.consumePending());
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
