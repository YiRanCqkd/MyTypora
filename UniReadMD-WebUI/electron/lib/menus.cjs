function buildApplicationMenuTemplate(options = {}) {
  const {
    isMac = process.platform === 'darwin',
    onAbout = () => {},
    onSetDefaultReader = () => {},
    sendCommand = () => {},
  } = options;
  const template = isMac ? [{ role: 'appMenu' }] : [];

  template.push({
    label: '文件',
    submenu: [
      { label: '打开...', accelerator: 'CmdOrCtrl+O', click: () => sendCommand('file-open') },
      { label: '保存', accelerator: 'CmdOrCtrl+S', click: () => sendCommand('file-save') },
      {
        label: '另存为...',
        accelerator: 'CmdOrCtrl+Shift+S',
        click: () => sendCommand('file-save-as'),
      },
      { type: 'separator' },
      {
        label: '导出',
        submenu: [
          {
            label: '导出 HTML...',
            accelerator: 'CmdOrCtrl+E',
            click: () => sendCommand('file-export-html'),
          },
          { label: '导出 PDF...', click: () => sendCommand('file-export-pdf') },
          { label: '导出 PNG...', click: () => sendCommand('file-export-png') },
          { label: '导出 JPEG...', click: () => sendCommand('file-export-jpeg') },
          { type: 'separator' },
          { label: '导出 Word (docx)...', click: () => sendCommand('file-export-docx') },
        ],
      },
      { type: 'separator' },
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
      { label: '查找', accelerator: 'CmdOrCtrl+F', click: () => sendCommand('edit-search') },
      { label: '专注模式', click: () => sendCommand('edit-focus') },
      { label: '打字机模式', click: () => sendCommand('edit-typewriter') },
      { type: 'separator' },
      ...(isMac
        ? [
            { role: 'pasteAndMatchStyle' },
            { role: 'delete' },
            { type: 'separator' },
            { role: 'selectAll' },
          ]
        : [{ role: 'delete' }, { type: 'separator' }, { role: 'selectAll' }]),
    ],
  });
  template.push({
    label: '视图',
    submenu: [
      {
        label: '主题',
        submenu: [
          { label: '浅色', type: 'radio', click: () => sendCommand('theme-light') },
          { label: '深色', type: 'radio', click: () => sendCommand('theme-dark') },
          {
            label: '跟随系统',
            type: 'radio',
            checked: true,
            click: () => sendCommand('theme-system'),
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
      { label: '设为默认 md 阅读器', click: onSetDefaultReader },
      { type: 'separator' },
      { label: '关于 UniReadMD', click: onAbout },
      { type: 'separator' },
      { label: '拼写检查', click: () => sendCommand('help-spell') },
    ],
  });
  return template;
}

function setupApplicationMenu(Menu, options) {
  const menu = Menu.buildFromTemplate(buildApplicationMenuTemplate(options));
  Menu.setApplicationMenu(menu);
  return menu;
}

module.exports = {
  buildApplicationMenuTemplate,
  setupApplicationMenu,
};
