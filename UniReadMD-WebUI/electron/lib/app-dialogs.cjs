const { safeErrorMessage } = require('./errors.cjs');

function createAppDialogs(options) {
  const { app, dialog, registry, shell } = options;

  async function show(target, dialogOptions) {
    if (target && !target.isDestroyed()) return dialog.showMessageBox(target, dialogOptions);
    return dialog.showMessageBox(dialogOptions);
  }

  async function showAbout(target) {
    return show(target, {
      type: 'info',
      title: '关于 UniReadMD',
      message: 'UniReadMD',
      detail: [`版本：${app.getVersion()}`, '问题请联系：terminator@leagsoft.com'].join('\n'),
      buttons: ['确定'],
      noLink: true,
    });
  }

  async function showSetDefaultReader(target) {
    if (process.platform !== 'win32') {
      await show(target, {
        type: 'info',
        title: '设为默认 md 阅读器',
        message: '当前功能仅支持 Windows 系统。',
        buttons: ['确定'],
        noLink: true,
      });
      return;
    }
    try {
      const result = await registry.register();
      await shell.openExternal(result.settingsUri || 'ms-settings:defaultapps');
      await show(target, {
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
      await show(target, {
        type: 'error',
        title: '设为默认 md 阅读器',
        message: '设置默认 md 阅读器失败。',
        detail: safeErrorMessage(error),
        buttons: ['确定'],
        noLink: true,
      });
    }
  }

  return { showAbout, showSetDefaultReader };
}

module.exports = { createAppDialogs };
