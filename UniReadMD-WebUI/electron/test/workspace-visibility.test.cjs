const assert = require('node:assert/strict');
const test = require('node:test');
const { buildApplicationMenuTemplate } = require('../lib/menus.cjs');
const { mergeSettings, validateSettingUpdate } = require('../lib/settings.cjs');

test('视图显示开关默认关闭，读取旧设置时采用隐藏基线', () => {
  const settings = mergeSettings({ theme: 'dark', documentToolsVisible: 'true' });
  assert.equal(settings.documentToolsVisible, false);
  assert.equal(settings.preferencesVisible, false);
  assert.equal(settings.theme, 'dark');
  for (const key of ['documentToolsVisible', 'preferencesVisible']) {
    assert.equal(validateSettingUpdate(key, true).value, true);
    assert.throws(() => validateSettingUpdate(key, 'true'));
  }
});

test('视图菜单提供两个独立勾选项和对应命令', () => {
  const commands = [];
  const view = buildApplicationMenuTemplate({ documentToolsVisible: true,
    sendCommand: (command) => commands.push(command) })
    .find((item) => item.label === '视图');
  const toolbar = view.submenu.find((item) => item.id === 'view-document-tools');
  const preferences = view.submenu.find((item) => item.id === 'view-preferences');
  assert.equal(toolbar.type, 'checkbox');
  assert.equal(toolbar.checked, true);
  assert.equal(preferences.checked, false);
  toolbar.click();
  preferences.click();
  assert.deepEqual(commands, ['view-toggle-document-tools', 'view-toggle-preferences']);
});
