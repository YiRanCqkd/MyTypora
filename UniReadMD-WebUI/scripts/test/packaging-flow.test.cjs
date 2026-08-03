const assert = require('node:assert/strict');
const test = require('node:test');
const builderConfig = require('../../electron/builder.json');
const manifest = require('../../package.json');

test('Given 普通 Electron 合同, when 无打包产物, then 不运行 packaging 合同', () => {
  assert.doesNotMatch(manifest.scripts['test:electron'], /packaging\.test\.cjs/u);
  assert.doesNotMatch(manifest.scripts['test:electron'], /\*\.test\.cjs/u);
});

test('Given pack:dir, when builder 成功, then 显式运行 packaging 合同', () => {
  assert.equal(
    manifest.scripts['test:packaging'],
    'node --test electron/test/packaging.test.cjs',
  );
  assert.match(manifest.scripts['pack:dir'], /npm run test:packaging/u);
});

test('Given 已安装 Electron, when pack:dir, then 使用本地 dist 且不重复下载', () => {
  assert.equal(builderConfig.electronDist, 'node_modules/electron/dist');
});
