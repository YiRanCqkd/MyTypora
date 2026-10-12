const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const { listPackage, extractFile } = require('@electron/asar');

const projectRoot = path.resolve(__dirname, '..', '..');
const unpackedDirectory = process.env.UNIREADMD_PACK_DIR
  || path.join(projectRoot, 'release', 'electron', 'win-unpacked');

test('Given unpacked Windows 应用, when 检查发布内容, then 主文件、preload 和资源存在', () => {
  const executable = path.join(unpackedDirectory, 'UniReadMD.exe');
  const archive = path.join(unpackedDirectory, 'resources', 'app.asar');
  const icon = path.join(unpackedDirectory, 'resources', 'res', 'unireadmd.ico');

  assert.equal(fs.existsSync(executable), true, '缺少 unpacked 主程序');
  assert.equal(fs.existsSync(archive), true, '缺少 app.asar');
  assert.equal(fs.existsSync(icon), true, '缺少打包资源图标');

  const archiveFiles = listPackage(archive).map((entry) => entry.replaceAll('\\', '/'));
  assert.equal(archiveFiles.includes('/electron/main.cjs'), true, '缺少 Electron 主文件');
  assert.equal(archiveFiles.includes('/electron/preload.cjs'), true, '缺少 preload 文件');
});

test('发布包包含本次主进程代码，搜索线程可从解包目录执行', async () => {
  const archive = path.join(unpackedDirectory, 'resources', 'app.asar');
  for (const entry of ['main.cjs', 'preload.cjs',
    ...fs.readdirSync(path.join(projectRoot, 'electron', 'lib'))
      .filter((name) => name.endsWith('.cjs')).map((name) => `lib/${name}`)]) {
    const archivePath = path.join('electron', ...entry.split('/'));
    assert.deepEqual(extractFile(archive, archivePath),
      fs.readFileSync(path.join(projectRoot, 'electron', entry)), `发布代码不一致：${entry}`);
  }
  const worker = path.join(unpackedDirectory, 'resources',
    'app.asar.unpacked', 'electron', 'lib', 'search-lines.cjs');
  assert.equal(fs.existsSync(worker), true, '缺少解包搜索线程');
  const { searchLines } = require(worker);
  assert.deepEqual(await searchLines(['测试', '测试文档', 'x 测试 y'], '测试', {
    wholeWord: true,
  }), [0, 2]);
});

test('发布包前端入口与当前构建一致，避免打包旧编辑器', () => {
  const archive = path.join(unpackedDirectory, 'resources', 'app.asar');
  const dist = path.join(projectRoot, 'frontend', 'dist');
  for (const entry of ['index.html',
    ...fs.readdirSync(path.join(dist, 'assets')).map((name) => `assets/${name}`)]) {
    assert.deepEqual(extractFile(archive, path.join('frontend', 'dist', ...entry.split('/'))),
      fs.readFileSync(path.join(dist, ...entry.split('/'))), `前端产物不一致：${entry}`);
  }
});
