const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const { listPackage } = require('@electron/asar');

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
