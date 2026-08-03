const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const projectRoot = path.resolve(__dirname, '..', '..');

test('Given Electron manifest, when resolving entry files, then main and preload exist', () => {
  const manifest = JSON.parse(
    fs.readFileSync(path.join(projectRoot, 'package.json'), 'utf8'),
  );
  const mainFile = path.join(projectRoot, manifest.main);
  const preloadFile = path.join(projectRoot, 'electron', 'preload.cjs');

  assert.equal(manifest.main, 'electron/main.cjs');
  assert.equal(fs.existsSync(mainFile), true);
  assert.equal(fs.existsSync(preloadFile), true);
});
