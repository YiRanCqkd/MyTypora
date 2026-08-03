const { spawnSync } = require('node:child_process');
const fsPromises = require('node:fs/promises');
const path = require('node:path');
const {
  renameWithTransientRetry,
} = require('./transient-rename.cjs');

const rename = fsPromises.rename.bind(fsPromises);
const projectRoot = path.resolve(__dirname, '..');
const npmCache = process.env.npm_config_cache;

if (!npmCache || !path.isAbsolute(npmCache)) {
  process.stderr.write('ERR_PACK_ELECTRON: npm cache 必须是绝对路径\n');
  process.exit(72);
}

const installResult = spawnSync(
  process.execPath,
  [
    path.join(projectRoot, 'node_modules', 'electron', 'install.js'),
  ],
  {
    cwd: projectRoot,
    env: {
      ...process.env,
      electron_config_cache: path.join(npmCache, 'electron'),
    },
    stdio: 'inherit',
  },
);

if (installResult.status !== 0) {
  process.exit(installResult.status ?? 1);
}

fsPromises.rename = (source, destination) => renameWithTransientRetry({
  source,
  destination,
  rename,
  onRetry: ({
    attempt,
    delayMs,
  }) => {
    process.stderr.write(
      `[pack:dir] 瞬态 EPERM，第 ${attempt} 次等待 ${delayMs} ms 后重试\n`,
    );
  },
});

process.argv.push(
  '--config',
  'electron\\builder.json',
  '--win',
  '--dir',
);

require('electron-builder/cli.js');
