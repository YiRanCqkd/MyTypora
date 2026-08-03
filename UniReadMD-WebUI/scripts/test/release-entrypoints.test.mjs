import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { spawnSync } from 'node:child_process';

const repositoryRoot = path.resolve('..');
const helper = path.resolve('scripts', 'invoke-node-tool.ps1');
const powershell = path.join(
  process.env.SystemRoot || 'C:\\Windows',
  'System32',
  'WindowsPowerShell',
  'v1.0',
  'powershell.exe',
);

function createFixedEnvironment(extra = {}) {
  const environment = {};
  for (const [key, value] of Object.entries(process.env)) {
    if (key.toUpperCase() !== 'PATH' && value !== undefined) {
      environment[key] = value;
    }
  }
  environment.Path = 'D:\\Program Files\\nodejs';
  const windowsPowerShellModules = path.join(
    environment.SystemRoot || 'C:\\Windows',
    'System32',
    'WindowsPowerShell',
    'v1.0',
    'Modules',
  );
  const modulePaths = String(environment.PSModulePath || '')
    .split(path.delimiter)
    .filter((entry) => (
      entry && entry.toLowerCase() !== windowsPowerShellModules.toLowerCase()
    ));
  environment.PSModulePath = [windowsPowerShellModules, ...modulePaths].join(path.delimiter);
  return { ...environment, ...extra };
}

function invokeCommand(command, extraEnvironment = {}) {
  return spawnSync(
    powershell,
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', command],
    {
      cwd: path.resolve('.'),
      encoding: 'utf8',
      env: createFixedEnvironment(extraEnvironment),
      timeout: 60_000,
    },
  );
}

function invokeFile(arguments_) {
  return spawnSync(
    powershell,
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', helper, ...arguments_],
    {
      cwd: path.resolve('.'),
      encoding: 'utf8',
      env: createFixedEnvironment(),
      timeout: 60_000,
    },
  );
}

test('helper 拒绝非法工具', () => {
  const result = invokeFile(['invalid', '--version']);

  assert.equal(result.status, 64);
  assert.match(result.stderr, /ERR_INVOKE_NODE_TOOL/);
});

test('显式参数协议把 npm 脚本分隔符可靠传给子命令', () => {
  const fixture = fs.mkdtempSync('D:\\DevTools\\unireadmd-helper-');
  const packageJson = {
    name: 'unireadmd-helper-probe',
    private: true,
    scripts: {
      probe: 'node probe.mjs',
    },
  };
  fs.writeFileSync(
    path.join(fixture, 'package.json'),
    `${JSON.stringify(packageJson, null, 2)}\r\n`,
    'utf8',
  );
  fs.writeFileSync(
    path.join(fixture, 'probe.mjs'),
    [
      'process.stdout.write(JSON.stringify({',
      '  arguments: process.argv.slice(2),',
      '  path: process.env.Path || process.env.PATH,',
      '}));',
      '',
    ].join('\r\n'),
    'utf8',
  );

  try {
    const command = [
      '& $env:HELPER npm -LiteralArguments @(',
      "'--cache', 'D:\\DevTools\\npm-cache', '--prefix', $env:FIXTURE,",
      "'run', 'probe', '--', '--reporter=json')",
    ].join(' ');
    const version = invokeCommand(
      "& $env:HELPER node -LiteralArguments @('--version')",
      { HELPER: helper },
    );
    const result = invokeCommand(command, { FIXTURE: fixture, HELPER: helper });

    assert.equal(version.status, 0, version.stderr);
    assert.equal(version.stdout.trim(), 'v24.8.0');
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout, /"arguments":\["--reporter=json"\]/u);
    assert.match(result.stdout, /Windows\\\\System32/iu);
  } finally {
    fs.rmSync(fixture, { recursive: true, force: true });
  }
});

test('混用兼容参数与显式参数协议时返回稳定错误', () => {
  const result = invokeFile([
    'node',
    '--version',
    '-LiteralArguments',
    '--version',
  ]);

  assert.equal(result.status, 64);
  assert.match(result.stderr, /ERR_INVOKE_NODE_TOOL_ARGUMENTS/);
});

test('CI 精确固定 Task 1 的 Node 补丁版本', () => {
  const workflow = fs.readFileSync(
    path.join(repositoryRoot, '.github', 'workflows', 'frontend-quality.yml'),
    'utf8',
  );

  assert.match(workflow, /node-version:\s*'24\.8\.0'/u);
  assert.doesNotMatch(workflow, /node-version:\s*'24'/u);
});

test('Playwright 标准入口直接启动可归属的 Vite 服务', () => {
  const config = fs.readFileSync(
    path.join(repositoryRoot, 'UniReadMD-WebUI', 'frontend', 'playwright.config.ts'),
    'utf8',
  );

  assert.match(config, /command:\s*'node \.\/node_modules\/vite\/bin\/vite\.js/u);
  assert.match(config, /reuseExistingServer:\s*false/u);
  assert.doesNotMatch(config, /command:\s*'npm run dev/u);
});
