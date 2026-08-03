import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { spawnSync } from 'node:child_process';

import { findRawColorViolations } from '../check-design-tokens.mjs';
import { E2eBaselineError, evaluateE2eBaseline } from '../check-e2e-baseline.mjs';
import { countScannedLines, extractCountableSource } from '../check-source-size.mjs';

function createReport(status, title = '已知失败') {
  return {
    stats: {
      flaky: 0,
      skipped: 0,
    },
    suites: [
      {
        title: 'sample.spec.ts',
        file: 'sample.spec.ts',
        specs: [
          {
            file: 'sample.spec.ts',
            line: 3,
            column: 1,
            title,
            tests: [
              {
                projectName: 'chromium',
                results: [{ status }],
              },
            ],
          },
        ],
      },
    ],
  };
}

function createFixedNodeEnvironment() {
  const environment = {};
  for (const [key, value] of Object.entries(process.env)) {
    if (key.toUpperCase() === 'PATH' || value === undefined) {
      continue;
    }
    environment[key] = value;
  }
  environment.Path = 'D:\\Program Files\\nodejs';
  const windowsPowerShellModules = path.join(
    environment.SystemRoot || 'C:\\Windows',
    'System32',
    'WindowsPowerShell',
    'v1.0',
    'Modules',
  );
  const existingModulePaths = String(environment.PSModulePath || '')
    .split(path.delimiter)
    .filter((entry) => (
      entry
      && entry.toLowerCase() !== windowsPowerShellModules.toLowerCase()
    ));
  environment.PSModulePath = [
    windowsPowerShellModules,
    ...existingModulePaths,
  ].join(path.delimiter);
  return environment;
}

test(
  'Given comments and blank lines, when scanning source, then only physical code lines count',
  () => {
  const source = [
    '// 纯注释',
    '',
    'const value = 1; // 行尾注释',
    '/* 多行',
    ' * 注释 */',
    'value.toString();',
  ].join('\n');

  assert.equal(countScannedLines(source), 2);
  },
);

test('Given a Vue SFC, when extracting source, then template and style are excluded', () => {
  const source = [
    '<template><main>不计数</main></template>',
    '<script setup lang=\'ts\'>',
    'const value = 1;',
    '</script>',
    '<style>.sample { color: red; }</style>',
  ].join('\n');

  assert.equal(countScannedLines(extractCountableSource('Sample.vue', source)), 1);
});

test('Given token declarations and raw usage, when checking colors, then only usage fails', () => {
  const source = [
    ':root {',
    '  --surface: #ffffff;',
    '}',
    '.card { color: #ffffff; }',
  ].join('\n');

  assert.deepEqual(findRawColorViolations(source), [
    {
      file: '<memory>',
      line: 4,
      value: '#ffffff',
    },
  ]);
});

test('Given a known failure, when evaluating the report, then the baseline passes', () => {
  const id = '[chromium] › sample.spec.ts:3:1 › 已知失败';
  const result = evaluateE2eBaseline(createReport('failed'), {
    minimumPassed: 0,
    knownFailureIds: [id],
  });

  assert.deepEqual(result, {
    passed: 0,
    failed: 1,
    skipped: 0,
    knownFailures: 1,
  });
});

test('Given a new failure, when evaluating the report, then a stable error is returned', () => {
  assert.throws(
    () => evaluateE2eBaseline(createReport('failed', '新增失败'), {
      minimumPassed: 0,
      knownFailureIds: [],
    }),
    (error) => error instanceof E2eBaselineError && error.code === 'ERR_E2E_NEW_FAILURE',
  );
});

test(
  'Given the fixed helper, when invoking Node, then the resolver-selected version is used',
  () => {
  const powershell = path.join(
    process.env.SystemRoot || 'C:\\Windows',
    'System32',
    'WindowsPowerShell',
    'v1.0',
    'powershell.exe',
  );
  const helper = path.resolve('scripts', 'invoke-node-tool.ps1');
  const result = spawnSync(
    powershell,
    [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      helper,
      'node',
      '--version',
    ],
    {
      cwd: path.resolve('.'),
      encoding: 'utf8',
      env: createFixedNodeEnvironment(),
      timeout: 30_000,
    },
  );

  assert.equal(result.status, 0, result.stderr);
  assert.equal(result.stdout.trim(), 'v24.8.0');
  },
);

test('Given a child exit code, when invoking Node, then the helper passes it through', () => {
  const powershell = path.join(
    process.env.SystemRoot || 'C:\\Windows',
    'System32',
    'WindowsPowerShell',
    'v1.0',
    'powershell.exe',
  );
  const helper = path.resolve('scripts', 'invoke-node-tool.ps1');
  const result = spawnSync(
    powershell,
    [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      helper,
      'node',
      '-e',
      'process.exit(23)',
    ],
    {
      cwd: path.resolve('.'),
      encoding: 'utf8',
      env: createFixedNodeEnvironment(),
      timeout: 30_000,
    },
  );

  assert.equal(result.status, 23, result.stderr);
});

test('Given an over-limit file, when running source-size CLI, then it fails visibly', () => {
  const directory = fs.mkdtempSync('D:\\DevTools\\unireadmd-source-size-');
  const fixture = path.join(directory, 'oversized.ts');
  fs.writeFileSync(fixture, 'const one = 1;\nconst two = 2;\n', 'utf8');
  const script = path.resolve('scripts', 'check-source-size.mjs');

  try {
    const result = spawnSync(process.execPath, [script, '--limit', '1', fixture], {
      encoding: 'utf8',
    });
    assert.equal(result.status, 1);
    assert.match(result.stdout, /ERR_SOURCE_SIZE/);
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});
