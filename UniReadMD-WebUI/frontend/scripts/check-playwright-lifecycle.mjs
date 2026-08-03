import { spawn, spawnSync } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const EXIT_TIMEOUT_MS = 45_000;
const TASK_COUNT = 6;
const SERVER_HOST = '127.0.0.1';
const SERVER_PORT = 4173;

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const frontendDirectory = path.resolve(scriptDirectory, '..');
const playwrightCli = path.join(
  frontendDirectory,
  'node_modules',
  '@playwright',
  'test',
  'cli.js',
);
const systemDirectory = path.join(process.env.SystemRoot ?? 'C:\\Windows', 'System32');
const normalizedSystemDirectory = path.normalize(systemDirectory).toLowerCase();
const pathWithoutSystemDirectory = (process.env.PATH ?? '')
  .split(path.delimiter)
  .filter((entry) => path.normalize(entry).toLowerCase() !== normalizedSystemDirectory)
  .join(path.delimiter);

const child = spawn(
  process.execPath,
  [
    playwrightCli,
    'test',
    'task2-characterization.spec.ts',
    '--reporter=json',
  ],
  {
    cwd: frontendDirectory,
    env: {
      ...process.env,
      PATH: pathWithoutSystemDirectory,
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  },
);

let stdout = '';
let stderr = '';
child.stdout.setEncoding('utf8');
child.stderr.setEncoding('utf8');
child.stdout.on('data', (chunk) => {
  stdout += chunk;
});
child.stderr.on('data', (chunk) => {
  stderr += chunk;
});

const exitCode = await new Promise((resolve, reject) => {
  const timer = setTimeout(() => {
    const taskkill = path.join(systemDirectory, 'taskkill.exe');
    if (child.pid) {
      spawnSync(taskkill, ['/pid', String(child.pid), '/T', '/F'], {
        stdio: 'ignore',
      });
    }
    reject(new Error(
      `Task 2 断言运行后 Playwright 未在 ${EXIT_TIMEOUT_MS}ms 内自然退出。`,
    ));
  }, EXIT_TIMEOUT_MS);

  child.once('error', (error) => {
    clearTimeout(timer);
    reject(error);
  });
  child.once('close', (code) => {
    clearTimeout(timer);
    resolve(code);
  });
});

if (exitCode !== 0) {
  throw new Error(`Playwright 退出码为 ${exitCode}。\n${stderr}`);
}

const report = JSON.parse(stdout);
const stats = report.stats;
if (
  !stats
  || stats.expected !== TASK_COUNT
  || stats.unexpected !== 0
  || stats.skipped !== 0
  || stats.flaky !== 0
) {
  throw new Error(`Playwright JSON 终态不符合预期：${JSON.stringify(stats)}`);
}

await new Promise((resolve, reject) => {
  const server = net.createServer();
  server.once('error', reject);
  server.listen(SERVER_PORT, SERVER_HOST, () => {
    server.close((error) => {
      if (error) {
        reject(error);
        return;
      }
      resolve();
    });
  });
});

process.stdout.write(`${JSON.stringify({
  exitCode,
  expected: stats.expected,
  flaky: stats.flaky,
  skipped: stats.skipped,
  unexpected: stats.unexpected,
  serverPortReleased: true,
})}\n`);
