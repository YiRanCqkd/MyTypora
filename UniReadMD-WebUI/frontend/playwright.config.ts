import path from 'node:path';
import { defineConfig, devices } from '@playwright/test';

// Playwright 1.55 在 Windows 通过 PATH 查找 taskkill 结束 webServer 进程树。
// 隔离环境可能移除 System32；配置加载时补回它，避免 teardown 永久等待 close。
if (process.platform === 'win32' && process.env.SystemRoot) {
  const systemDirectory = path.join(process.env.SystemRoot, 'System32');
  process.env.PATH = [systemDirectory, process.env.PATH]
    .filter(Boolean)
    .join(path.delimiter);
}

export default defineConfig({
  testDir: './e2e',
  fullyParallel: true,
  // Windows 页面进程初始化较慢；断言超时独立保持 10 秒。
  timeout: 60_000,
  expect: {
    timeout: 10_000,
  },
  use: {
    baseURL: 'http://127.0.0.1:4173',
    trace: 'retain-on-failure',
  },
  webServer: {
    command: 'node ./node_modules/vite/bin/vite.js build && '
      + 'node ./node_modules/vite/bin/vite.js preview --host 127.0.0.1 --port 4173 --strictPort',
    url: 'http://127.0.0.1:4173',
    reuseExistingServer: false,
    timeout: 180_000,
  },
  projects: [
    {
      name: 'chromium',
      use: {
        ...devices['Desktop Chrome'],
      },
    },
  ],
});
