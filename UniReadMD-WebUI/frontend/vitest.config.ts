import { defineConfig } from 'vitest/config';
import vue from '@vitejs/plugin-vue';

const coverageFloor = 90;
const criticalCoverageFiles = [
  'src/features/markdown/safe-html.ts',
  'src/features/workspace/document-integrity.ts',
  ...(process.env.VITEST_CRITICAL_COVERAGE_FILES || '')
    .split(';')
    .map((file) => file.trim())
    .filter(Boolean),
];

const perFileThresholds = Object.fromEntries(
  criticalCoverageFiles.map((file) => [
    file,
    {
      statements: coverageFloor,
      branches: coverageFloor,
      functions: coverageFloor,
      lines: coverageFloor,
    },
  ]),
);

export default defineConfig({
  plugins: [
    vue(),
  ],
  test: {
    // 限制并发，避免 Windows 上多个工作台和覆盖率实例争抢内存。
    maxWorkers: 2,
    environment: 'jsdom',
    globals: true,
    include: [
      'src/**/*.spec.ts',
      'tests/**/*.spec.ts',
    ],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json-summary', 'html'],
      reportsDirectory: 'coverage/core',
      include: [
        'src/services/**/*.ts',
        'src/utils/**/*.ts',
        'src/features/markdown/safe-html.ts',
        'src/features/workspace/document-integrity.ts',
      ],
      exclude: [
        'src/**/*.spec.ts',
      ],
      thresholds: {
        statements: coverageFloor,
        branches: coverageFloor,
        functions: coverageFloor,
        lines: coverageFloor,
        ...perFileThresholds,
      },
    },
  },
});
