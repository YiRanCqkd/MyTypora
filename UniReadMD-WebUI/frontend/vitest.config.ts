import { defineConfig } from 'vitest/config';
import vue from '@vitejs/plugin-vue';

const coverageFloor = 90;
const criticalCoverageFiles = (process.env.VITEST_CRITICAL_COVERAGE_FILES || '')
  .split(';')
  .map((file) => file.trim())
  .filter(Boolean);

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
    environment: 'jsdom',
    globals: true,
    include: ['src/**/*.spec.ts'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json-summary', 'html'],
      reportsDirectory: 'coverage/core',
      include: [
        'src/services/**/*.ts',
        'src/utils/**/*.ts',
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
