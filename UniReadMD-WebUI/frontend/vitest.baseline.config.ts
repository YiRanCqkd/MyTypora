import vue from '@vitejs/plugin-vue';
import { defineConfig } from 'vitest/config';

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
      reportsDirectory: 'coverage/baseline',
      include: [
        'src/**/*.ts',
        'src/**/*.vue',
      ],
      exclude: [
        'src/**/*.spec.ts',
        'src/**/*.d.ts',
        'src/vendor/**',
      ],
    },
  },
});
