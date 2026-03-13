import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';
import { fileURLToPath, URL } from 'node:url';

export default defineConfig(({ command }) => {
  return {
    // 生产打包用于 file:// 加载，需要相对路径；开发态维持根路径。
    base: command === 'serve' ? '/' : './',
    plugins: [
      vue(),
    ],
    resolve: {
      alias: {
        'markdown-it-mathjax3': fileURLToPath(
          new URL('./src/vendor/markdown-it-mathjax3.ts', import.meta.url),
        ),
      },
    },
    build: {
      modulePreload: false,
      chunkSizeWarningLimit: 900,
      rollupOptions: {
        output: {
          manualChunks(id) {
            if (!id.includes('node_modules')) {
              return;
            }

            if (id.includes('@codemirror')) {
              return 'vendor-codemirror';
            }

            if (id.includes('mermaid')) {
              return 'vendor-mermaid';
            }

            if (id.includes('markdown-it')) {
              return 'vendor-markdown';
            }

            if (id.includes('pinia') || id.includes('vue')) {
              return 'vendor-vue';
            }
          },
        },
      },
    },
    server: {
      host: '127.0.0.1',
      port: 5173,
      strictPort: true,
    },
  };
});
