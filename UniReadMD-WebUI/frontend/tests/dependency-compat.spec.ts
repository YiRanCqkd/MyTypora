import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import MarkdownIt from 'markdown-it';
import mermaid from 'mermaid';
import { describe, expect, it } from 'vitest';
import markdownItMathJax3 from '../src/vendor/markdown-it-mathjax3';

const fixture = [
  '# 依赖兼容性',
  '',
  '固定公式：$E = mc^2$',
  '',
  '```mermaid',
  'flowchart LR',
  '  A[Markdown] --> B[DOM]',
  '```',
].join('\n');

const directDependencyPins = [
  '"dompurify": "3.4.12"',
  '"markdown-it": "14.3.0"',
  '"mermaid": "11.16.0"',
  '"postcss": "8.5.23"',
] as const;

describe('升级依赖兼容性', () => {
  it(
    'Given 固定 Markdown/KaTeX/Mermaid fixture, when 渲染并净化, then DOM 完整且无错误',
    async () => {
      const lockfile = readFileSync(resolve(process.cwd(), 'package-lock.json'), 'utf8');
      const directDependencyHeader = lockfile.slice(0, 4_000);
      for (const dependencyPin of directDependencyPins) {
        expect(directDependencyHeader).toContain(dependencyPin);
      }

      Object.defineProperties(SVGElement.prototype, {
        getBBox: {
          configurable: true,
          value: () => ({ height: 24, width: 120, x: 0, y: 0 }),
        },
        getComputedTextLength: {
          configurable: true,
          value: () => 120,
        },
      });

      const markdown = new MarkdownIt({ html: true });
      markdown.use(markdownItMathJax3);
      const host = document.createElement('main');
      host.innerHTML = markdown.render(fixture);

      const diagramCode = host.querySelector<HTMLElement>('code.language-mermaid');
      const diagramContainer = document.createElement('section');
      const capturedErrors: unknown[] = [];
      const originalConsoleError = console.error;
      console.error = (...messages: unknown[]) => {
        capturedErrors.push(messages);
      };

      try {
        mermaid.initialize({ securityLevel: 'strict', startOnLoad: false });
        const rendered = await mermaid.render(
          'dependency-compat-mermaid',
          diagramCode?.textContent ?? '',
        );
        diagramContainer.innerHTML = rendered.svg;
        diagramCode?.parentElement?.replaceWith(diagramContainer);
      } finally {
        console.error = originalConsoleError;
      }

      const { default: DOMPurify } = await import('dompurify');
      host.innerHTML = DOMPurify.sanitize(host.innerHTML);

      expect(host.querySelector('h1')?.textContent).toBe('依赖兼容性');
      expect(host.querySelector('.katex')).not.toBeNull();
      expect(host.querySelector('section svg')).not.toBeNull();
      expect(capturedErrors).toEqual([]);
    },
  );
});
