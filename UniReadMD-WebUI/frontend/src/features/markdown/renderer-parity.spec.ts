import { describe, expect, it } from 'vitest';
import { createMarkdownRenderer } from './renderer';
import { createWorkbenchRuntime } from '../editor/workbench-runtime';
import { sanitizeMarkdownHtml } from './safe-html';
import { highlightCode } from './code-highlight';
import { createMarkdownSourceMapping } from './source-mapping';

describe('Markdown 预览标准能力', () => {
  it('围栏与缩进代码的背景留白包含完整源码映射', () => {
    const renderer = createMarkdownRenderer(createWorkbenchRuntime());
    createMarkdownSourceMapping(renderer);
    for (const source of ['```js\nconst a = 1;\n```', '    code']) {
      const host = document.createElement('div');
      host.innerHTML = sanitizeMarkdownHtml(renderer.markdown.render(source));
      expect(host.querySelector('pre')?.dataset.sourceStart).toBe('1');
      expect(host.querySelector('pre')?.dataset.sourceToken)
        .toBe(source.startsWith('```') ? 'fence' : 'code_block');
    }
  });

  it('标题生成稳定且不重复的锚点，目录指向可找到的标题', () => {
    const { markdown } = createMarkdownRenderer(createWorkbenchRuntime());
    const host = document.createElement('div');
    host.innerHTML = sanitizeMarkdownHtml(markdown.render('[[toc]]\n\n# 中文标题\n\n# 中文标题'));
    const headings = [...host.querySelectorAll('h1')];
    expect(headings.map((heading) => heading.id)).toEqual(['md-中文标题', 'md-中文标题-1']);
    const links = [...host.querySelectorAll('a')];
    expect(links).toHaveLength(2);
    for (const link of links) {
      const id = decodeURIComponent(link.getAttribute('href')!.slice(1));
      expect(headings.some((heading) => heading.id === id)).toBe(true);
      expect(link.hasAttribute('target')).toBe(false);
    }
  });
  it('已知语言高亮并转义代码，未知语言降级为普通文本', () => {
    expect(highlightCode('const value = "<script>";', 'js')).toContain('hljs-keyword');
    expect(highlightCode('const value = "<script>";', 'js')).not.toContain('<script>');
    expect(highlightCode('<script>', 'unknown')).toBe('');
  });
});
