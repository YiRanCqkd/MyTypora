import { describe, expect, it } from 'vitest';
import { createMarkdownRenderer } from './renderer';
import { createMarkdownSourceMapping } from './source-mapping';
import { createWorkbenchRuntime } from '../editor/workbench-runtime';
import { sanitizeMarkdownHtml } from './safe-html';
import { createPlainRenderCache } from './plain-render-cache';
import { getCachedDocumentWords } from './document-text';

describe('大文档纯正文缓存', () => {
  it('缓存命中、重复标题、空行和多行正文与完整解析输出一致，源行保持准确', () => {
    const renderer = createMarkdownRenderer(createWorkbenchRuntime());
    createMarkdownSourceMapping(renderer);
    const render = createPlainRenderCache(renderer.markdown, 0);
    const content = '# 标题\n\n正文 中文\n第二行\n\n# 标题\n\n正文 中文\n第二行\n';
    const expected = sanitizeMarkdownHtml(renderer.markdown.render(content));
    expect(render(content)).toBe(expected);
    expect(render(content)).toBe(expected);
    expect(getCachedDocumentWords(content)).toBeGreaterThan(0);
    const edited = content + '\n尾部新增';
    expect(render(edited)).toBe(sanitizeMarkdownHtml(renderer.markdown.render(edited)));
  });
  it('特殊语法、HTML、列表、链接、缩进和多标题块回到完整解析', () => {
    const renderer = createMarkdownRenderer(createWorkbenchRuntime());
    const render = createPlainRenderCache(renderer.markdown, 0);
    for (const text of ['**粗体**', '<script>1</script>', '- 列表', 'https://example.com',
      '    缩进', '正文\n---', '# 甲\n# 乙', '$x$', '&copy;']) {
      expect(render(text)).toBeNull();
    }
  });
});
