import type MarkdownIt from 'markdown-it';
import { sanitizeMarkdownHtml } from './safe-html';
import { getCachedDocumentWords, setCachedDocumentWords } from './document-text';

export function createPlainRenderCache(markdown: MarkdownIt, minimumLength = 1024 * 1024) {
  const cache = new Map<string, { html: string; words: number }>();
  let cacheSize = 0;
  return (content: string): string | null => {
    // 仅对不含特殊语法、链接和 HTML 的正文加速，其他文档保留完整解析与清洗。
    if (content.length < minimumLength || /[\\`*_<>[\]{}|$~=:&\t\r]/.test(content)
      || /^ {0,3}(?:[-+>]\s|\d+[.)]\s|[-=]{2,}\s*$)|^ {4}/m.test(content)) return null;
    const lines = content.split('\n');
    const blocks: Array<{ text: string; start: number }> = [];
    let block: string[] = [];
    let start = 0;
    const flush = () => {
      if (block.length) blocks.push({ text: block.join('\n'), start });
      block = [];
    };
    for (let index = 0; index < lines.length; index += 1) {
      if (!lines[index].trim()) { flush(); continue; }
      if (!block.length) start = index;
      block.push(lines[index]);
    }
    flush();
    // 一个缓存块最多一个标题，确保重复标题的全局编号与标准解析完全一致。
    if (blocks.some(({ text }) => (text.match(/^ {0,3}#{1,6}\s/gm)?.length || 0) > 1)) return null;
    let words = 0;
    const ids = new Set<string>();
    const html = blocks.map(({ text, start: offset }) => {
      let item = cache.get(text);
      if (!item) {
        item = { html: sanitizeMarkdownHtml(markdown.render(text)),
          words: getCachedDocumentWords(text) || 0 };
        cache.set(text, item);
        cacheSize += text.length + item.html.length;
        while (cache.size > 256 || cacheSize > 4 * 1024 * 1024) {
          const oldest = cache.keys().next().value!;
          const entry = cache.get(oldest)!;
          cacheSize -= oldest.length + entry.html.length;
          cache.delete(oldest);
        }
      }
      words += item.words;
      return item.html.replace(/data-source-(start|end)="(\d+)"/g,
        (_match, kind, line) => `data-source-${kind}="${Number(line) + offset}"`)
        .replace(/id="(md-[^"]+)"/g, (_match, base) => {
          let id = base;
          let suffix = 1;
          while (ids.has(id)) id = `${base}-${suffix++}`;
          ids.add(id);
          return `id="${id}"`;
        });
    }).join('');
    setCachedDocumentWords(content, words);
    return html;
  };
}
