import { markdownLanguage } from '@codemirror/lang-markdown';
import type MarkdownIt from 'markdown-it';
import { sanitizeMarkdownHtml } from '../markdown/safe-html';

export type ImageSource = { from: number; to: number; html: string };

export function findPreviewImageSource(source: string, image: HTMLImageElement,
  host: HTMLElement, markdown: MarkdownIt, resourceBaseUrl?: string): ImageSource | null {
  if (!host.contains(image)) return null;
  const block = image.closest<HTMLElement>('[data-source-start][data-source-end]') || host;
  const lines = source.split('\n');
  const start = Number(block.dataset.sourceStart) || 1;
  const end = Number(block.dataset.sourceEnd) || lines.length;
  const from = lines.slice(0, start - 1).reduce((sum, line) => sum + line.length + 1, 0);
  const to = from + lines.slice(start - 1, end).join('\n').length;
  const env = {};
  markdown.parse(source, env);
  const candidates: ImageSource[] = [];
  function add(left: number, right: number, html = false) {
    if (left < from || right > to) return;
    const fragment = source.slice(left, right);
    const rendered = html ? fragment : markdown.renderInline(fragment, env);
    const safe = sanitizeMarkdownHtml(rendered, { trustedAttributes: false });
    const template = document.createElement('template');
    template.innerHTML = safe;
    const original = template.content.querySelector('img');
    if (original) candidates.push({ from: left, to: right, html: original.outerHTML });
  }
  markdownLanguage.parser.parse(source).iterate({ enter(node) {
    if (node.to < from || node.from > to) return false;
    if (node.name === 'Image') { add(node.from, node.to); return false; }
    if (node.name === 'HTMLTag' || node.name === 'HTMLBlock') {
      // 跳过注释和脚本里的伪标签，带引号的属性值可以含大于号。
      const fragment = source.slice(node.from, node.to).replace(
        /<!--[\s\S]*?-->|<(script|style)\b[^>]*>[\s\S]*?<\/\1\s*>/gi,
        (text) => ' '.repeat(text.length),
      );
      for (const match of fragment.matchAll(/<img\b(?:[^>"']|"[^"]*"|'[^']*')*>/gi)) {
        add(node.from + match.index, node.from + match.index + match[0].length, true);
      }
      return false;
    }
  } });
  const renderedImages = [...block.querySelectorAll('img')];
  if (renderedImages.length !== candidates.length) return null;
  // 校验所有对应关系，遇到插件重排或过期 DOM 时宁可拒绝，也不写错图片。
  for (let index = 0; index < candidates.length; index += 1) {
    const template = document.createElement('template');
    template.innerHTML = sanitizeMarkdownHtml(candidates[index].html, { resourceBaseUrl });
    const expected = template.content.querySelector('img');
    const actual = renderedImages[index];
    if (expected?.getAttribute('src') !== actual.getAttribute('src')
      || expected?.alt !== actual.alt) return null;
  }
  return candidates[renderedImages.indexOf(image)] || null;
}

export function resizeImageSource(selection: ImageSource, naturalWidth: number,
  percent: number | null): string {
  if (percent !== null && (!Number.isFinite(percent) || percent < 1 || percent > 400
    || !Number.isFinite(naturalWidth) || naturalWidth <= 0)) {
    throw new Error('请输入 1%～400% 的比例，并等待原图加载完成。');
  }
  const template = document.createElement('template');
  template.innerHTML = selection.html;
  const image = template.content.querySelector('img');
  if (!image) throw new Error('无法识别图片源码，请重新右键选择图片。');
  image.removeAttribute('width');
  image.removeAttribute('height');
  for (const property of ['width', 'height', 'min-width', 'min-height', 'max-width',
    'max-height', 'zoom', 'aspect-ratio']) image.style.removeProperty(property);
  if (percent !== null) {
    const width = Math.max(1, Math.round(naturalWidth * percent / 100));
    image.setAttribute('width', String(width));
    image.style.width = `${width}px`;
    image.style.height = 'auto';
    image.style.maxWidth = '100%';
  }
  if (!image.style.length) image.removeAttribute('style');
  return image.outerHTML;
}

export function buildImageResizeSource(source: string, selection: ImageSource,
  naturalWidth: number, percent: number | null) {
  const image = resizeImageSource(selection, naturalWidth, percent);
  const lineStart = source.lastIndexOf('\n', Math.max(0, selection.from - 1)) + 1;
  const before = source.slice(lineStart, selection.from);
  const after = source.slice(selection.to);
  // 独占一行的 HTML 图片会开启 HTML 块，补空行以免下一行 Markdown 被吞掉。
  if (source.slice(selection.from, selection.to).startsWith('![') && !before.trim()
    && /^\r?\n(?=[^\r\n]*\S)/.test(after)) {
    return image + (after.startsWith('\r\n') ? '\r\n' : '\n');
  }
  return image;
}
