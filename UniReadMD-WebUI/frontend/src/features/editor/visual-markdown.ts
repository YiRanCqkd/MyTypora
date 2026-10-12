import { markdownLanguage } from '@codemirror/lang-markdown';
import { sanitizeMarkdownHtml } from '../markdown/safe-html';

export interface VisualBlock {
  prefix: string;
  kind: 'paragraph' | 'heading' | 'list' | 'quote';
}

// 仅转换能够无损保留语义的简单块；图片、数学、HTML 和嵌套列表沿用源码入口。
export function describeVisualBlock(source: string): VisualBlock | null {
  let unsupported = false;
  markdownLanguage.parser.parse(source).iterate({ enter(node) {
    if (['Image', 'HTMLTag', 'HTMLBlock', 'FencedCode', 'CodeBlock', 'Table']
      .includes(node.name)) unsupported = true;
  } });
  if (unsupported || /\$[^$\n]+\$|\[\^[^\]]+\]|^\s*[-+*]\s+\[[ xX]\]/m.test(source)) {
    return null;
  }
  const heading = /^(#{1,6})[ \t]+/.exec(source);
  if (heading) return { prefix: heading[0], kind: 'heading' };
  const list = /^( *)(?:[-+*]|\d+[.)])[ \t]+/.exec(source);
  if (list) {
    if (source.split('\n').slice(1).some((line) => /^\s*(?:[-+*]|\d+[.)])\s/.test(line))) {
      return null;
    }
    return { prefix: list[0], kind: 'list' };
  }
  const quote = /^( *>[ \t]?)/.exec(source);
  if (quote) {
    if (/^ *>\s*(?:>|[-+*]\s|\d+[.)]\s|#{1,6}\s)/m.test(source)) return null;
    return { prefix: quote[0], kind: 'quote' };
  }
  if (/\n(?:={2,}|-{2,})\s*$/.test(source) || /^(?:---|\[\[?toc\]?\])/i.test(source)) return null;
  return { prefix: '', kind: 'paragraph' };
}

function escapeText(text: string) {
  return text.replace(/\u00a0/g, ' ').replace(/([\\`*_[\]<>~^=])/g, '\\$1')
    .replace(/^([ \t]*)(#{1,6}\s|>\s?|[-+]\s)/gm, '$1\\$2')
    .replace(/^([ \t]*)(\d+)([.)])(\s)/gm, '$1$2\\$3$4');
}

export function visualHtmlToMarkdown(html: string): string {
  const host = document.createElement('div');
  host.innerHTML = sanitizeMarkdownHtml(html, { trustedAttributes: false });
  function wrap(text: string, marker: string) {
    const [, before, value, after] = /^(\s*)([\s\S]*?)(\s*)$/.exec(text)!;
    return before + (value ? marker + value + marker : '') + after;
  }
  function walk(node: Node): string {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent || '';
      // markdown-it 输出 <br> 后会追加排版换行，写回时不能重复计算。
      const afterBreak = node.previousSibling instanceof HTMLElement
        && node.previousSibling.tagName === 'BR';
      return escapeText(afterBreak ? text.replace(/^\r?\n/, '') : text);
    }
    if (!(node instanceof HTMLElement)) return '';
    const text = [...node.childNodes].map(walk).join('');
    const tag = node.tagName.toLowerCase();
    if (tag === 'br') return '\n';
    if (tag === 'strong' || tag === 'b') return wrap(text, '**');
    if (tag === 'em' || tag === 'i') return wrap(text, '*');
    if (tag === 's' || tag === 'del' || tag === 'strike') return wrap(text, '~~');
    if (tag === 'mark') return wrap(text, '==');
    if (tag === 'sub' || tag === 'sup') {
      const mark = tag === 'sub' ? '~' : '^';
      return `${mark}${text.replace(/ /g, '\\ ')}${mark}`;
    }
    if (tag === 'code') {
      const value = node.textContent || '';
      const fence = '`'.repeat(Math.max(1, ...[...value.matchAll(/`+/g)]
        .map((match) => match[0].length + 1)));
      const padding = /^`|`$|^ .* $/.test(value) ? ' ' : '';
      return `${fence}${padding}${value}${padding}${fence}`;
    }
    if (tag === 'a' && node.hasAttribute('href')) {
      const href = node.getAttribute('href')!.replace(/ /g, '%20').replace(/>/g, '%3E');
      const title = node.getAttribute('title');
      return `[${text}](<${href}>${title ? ` "${title.replace(/"/g, '&quot;')}"` : ''})`;
    }
    if (tag === 'div' || tag === 'p') return `\n\n${text}\n\n`;
    return text;
  }
  return [...host.childNodes].map(walk).join('').replace(/\n{3,}/g, '\n\n')
    .replace(/^\n+|\n+$/g, '');
}

export function applyVisualPrefix(text: string, block: VisualBlock) {
  if (block.kind === 'quote') return text.split('\n').map((line) => block.prefix + line).join('\n');
  if (block.kind === 'list') {
    return block.prefix + text.replace(/\n/g, '\n' + ' '.repeat(block.prefix.length));
  }
  return block.prefix + text;
}

export function nextVisualPrefix(block: VisualBlock) {
  if (block.kind === 'heading' || block.kind === 'paragraph') return '';
  if (block.kind === 'quote') return block.prefix;
  return block.prefix.replace(/\d+(?=[.)])/, (number) => String(Number(number) + 1));
}
