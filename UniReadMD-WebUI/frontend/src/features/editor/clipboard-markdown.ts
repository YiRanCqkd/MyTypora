import MarkdownIt from 'markdown-it';
import DOMPurify from 'dompurify';
import { serializeMarkdownTable } from '../../utils/preview-table';

const renderer = new MarkdownIt({ html: false, breaks: true });
const escapeText = (text: string) => text.replace(/([\\`*_[\]<>])/g, '\\$1');

function safeReference(value: string) {
  const normalized = [...value.trim()].filter((char) => char.charCodeAt(0) > 32).join('');
  if (/^(?:javascript|vbscript|file|data):/i.test(normalized)) return '';
  return value.trim().replace(/[\s()<>]/g, (char) => encodeURIComponent(char));
}

export function htmlToMarkdown(html: string) {
  if (html.length > 5 * 1024 * 1024) throw new Error('粘贴内容超过 5 MiB');
  const clean = DOMPurify.sanitize(html, {
    FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed', 'form', 'input'],
  });
  const body = new DOMParser().parseFromString(clean, 'text/html').body;
  function children(node: Node): string {
    return [...node.childNodes].map(convert).reduce((result, text) => {
      return result.endsWith('\n\n') && text.startsWith('\n\n')
        ? result + text.replace(/^\n+/, '') : result + text;
    }, '');
  }
  function convert(node: Node): string {
    if (node.nodeType === Node.TEXT_NODE) {
      return escapeText((node.textContent || '').replace(/\s+/g, ' '));
    }
    if (!(node instanceof Element)) return '';
    const tag = node.tagName.toLowerCase();
    if (tag === 'pre') {
      const code = node.querySelector('code');
      const value = (code || node).textContent || '';
      const longest = Math.max(2, ...[...value.matchAll(/`+/g)].map((item) => item[0].length));
      const fence = '`'.repeat(longest + 1);
      const language = /language-([\w+-]+)/.exec(code?.className || '')?.[1] || '';
      return `\n\n${fence}${language}\n${value.replace(/\n$/, '')}\n${fence}\n\n`;
    }
    if (tag === 'code') {
      const value = node.textContent || '';
      const longest = Math.max(0, ...[...value.matchAll(/`+/g)].map((item) => item[0].length));
      const fence = '`'.repeat(longest + 1);
      const padding = /^`|`$|^ .* $/.test(value) ? ' ' : '';
      return `${fence}${padding}${value}${padding}${fence}`;
    }
    if (tag === 'table') {
      const rows = [...node.querySelectorAll('tr')].filter((row) => row.closest('table') === node)
        .map((row) => [...row.children].filter((cell) => /^(TH|TD)$/.test(cell.tagName))
          .map((cell) => children(cell).trim().replace(/\n+/g, '<br>')));
      if (!rows.length) return '';
      const count = Math.max(...rows.map((row) => row.length));
      rows.forEach((row) => { while (row.length < count) row.push(''); });
      return `\n\n${serializeMarkdownTable({ header: rows[0], rows: rows.slice(1) })}\n\n`;
    }
    if (tag === 'ul' || tag === 'ol') {
      let index = Number(node.getAttribute('start')) || 1;
      const items = [...node.children].filter((child) => child.tagName === 'LI').map((item) => {
        const marker = tag === 'ol' ? `${index++}. ` : '- ';
        const lines = children(item).trim().split('\n');
        return marker + lines.join(`\n${' '.repeat(marker.length)}`);
      });
      return `\n\n${items.join('\n')}\n\n`;
    }
    if (tag === 'img') {
      const source = safeReference(node.getAttribute('src') || '');
      return source ? `![${escapeText(node.getAttribute('alt') || '')}](${source})` : '';
    }
    const text = children(node);
    if (tag === 'span' && node instanceof HTMLElement) {
      let formatted = text;
      if (/^(?:bold|[6-9]00)$/.test(node.style.fontWeight)) formatted = `**${formatted}**`;
      if (node.style.fontStyle === 'italic') formatted = `*${formatted}*`;
      if (node.style.textDecoration.includes('line-through')) formatted = `~~${formatted}~~`;
      return formatted;
    }
    if (tag === 'mark') return `==${text}==`;
    if (tag === 'sup') return `^${text}^`;
    if (tag === 'sub') return `~${text}~`;
    if (tag === 'a') {
      const href = safeReference(node.getAttribute('href') || '');
      return href ? `[${text}](${href})` : text;
    }
    if (tag === 'strong' || tag === 'b') return text.trim() ? `**${text.trim()}**` : text;
    if (tag === 'em' || tag === 'i') return text.trim() ? `*${text.trim()}*` : text;
    if (tag === 's' || tag === 'del') return `~~${text}~~`;
    if (/^h[1-6]$/.test(tag)) return `\n\n${'#'.repeat(Number(tag[1]))} ${text.trim()}\n\n`;
    if (tag === 'blockquote') return `\n\n${text.trim().split('\n')
      .map((line) => `> ${line}`).join('\n')}\n\n`;
    if (tag === 'br') return '\n';
    if (tag === 'hr') return '\n\n---\n\n';
    if (['p', 'div', 'section', 'article'].includes(tag)) return `\n\n${text.trim()}\n\n`;
    return text;
  }
  // 只整理块间空行，不能重写代码块内部的空白。
  return children(body).trim();
}

export function markdownClipboardHtml(markdown: string) {
  return DOMPurify.sanitize(renderer.render(markdown));
}

export async function writeMarkdownClipboard(markdown: string) {
  if (typeof ClipboardItem !== 'undefined' && navigator.clipboard.write) {
    await navigator.clipboard.write([new ClipboardItem({
      'text/plain': new Blob([markdown], { type: 'text/plain' }),
      'text/html': new Blob([markdownClipboardHtml(markdown)], { type: 'text/html' }),
    })]);
  } else await navigator.clipboard.writeText(markdown);
}

export async function readMarkdownClipboard() {
  if (navigator.clipboard.read) {
    const items = await navigator.clipboard.read();
    for (const item of items) {
      if (item.types.includes('text/html')) {
        return htmlToMarkdown(await (await item.getType('text/html')).text());
      }
      if (item.types.includes('text/plain')) {
        return await (await item.getType('text/plain')).text();
      }
    }
  }
  return navigator.clipboard.readText();
}
