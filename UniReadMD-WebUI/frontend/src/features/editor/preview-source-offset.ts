import { markdownLanguage } from '@codemirror/lang-markdown';
import { buildInlineShortcutPatch, type InlineShortcutKind } from '../../utils/markdown-shortcuts';

export function formatPreviewRange(source: string, from: number, to: number,
  kind: InlineShortcutKind) {
  const selected = source.slice(from, to);
  if (!selected.includes('\n')) return buildInlineShortcutPatch(selected, kind).insert;
  return selected.split('\n').map((line) => {
    if (!line.trim()) return line;
    const marker = /^(?:\s*(?:>\s*)*)(?:#{1,6}\s+|(?:[-+*]|\d+[.)])\s+)?/.exec(line)?.[0] || '';
    const value = line.slice(marker.length);
    return value ? marker + buildInlineShortcutPatch(value, kind).insert : line;
  }).join('\n');
}

export function sourceOffsetForRenderedText(source: string, prefix: string) {
  const hidden = new Uint8Array(source.length);
  const entities = new Map<number, { text: string; end: number }>();
  markdownLanguage.parser.parse(source).iterate({ enter(node) {
    const marker = ['HeaderMark', 'QuoteMark', 'ListMark', 'EmphasisMark', 'CodeMark',
      'CodeInfo', 'LinkMark', 'LinkTitle'].includes(node.name);
    if (marker
      || node.name === 'URL' && node.node.parent?.name === 'Link') {
      hidden.fill(1, node.from, node.to);
      if (node.node.parent?.name === 'FencedCode' && source[node.to] === '\n') {
        hidden[node.to] = 1;
      }
      if (/^(?:HeaderMark|QuoteMark|ListMark)$/.test(node.name)) {
        let end = node.to;
        while (source[end] === ' ' || source[end] === '\t') hidden[end++] = 1;
      }
    }
    if (node.name === 'Escape') hidden[node.from] = 1;
    if (node.name === 'Entity') {
      const element = document.createElement('textarea');
      element.innerHTML = source.slice(node.from, node.to);
      entities.set(node.from, { text: element.value, end: node.to });
    }
  } });
  const chars: Array<{ char: string; from: number; to: number }> = [];
  for (let index = 0; index < source.length; index += 1) {
    if (hidden[index]) continue;
    const entity = entities.get(index);
    if (entity) {
      for (const char of entity.text.split('')) {
        chars.push({ char, from: index, to: entity.end });
      }
      index = entity.end - 1;
    } else chars.push({ char: source[index], from: index, to: index + 1 });
  }
  if (!prefix) return chars[0]?.from || 0;
  let cursor = 0;
  let offset = chars[0]?.from || 0;
  for (const char of prefix.split('')) {
    // DOM 的 br 不产生文本字符，跨行软换行不能挤偏后续位置。
    while (cursor < chars.length && chars[cursor].char !== char
      && /\s/.test(chars[cursor].char)) cursor += 1;
    if (chars[cursor]?.char !== char) break;
    offset = chars[cursor++].to;
  }
  return Math.min(source.length, offset);
}

export function previewPointToSource(
  host: HTMLElement, source: string, node: Node | null, offset: number,
) {
  const element = node instanceof Element ? node : node?.parentElement;
  const block = element?.closest<HTMLElement>('[data-source-start][data-source-end]');
  if (!node || !block || !host.contains(block)) return null;
  const startLine = Number(block.dataset.sourceStart);
  const endLine = Number(block.dataset.sourceEnd);
  const lines = source.split('\n');
  if (!startLine || !endLine) return null;
  const from = lines.slice(0, startLine - 1).reduce((total, line) => total + line.length + 1, 0);
  const fragment = lines.slice(startLine - 1, endLine).join('\n');
  const range = document.createRange();
  range.selectNodeContents(block);
  range.setEnd(node, offset);
  return from + sourceOffsetForRenderedText(fragment, range.toString());
}
