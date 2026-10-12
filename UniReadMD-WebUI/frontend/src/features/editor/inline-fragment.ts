import { markdownLanguage } from '@codemirror/lang-markdown';

const INLINE_NODES: Record<string, string> = {
  strong: 'StrongEmphasis', em: 'Emphasis', code: 'InlineCode', a: 'Link',
};

export function findInlineFragment(source: string, offset: number, tag: string):
  { from: number; to: number } | null {
  const name = INLINE_NODES[tag];
  if (!name) return null;
  let result: { from: number; to: number } | null = null;
  markdownLanguage.parser.parse(source).iterate({ enter(node) {
    if (node.name === name && node.from <= offset && node.to >= offset
      && (!result || node.to - node.from < result.to - result.from)) {
      result = { from: node.from, to: node.to };
    }
  } });
  return result;
}
