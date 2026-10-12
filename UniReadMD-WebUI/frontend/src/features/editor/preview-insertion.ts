export type PreviewInsertKind = 'paragraph' | 'heading1' | 'heading2' | 'heading3'
  | 'heading4' | 'heading5' | 'heading6' | 'table' | 'code' | 'math' | 'quote'
  | 'unordered' | 'ordered' | 'task' | 'rule';

export function buildPreviewInsertion(source: string, offset: number, kind: PreviewInsertKind,
  rows = 2, columns = 2) {
  const position = Math.max(0, Math.min(source.length, offset));
  let text: string;
  if (/^heading[1-6]$/.test(kind)) text = `${'#'.repeat(Number(kind.at(-1)))} 标题`;
  else if (kind === 'table') {
    if (!Number.isInteger(rows) || rows < 1 || rows > 20
      || !Number.isInteger(columns) || columns < 1 || columns > 12) {
      throw new Error('表格行数须为 1～20，列数须为 1～12。');
    }
    const line = (cells: string[]) => `| ${cells.join(' | ')} |`;
    text = [line(Array.from({ length: columns }, (_, index) => `列${index + 1}`)),
      line(Array(columns).fill('---')),
      ...Array.from({ length: rows }, () => line(Array(columns).fill('内容')))].join('\n');
  } else {
    text = { paragraph: '正文', code: '```\n代码\n```', math: '$$\n公式\n$$',
      quote: '> 引用', unordered: '- 列表项', ordered: '1. 列表项',
      task: '- [ ] 任务', rule: '***' }[kind as Exclude<PreviewInsertKind,
        'table' | 'heading1' | 'heading2' | 'heading3' | 'heading4' | 'heading5' | 'heading6'>];
  }
  const before = source.slice(0, position);
  const after = source.slice(position);
  const prefix = before ? '\n'.repeat(Math.max(0, 2 - (/\n*$/.exec(before)?.[0].length || 0))) : '';
  const suffix = after ? '\n'.repeat(Math.max(0, 2 - (/^\n*/.exec(after)?.[0].length || 0))) : '';
  return { from: position, insert: prefix + text + suffix, contentFrom: position + prefix.length,
    contentTo: position + prefix.length + text.length };
}
