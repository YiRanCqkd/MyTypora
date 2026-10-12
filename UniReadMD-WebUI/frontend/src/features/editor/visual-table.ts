import type { PreviewTableSelection } from './editor-contracts';
import { parseMarkdownTable, serializeMarkdownTable } from '../../utils/preview-table';

export function writeVisualTableCell(
  source: string, selection: PreviewTableSelection, value: string,
) {
  const table = parseMarkdownTable(source);
  const row = selection.rowIndex === 0 ? table?.header : table?.rows[selection.rowIndex - 1];
  if (!table || !row || selection.colIndex < 0 || selection.colIndex >= row.length) return null;
  // 单元格换行必须保留在同一 Markdown 行内，不能生成新的表格行。
  row[selection.colIndex] = value.replace(/\r?\n/g, '<br>').trim();
  return serializeMarkdownTable(table);
}

export function navigateVisualTable(
  source: string, selection: PreviewTableSelection, direction: 'next' | 'previous' | 'down',
) {
  const table = parseMarkdownTable(source);
  if (!table) return null;
  let { rowIndex, colIndex } = selection;
  if (direction === 'previous') {
    if (colIndex > 0) colIndex -= 1;
    else if (rowIndex > 0) { rowIndex -= 1; colIndex = table.header.length - 1; }
  } else if (direction === 'next' && colIndex < table.header.length - 1) colIndex += 1;
  else {
    rowIndex += 1;
    if (direction === 'next') colIndex = 0;
  }
  const appended = rowIndex > table.rows.length;
  if (appended) table.rows.push(Array.from({ length: table.header.length }, () => ''));
  const content = appended ? serializeMarkdownTable(table) : source;
  return { content, selection: { ...selection, rowIndex, colIndex,
    endLine: selection.startLine + content.split('\n').length - 1 } };
}
