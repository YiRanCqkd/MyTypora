import { describe, expect, it } from 'vitest';
import { navigateVisualTable, writeVisualTableCell } from './visual-table';

const source = '| 甲 | 乙 |\n| :---: | ---: |\n| **原文** | a\\|b |';
const selection = { startLine: 3, endLine: 5, rowIndex: 1, colIndex: 0 };

describe('表格可视编辑写回', () => {
  it('编辑省略竖线的数据行仍保留其他数据行', () => {
    expect(writeVisualTableCell('| A |\n| --- |\none\ntwo',
      { startLine: 1, endLine: 4, rowIndex: 2, colIndex: 0 }, '修改值'))
      .toBe('| A |\n| --- |\n| one |\n| 修改值 |');
  });

  it('保留列对齐、其他单元格格式及竖线转义', () => {
    expect(writeVisualTableCell(source, selection, '新|内容'))
      .toBe('| 甲 | 乙 |\n| :---: | ---: |\n| 新\\|内容 | a\\|b |');
  });
  it('单元格换行写为行内换行，不增加表格数据行', () => {
    expect(writeVisualTableCell(source, selection, '首行\n次行'))
      .toContain('| 首行<br>次行 | a\\|b |');
  });
  it('支持空内容及表头更新', () => {
    expect(writeVisualTableCell(source, { ...selection, rowIndex: 0 }, ''))
      .toContain('|  | 乙 |');
  });
  it('无效表格或单元格坐标不写入', () => {
    expect(writeVisualTableCell('普通正文', selection, '新内容')).toBeNull();
    expect(writeVisualTableCell(source, { ...selection, rowIndex: 9 }, '新内容')).toBeNull();
    expect(writeVisualTableCell(source, { ...selection, colIndex: 9 }, '新内容')).toBeNull();
  });
});

describe('表格可视编辑导航', () => {
  it('同一行 Tab 与反向 Tab 不改写源码', () => {
    expect(navigateVisualTable(source, selection, 'next'))
      .toMatchObject({ content: source, selection: { rowIndex: 1, colIndex: 1 } });
    expect(navigateVisualTable(source, selection, 'previous'))
      .toMatchObject({ content: source, selection: { rowIndex: 0, colIndex: 1 } });
  });
  it('表头首格反向 Tab 停留原位', () => {
    expect(navigateVisualTable(source, { ...selection, rowIndex: 0 }, 'previous'))
      .toMatchObject({ content: source, selection: { rowIndex: 0, colIndex: 0 } });
  });
  it('末格 Tab 追加空行，并更新表格源码范围', () => {
    expect(navigateVisualTable(source, { ...selection, colIndex: 1 }, 'next'))
      .toMatchObject({ content: source + '\n|  |  |',
        selection: { rowIndex: 2, colIndex: 0, endLine: 6 } });
  });
  it('Enter 保持列号，表头向下不追加行，末行向下追加行', () => {
    expect(navigateVisualTable(source, { ...selection, rowIndex: 0, colIndex: 1 }, 'down'))
      .toMatchObject({ content: source, selection: { rowIndex: 1, colIndex: 1 } });
    expect(navigateVisualTable(source, { ...selection, colIndex: 1 }, 'down'))
      .toMatchObject({ selection: { rowIndex: 2, colIndex: 1, endLine: 6 } });
  });
  it('无效表格不导航', () => {
    expect(navigateVisualTable('普通正文', selection, 'next')).toBeNull();
  });
});
