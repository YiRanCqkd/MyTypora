import { describe, expect, it } from 'vitest';
import {
  applyPreviewTableActionToModel,
  parseMarkdownTable,
  serializeMarkdownTable,
} from './preview-table';

describe('preview table utils', () => {
  it('parseMarkdownTable + serializeMarkdownTable 应能补齐空单元格', () => {
    const input = [
      '| A | B |',
      '| --- | --- |',
      '| 1 |',
    ].join('\n');

    const table = parseMarkdownTable(input);

    expect(table).not.toBeNull();
    expect(table?.header).toEqual([
      'A',
      'B',
    ]);
    expect(table?.rows[0]).toEqual([
      '1',
      '',
    ]);
    expect(serializeMarkdownTable(table!)).toContain('| 1 |  |');
  });

  it('applyPreviewTableActionToModel 应支持行插入和行移动', () => {
    const table = {
      header: [
        'A',
        'B',
      ],
      rows: [
        [
          '1',
          '2',
        ],
        [
          '3',
          '4',
        ],
      ],
    };

    const insertResult = applyPreviewTableActionToModel(
      table,
      {
        rowIndex: 1,
        colIndex: 0,
      },
      'insert-row-below',
    );
    expect(insertResult).toEqual({
      changed: true,
    });
    expect(table.rows).toEqual([
      [
        '1',
        '2',
      ],
      [
        '',
        '',
      ],
      [
        '3',
        '4',
      ],
    ]);

    const moveResult = applyPreviewTableActionToModel(
      table,
      {
        rowIndex: 2,
        colIndex: 0,
      },
      'move-row-up',
    );
    expect(moveResult).toEqual({
      changed: true,
    });
    expect(table.rows[0][0]).toBe('');
  });

  it('applyPreviewTableActionToModel 应返回列删除边界错误信息', () => {
    const table = {
      header: [
        'Only',
      ],
      rows: [
        [
          '1',
        ],
      ],
    };

    const result = applyPreviewTableActionToModel(
      table,
      {
        rowIndex: 1,
        colIndex: 0,
      },
      'delete-col',
    );

    expect(result).toEqual({
      changed: false,
      errorMessage: '表格至少需要保留一列。',
    });
  });

  it('applyPreviewTableActionToModel 应支持列移动并保持列顺序同步', () => {
    const table = {
      header: [
        'A',
        'B',
        'C',
      ],
      rows: [
        [
          '1',
          '2',
          '3',
        ],
      ],
    };

    const result = applyPreviewTableActionToModel(
      table,
      {
        rowIndex: 1,
        colIndex: 2,
      },
      'move-col-left',
    );

    expect(result).toEqual({
      changed: true,
    });
    expect(table.header).toEqual([
      'A',
      'C',
      'B',
    ]);
    expect(table.rows[0]).toEqual([
      '1',
      '3',
      '2',
    ]);
  });
});
