import { describe, expect, it } from 'vitest';
import {
  applyPreviewTableActionToModel,
  parseMarkdownTable,
  serializeMarkdownTable,
} from './preview-table';

describe('preview table utils', () => {
  it('编辑与移动列应保持对齐、转义竖线和代码单元格', () => {
    const table = parseMarkdownTable(
      '| 中间 | 右侧 |\n| :---: | ---: |\n| a\\|b | `x\\|y` |',
    )!;
    expect(table.header).toHaveLength(2);
    expect(table.rows[0]).toEqual(['a\\|b', '`x\\|y`']);
    table.rows[0][0] = '新值|尾部';
    applyPreviewTableActionToModel(table, { rowIndex: 1, colIndex: 1 }, 'move-col-left');
    const output = serializeMarkdownTable(table);
    expect(output).toContain('| ---: | :---: |');
    expect(output).toContain('| `x\\|y` | 新值\\|尾部 |');
    expect(parseMarkdownTable(output)?.rows[0]).toHaveLength(2);
    applyPreviewTableActionToModel(table, { rowIndex: 1, colIndex: 0 }, 'insert-col-left');
    expect(table.delimiters).toEqual(['---', '---:', ':---:']);
    applyPreviewTableActionToModel(table, { rowIndex: 1, colIndex: 1 }, 'delete-col');
    expect(table.delimiters).toEqual(['---', ':---:']);
  });

  it('奇偶转义与无边框末尾转义竖线应正确解析，错误分隔行应拒绝', () => {
    expect(parseMarkdownTable('A | B\n--- | ---\nx | y\\|')?.rows[0])
      .toEqual(['x', 'y\\|']);
    expect(parseMarkdownTable('| A | B |\n| --- | --- |\n| x\\\\ | y |')?.rows[0])
      .toEqual(['x\\\\', 'y']);
    expect(parseMarkdownTable('A | B\n普通 | 文字')).toBeNull();
  });

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

  it('parseMarkdownTable 应拒绝不完整输入并规范无边框表格', () => {
    expect(parseMarkdownTable('only one line')).toBeNull();
    expect(parseMarkdownTable('A\n---')).toBeNull();
    expect(parseMarkdownTable('A | B\n--- | ---\n1 | 2')).toEqual({
      header: ['A', 'B'],
      rows: [['1', '2']],
    });
  });

  it('serializeMarkdownTable 应补齐短行并截断长行', () => {
    expect(serializeMarkdownTable({
      header: ['A', 'B'],
      rows: [
        ['1'],
        ['2', '3', '4'],
      ],
    })).toBe([
      '| A | B |',
      '| --- | --- |',
      '| 1 |  |',
      '| 2 | 3 |',
    ].join('\n'));
  });

  it('行操作应覆盖插入、删除、上下移动及边界', () => {
    const table = {
      header: ['A'],
      rows: [['1'], ['2']],
    };
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 1,
      colIndex: 0,
    }, 'insert-row-above').changed).toBe(true);
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 0,
      colIndex: 0,
    }, 'delete-row')).toMatchObject({ changed: false });
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 2,
      colIndex: 0,
    }, 'delete-row').changed).toBe(true);
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 1,
      colIndex: 0,
    }, 'move-row-up')).toMatchObject({ changed: false });
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 1,
      colIndex: 0,
    }, 'move-row-down').changed).toBe(true);
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 3,
      colIndex: 0,
    }, 'move-row-down')).toMatchObject({ changed: false });
  });

  it('列操作应覆盖插入、删除、左右移动及边界', () => {
    const table = {
      header: ['A', 'B'],
      rows: [['1', '2']],
    };
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 1,
      colIndex: 0,
    }, 'insert-col-left').changed).toBe(true);
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 1,
      colIndex: 1,
    }, 'insert-col-right').changed).toBe(true);
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 1,
      colIndex: 99,
    }, 'delete-col').changed).toBe(true);
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 1,
      colIndex: 0,
    }, 'move-col-left')).toMatchObject({ changed: false });
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 1,
      colIndex: table.header.length - 1,
    }, 'move-col-right')).toMatchObject({ changed: false });
    expect(applyPreviewTableActionToModel(table, {
      rowIndex: 1,
      colIndex: 0,
    }, 'move-col-right').changed).toBe(true);
  });
});
