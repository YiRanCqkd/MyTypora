export interface ParsedMarkdownTable {
  header: string[];
  rows: string[][];
  delimiters?: string[];
}

export interface PreviewTableCellSelection {
  rowIndex: number;
  colIndex: number;
}

export type PreviewTableAction =
  | 'insert-row-above'
  | 'insert-row-below'
  | 'delete-row'
  | 'move-row-up'
  | 'move-row-down'
  | 'insert-col-left'
  | 'insert-col-right'
  | 'delete-col'
  | 'move-col-left'
  | 'move-col-right'
  | 'align-col-left'
  | 'align-col-center'
  | 'align-col-right';

export interface PreviewTableActionResult {
  changed: boolean;
  errorMessage?: string;
}

function parseMarkdownTableLine(line: string) {
  let normalized = line.trim();
  if (normalized.startsWith('|')) {
    normalized = normalized.slice(1);
  }
  if (normalized.endsWith('|') && !/(?:^|[^\\])(?:\\\\)*\\\|$/.test(normalized)) {
    normalized = normalized.slice(0, -1);
  }
  if (!normalized) {
    return [''];
  }

  // GFM 中只有未转义的竖线是列分隔符，代码片段中的竖线也必须转义。
  const cells: string[] = [];
  let start = 0;
  let backslashes = 0;
  for (let index = 0; index < normalized.length; index += 1) {
    const char = normalized[index];
    if (char === '|' && backslashes % 2 === 0) {
      cells.push(normalized.slice(start, index).trim());
      start = index + 1;
    }
    backslashes = char === '\\' ? backslashes + 1 : 0;
  }
  cells.push(normalized.slice(start).trim());
  return cells;
}

function buildMarkdownTableLine(cells: string[]) {
  const escaped = cells.map((cell) => cell.replace(/(\\*)\|/g, (match, slashes: string) => {
    return slashes.length % 2 ? match : `${slashes}\\|`;
  }).replace(/\r?\n/g, '<br>'));
  return `| ${escaped.join(' | ')} |`;
}

export function parseMarkdownTable(source: string): ParsedMarkdownTable | null {
  const lines = source
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  if (lines.length < 2) {
    return null;
  }

  if (!lines[0].includes('|') || !lines[1].includes('|')) {
    return null;
  }

  const header = parseMarkdownTableLine(lines[0]);
  const delimiters = parseMarkdownTableLine(lines[1]);
  if (delimiters.length !== header.length
    || delimiters.some((item) => !/^:?-+:?$/.test(item))) return null;
  const rows = lines
    .slice(2)
    // 数据行可以省略竖线；短行由 normalizeRow 补齐，不能直接丢弃。
    .map((line) => parseMarkdownTableLine(line));
  const colCount = Math.max(
    1,
    header.length,
    ...rows.map((row) => row.length),
  );
  const normalizeRow = (row: string[]) => {
    const nextRow = [...row];
    while (nextRow.length < colCount) {
      nextRow.push('');
    }
    return nextRow.slice(0, colCount);
  };

  const table: ParsedMarkdownTable = {
    header: normalizeRow(header),
    rows: rows.map((row) => normalizeRow(row)),
  };
  // 普通三横线表格保持旧模型形状，现有调用无需迁移。
  if (delimiters.some((item) => item !== '---')) {
    table.delimiters = [...delimiters];
    while (table.delimiters.length < colCount) table.delimiters.push('---');
  }
  return table;
}

export function serializeMarkdownTable(table: ParsedMarkdownTable) {
  const colCount = Math.max(1, table.header.length);
  const delimiter = buildMarkdownTableLine(Array.from({ length: colCount }, (_, index) => {
    return table.delimiters?.[index] || '---';
  }));
  const body = table.rows.map((row) => {
    const normalized = [...row];
    while (normalized.length < colCount) {
      normalized.push('');
    }
    return buildMarkdownTableLine(normalized.slice(0, colCount));
  });

  return [
    buildMarkdownTableLine(table.header.slice(0, colCount)),
    delimiter,
    ...body,
  ].join('\n');
}

function moveTableRow(rows: string[][], from: number, to: number) {
  if (from < 0 || from >= rows.length || to < 0 || to >= rows.length) {
    return false;
  }

  const [current] = rows.splice(from, 1);
  rows.splice(to, 0, current);
  return true;
}

function moveTableColumn(table: ParsedMarkdownTable, from: number, to: number) {
  const colCount = table.header.length;
  if (colCount <= 1) {
    return false;
  }

  if (from < 0 || from >= colCount || to < 0 || to >= colCount) {
    return false;
  }

  const moveRow = (row: string[]) => {
    const [current] = row.splice(from, 1);
    row.splice(to, 0, current);
  };
  moveRow(table.header);
  if (table.delimiters) moveRow(table.delimiters);
  table.rows.forEach((row) => moveRow(row));
  return true;
}

export function applyPreviewTableActionToModel(
  table: ParsedMarkdownTable,
  selection: PreviewTableCellSelection,
  action: PreviewTableAction,
): PreviewTableActionResult {
  const bodyIndex = selection.rowIndex - 1;
  const colCount = table.header.length;
  const createEmptyRow = () => new Array(colCount).fill('');
  const alignment: Record<string, string> = {
    'align-col-left': ':---', 'align-col-center': ':---:', 'align-col-right': '---:',
  };
  if (alignment[action]) {
    table.delimiters ||= new Array(colCount).fill('---');
    const index = Math.max(0, Math.min(selection.colIndex, colCount - 1));
    const changed = table.delimiters[index] !== alignment[action];
    table.delimiters[index] = alignment[action];
    return { changed };
  }

  if (action === 'insert-row-above') {
    const insertAt = Math.max(0, bodyIndex);
    table.rows.splice(insertAt, 0, createEmptyRow());
    return {
      changed: true,
    };
  }

  if (action === 'insert-row-below') {
    const insertAt = bodyIndex < 0 ? 0 : bodyIndex + 1;
    table.rows.splice(insertAt, 0, createEmptyRow());
    return {
      changed: true,
    };
  }

  if (action === 'delete-row') {
    if (bodyIndex < 0 || table.rows.length === 0) {
      return {
        changed: false,
        errorMessage: '表头行不支持删除。',
      };
    }

    const deleteAt = Math.min(bodyIndex, table.rows.length - 1);
    table.rows.splice(deleteAt, 1);
    return {
      changed: true,
    };
  }

  if (action === 'move-row-up') {
    if (bodyIndex <= 0 || table.rows.length <= 1) {
      return {
        changed: false,
        errorMessage: '当前行已位于顶部或无法继续上移。',
      };
    }

    return {
      changed: moveTableRow(table.rows, bodyIndex, bodyIndex - 1),
    };
  }

  if (action === 'move-row-down') {
    if (bodyIndex < 0 || bodyIndex >= table.rows.length - 1) {
      return {
        changed: false,
        errorMessage: '当前行已位于底部或无法继续下移。',
      };
    }

    return {
      changed: moveTableRow(table.rows, bodyIndex, bodyIndex + 1),
    };
  }

  if (action === 'insert-col-left' || action === 'insert-col-right') {
    const offset = action === 'insert-col-left' ? 0 : 1;
    const insertAt = Math.min(
      Math.max(selection.colIndex + offset, 0),
      table.header.length,
    );
    table.header.splice(insertAt, 0, '');
    table.delimiters?.splice(insertAt, 0, '---');
    table.rows.forEach((row) => {
      row.splice(insertAt, 0, '');
    });
    return {
      changed: true,
    };
  }

  if (action === 'delete-col') {
    if (table.header.length <= 1) {
      return {
        changed: false,
        errorMessage: '表格至少需要保留一列。',
      };
    }

    const deleteAt = Math.min(
      Math.max(selection.colIndex, 0),
      table.header.length - 1,
    );
    table.header.splice(deleteAt, 1);
    table.delimiters?.splice(deleteAt, 1);
    table.rows.forEach((row) => {
      row.splice(deleteAt, 1);
    });
    return {
      changed: true,
    };
  }

  if (action === 'move-col-left') {
    const target = selection.colIndex - 1;
    if (target < 0) {
      return {
        changed: false,
        errorMessage: '当前列已位于最左侧。',
      };
    }

    return {
      changed: moveTableColumn(table, selection.colIndex, target),
    };
  }

  if (action === 'move-col-right') {
    const target = selection.colIndex + 1;
    if (target >= table.header.length) {
      return {
        changed: false,
        errorMessage: '当前列已位于最右侧。',
      };
    }

    return {
      changed: moveTableColumn(table, selection.colIndex, target),
    };
  }

  return {
    changed: false,
  };
}
