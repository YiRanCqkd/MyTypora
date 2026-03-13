import { describe, expect, it } from 'vitest';
import {
  extractHeadings,
  findActiveHeadingLine,
  findMatchedLines,
  hasMathSyntax,
  validateMarkdownSyntax,
} from './markdown';

describe('markdown utils', () => {
  it('extractHeadings 应提取多级标题', () => {
    const input = '# A\ntext\n## B\n### C';

    const output = extractHeadings(input);

    expect(output).toEqual([
      { level: 1, text: 'A', line: 1 },
      { level: 2, text: 'B', line: 3 },
      { level: 3, text: 'C', line: 4 },
    ]);
  });

  it('extractHeadings 应忽略代码块中的伪标题并支持 Setext 标题', () => {
    const input = [
      '# A',
      '',
      '```md',
      '# fake',
      '## fake-2',
      '```',
      '',
      'Section Title',
      '---',
    ].join('\n');

    const output = extractHeadings(input);

    expect(output).toEqual([
      { level: 1, text: 'A', line: 1 },
      { level: 2, text: 'Section Title', line: 8 },
    ]);
  });

  it('findMatchedLines 应返回匹配行号', () => {
    const input = 'hello\nworld\nhello world';

    const output = findMatchedLines(input, 'world');

    expect(output).toEqual([
      { lineNo: 2, line: 'world' },
      { lineNo: 3, line: 'hello world' },
    ]);
  });

  it('findActiveHeadingLine 应返回当前激活标题行', () => {
    const headings = [
      { level: 1, text: 'A', line: 1 },
      { level: 2, text: 'B', line: 5 },
      { level: 2, text: 'C', line: 10 },
    ];

    expect(findActiveHeadingLine(headings, 1)).toBe(1);
    expect(findActiveHeadingLine(headings, 7)).toBe(5);
    expect(findActiveHeadingLine(headings, 999)).toBe(10);
  });

  it('hasMathSyntax 应判断数学语法标记', () => {
    expect(hasMathSyntax('plain text')).toBe(false);
    expect(hasMathSyntax('a + b = $c$')).toBe(true);
  });

  it('validateMarkdownSyntax 应识别未闭合代码块与行内数学公式', () => {
    const input = [
      'price = $a + b',
      '',
      '```ts',
      'const value = 1;',
    ].join('\n');

    const output = validateMarkdownSyntax(input);

    expect(output.some((item) => item.rule === 'fence')).toBe(true);
    expect(output.some((item) => item.rule === 'inline-math')).toBe(true);
  });

  it('validateMarkdownSyntax 应识别脚注与表格列数问题', () => {
    const input = [
      '| A | B |',
      '| --- | --- | --- |',
      '| 1 | 2 |',
      '[^miss]',
      '',
      '[^unused]: text',
    ].join('\n');

    const output = validateMarkdownSyntax(input);

    expect(output.some((item) => item.rule === 'table-header')).toBe(true);
    expect(output.some((item) => item.rule === 'footnote-ref')).toBe(true);
    expect(output.some((item) => item.rule === 'footnote-def')).toBe(true);
  });

  it('validateMarkdownSyntax 应识别任务列表、链接与图表语法问题', () => {
    const input = [
      '- [a] invalid task',
      '[bad](<https://exa mple.com>)',
      '',
      '```mermaid',
      'A-->B',
      '```',
      '',
      '```flow',
      'st=>start: Start',
      '```',
    ].join('\n');

    const output = validateMarkdownSyntax(input);

    expect(output.some((item) => item.rule === 'task-list')).toBe(true);
    expect(output.some((item) => item.rule === 'link-target')).toBe(true);
    expect(output.some((item) => item.rule === 'diagram-mermaid')).toBe(true);
    expect(output.some((item) => item.rule === 'diagram-flow')).toBe(true);
  });
});
