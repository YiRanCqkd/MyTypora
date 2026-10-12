import { describe, expect, it } from 'vitest';
import { buildPreviewInsertion } from './preview-insertion';

describe('预览插入块', () => {
  it('空文档、段落间和文末插入标题，保留原正文和必要的空行', () => {
    expect(buildPreviewInsertion('', 0, 'heading2').insert).toBe('## 标题');
    const source = '前文\n\n后文';
    const patch = buildPreviewInsertion(source, 2, 'heading1');
    expect(source.slice(0, patch.from) + patch.insert + source.slice(patch.from))
      .toBe('前文\n\n# 标题\n\n后文');
    expect(buildPreviewInsertion('正文', 999, 'quote').insert).toBe('\n\n> 引用');
    expect(buildPreviewInsertion('正文\n\n', 4, 'task').contentFrom).toBe(4);
  });
  it('生成指定行列的标准表格，拒绝过大和非法尺寸', () => {
    const patch = buildPreviewInsertion('', 0, 'table', 3, 4);
    expect(patch.insert.split('\n')).toHaveLength(5);
    expect(patch.insert.split('\n')[0]).toBe('| 列1 | 列2 | 列3 | 列4 |');
    for (const [rows, columns] of [[0, 2], [21, 2], [2, 13], [2, 0], [1.5, 2]]) {
      expect(() => buildPreviewInsertion('', 0, 'table', rows, columns)).toThrow();
    }
  });
  it('常用块均生成可继续编辑的占位内容', () => {
    expect(buildPreviewInsertion('', -1, 'paragraph').insert).toBe('正文');
    const fence = String.fromCharCode(96).repeat(3);
    expect(buildPreviewInsertion('', 0, 'code').insert).toBe(fence + '\n代码\n' + fence);
    expect(buildPreviewInsertion('', 0, 'math').insert).toBe('$$\n公式\n$$');
    expect(buildPreviewInsertion('', 0, 'rule').insert).toBe('***');
    expect(buildPreviewInsertion('', 0, 'unordered').insert).toBe('- 列表项');
    expect(buildPreviewInsertion('', 0, 'ordered').insert).toBe('1. 列表项');
  });
});
