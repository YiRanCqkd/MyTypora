import { describe, expect, it } from 'vitest';
import { previewKeyEdit } from './preview-keyboard';

describe('预览连续编辑', () => {
  it('包围选区、跳过右括号和删除空配对', () => {
    expect(previewKeyEdit('中文', 0, 2, '(', 'paragraph_open'))
      .toEqual({ value: '(中文)', start: 1, end: 3 });
    expect(previewKeyEdit('()', 1, 1, ')', 'paragraph_open'))
      .toEqual({ value: '()', start: 2, end: 2 });
    expect(previewKeyEdit('()', 1, 1, 'Backspace', 'paragraph_open'))
      .toEqual({ value: '', start: 0, end: 0 });
  });
  it('正文回车只换一行，标题换段，代码块和软换行保持原生输入', () => {
    expect(previewKeyEdit('# 标题', 4, 4, 'Enter', 'heading_open')?.value).toBe('# 标题\n\n');
    expect(previewKeyEdit('正文', 2, 2, 'Enter', 'paragraph_open')?.value).toBe('正文\n');
    expect(previewKeyEdit('代码', 2, 2, 'Enter', 'fence')).toBeNull();
    expect(previewKeyEdit('正文', 2, 2, 'Enter', 'paragraph_open', true)).toBeNull();
  });
  it('续写有序列表、未完成任务和引用，空列表项退出列表', () => {
    expect(previewKeyEdit('9. 内容', 5, 5, 'Enter', 'list_item_open')?.value)
      .toBe('9. 内容\n10. ');
    expect(previewKeyEdit('> - [x] 完成', 10, 10, 'Enter', 'list_item_open')?.value)
      .toBe('> - [x] 完成\n> - [ ] ');
    expect(previewKeyEdit('- ', 2, 2, 'Enter', 'list_item_open')?.value).toBe('\n');
    expect(previewKeyEdit('> 引用', 4, 4, 'Enter', 'paragraph_open')?.value).toBe('> 引用\n> ');
  });
  it('正文单词内部的引号不自动配对', () => {
    expect(previewKeyEdit('dont', 3, 3, "'", 'paragraph_open')).toBeNull();
    expect(previewKeyEdit('dont', 4, 4, "'", 'paragraph_open')).toBeNull();
  });
  it('空正文和普通文字的退格交由编辑器处理', () => {
    expect(previewKeyEdit('', 0, 0, 'Backspace', 'paragraph_open')).toBeNull();
    expect(previewKeyEdit('abc', 3, 3, 'Backspace', 'paragraph_open')).toBeNull();
    expect(previewKeyEdit('abc', 0, 0, 'Backspace', 'paragraph_open')).toBeNull();
  });
});
