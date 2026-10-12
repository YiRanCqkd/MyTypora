import { describe, expect, it } from 'vitest';
import { formatPreviewRange, sourceOffsetForRenderedText } from './preview-source-offset';

describe('预览光标映射', () => {
  it('代码块隐藏围栏与语言名，光标仍指向实际代码', () => {
    const source = '```js\nconst x = 1;\n```';
    expect(sourceOffsetForRenderedText(source, 'const x')).toBe(source.indexOf('const x') + 7);
    expect(sourceOffsetForRenderedText(source, '')).toBe(source.indexOf('const'));
  });
  it('跨块加粗分别包裹正文，空行和列表结构保持原样', () => {
    const source = '第一段\n\n- 第二段';
    expect(formatPreviewRange(source, 0, source.length, 'bold'))
      .toBe('**第一段**\n\n- **第二段**');
  });
  it('标题中的嵌套格式与链接文字映射到实际内容', () => {
    const source = '# 标题 **加粗** [链接](a.md) `code`';
    expect(sourceOffsetForRenderedText(source, '标题 加')).toBe(source.indexOf('加') + 1);
    expect(sourceOffsetForRenderedText(source, '标题 加粗 链'))
      .toBe(source.indexOf('链') + 1);
    expect(sourceOffsetForRenderedText(source, '标题 加粗 链接 co'))
      .toBe(source.indexOf('code') + 2);
  });
  it('引用列表、实体、转义与软换行保持源码位置', () => {
    const source = '> - **甲** &amp; \\*乙\n>   丙';
    expect(sourceOffsetForRenderedText(source, '甲 & *乙')).toBe(source.indexOf('乙') + 1);
    expect(sourceOffsetForRenderedText(source, '甲 & *乙丙')).toBe(source.length);
    expect(sourceOffsetForRenderedText('## 标题', '')).toBe(3);
  });
});
