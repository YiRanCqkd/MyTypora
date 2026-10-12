import { describe, expect, it } from 'vitest';
import { findInlineFragment } from './inline-fragment';

describe('行内源码片段', () => {
  it('只选择点击的格式，并保留相邻与嵌套语法', () => {
    const text = '前 **粗 *斜* 体** 后 [链接](https://example.com)';
    expect(findInlineFragment(text, text.indexOf('斜'), 'em')).toEqual({ from: 6, to: 9 });
    const strong = findInlineFragment(text, text.indexOf('粗'), 'strong');
    expect(text.slice(strong!.from, strong!.to)).toBe('**粗 *斜* 体**');
    const link = findInlineFragment(text, text.indexOf('链接'), 'a');
    expect(text.slice(link!.from, link!.to)).toBe('[链接](https://example.com)');
  });
  it('代码围栏、普通文本和不匹配的 DOM 标签回退到块编辑', () => {
    expect(findInlineFragment('普通正文', 2, 'strong')).toBeNull();
    expect(findInlineFragment('```js\ncode\n```', 8, 'code')).toBeNull();
    expect(findInlineFragment('**加粗**', 3, 'em')).toBeNull();
  });
});
