import { describe, expect, it } from 'vitest';
import { buildSearchExpression, replaceCurrentSearchMatch, replaceSearchMatches }
  from './search-pattern';
import { toggleTaskLine } from './task-source';
import { countDocumentWords } from './document-statistics';

describe('Markdown 编辑正确性', () => {
  it('正则单次替换保留前后查找和全文替换变量的上下文', () => {
    const pattern = buildSearchExpression('(?<=a)(b)(?=c)', {
      regularExpression: true, global: true,
    })!;
    expect(replaceCurrentSearchMatch('abc abc', pattern, 1, '$1-$0', true))
      .toBe('ab-bc abc');
    expect(replaceCurrentSearchMatch('abc abc', pattern, 5, '$&', false))
      .toBe('abc a$&c');
  });
  it('中文全词匹配保留独立词且排除较长单词的一部分', () => {
    const expression = buildSearchExpression('测试', { wholeWord: true, global: true })!;
    expect('测试\n测试 测试 测试文档'.match(expression)).toHaveLength(3);
  });
  it('搜索保留首尾空格并正确转义特殊符号', () => {
    expect('a foo b foo'.match(buildSearchExpression(' foo ', { global: true })!))
      .toEqual([' foo ']);
    expect('a.b a-b'.match(buildSearchExpression('a.b', { global: true })!))
      .toEqual(['a.b']);
  });
  it('字面批量替换不展开美元字符', () => {
    const expression = buildSearchExpression('foo', { global: true })!;
    expect(replaceSearchMatches('foo foo', expression, '$&')).toBe('$& $&');
    expect(replaceSearchMatches('foo foo', expression, '$1 $$')).toBe('$1 $$ $1 $$');
  });
  it('正则捕获组替换且非法正则可安全拒绝', () => {
    const expression = buildSearchExpression('(foo)', { regularExpression: true, global: true })!;
    expect(replaceSearchMatches('foo foo', expression, '$1-$0', true)).toBe('foo-foo foo-foo');
    expect(buildSearchExpression('[', { regularExpression: true })).toBeNull();
  });
  it('引用和嵌套任务仅修改任务标记', () => {
    expect(toggleTaskLine('> - [ ] 引用任务')).toBe('> - [x] 引用任务');
    expect(toggleTaskLine('  > > + [X] 完成')).toBe('  > > + [ ] 完成');
    expect(toggleTaskLine('普通段落')).toBeNull();
  });
  it('中文逐字统计，英文逐词统计，不包含 Markdown 格式符号', () => {
    expect(countDocumentWords('中文内容'.repeat(500))).toBe(2000);
    expect(countDocumentWords('# 标题\n\n- **hello** world')).toBe(4);
    expect(countDocumentWords('hello中文world')).toBe(4);
  });
});
