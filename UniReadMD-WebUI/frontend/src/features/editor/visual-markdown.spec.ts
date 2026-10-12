import { describe, expect, it } from 'vitest';
import { applyVisualPrefix, describeVisualBlock, nextVisualPrefix,
  visualHtmlToMarkdown } from './visual-markdown';

describe('原位可视编辑 Markdown 转换', () => {
  it('忽略渲染器在换行标签后的排版换行，保留用户空格和连续换行标签', () => {
    expect(visualHtmlToMarkdown('首行<br>\n次行<br>\n末行')).toBe('首行\n次行\n末行');
    expect(visualHtmlToMarkdown('首行<br>\r\n<strong>次行</strong>'))
      .toBe('首行\n**次行**');
    expect(visualHtmlToMarkdown('首行<br>  次行')).toBe('首行\n  次行');
    expect(visualHtmlToMarkdown('首行<br>\n<br>\n次行')).toBe('首行\n\n次行');
    expect(visualHtmlToMarkdown('普通文本中的\n换行')).toBe('普通文本中的\n换行');
  });

  it('保留常用行内格式、连续空白、特殊字符及安全链接', () => {
    expect(visualHtmlToMarkdown('甲  <strong>乙</strong><em>丙</em><s>丁</s>'
      + '<mark>戊</mark><sub>a b</sub><sup>2</sup><br>*原文*'))
      .toBe('甲  **乙***丙*~~丁~~==戊==~a\\ b~^2^\n\\*原文\\*');
    expect(visualHtmlToMarkdown('<a href="https://example.com/a b" title="说明">链接</a>'))
      .toBe('[链接](<https://example.com/a%20b> "说明")');
    expect(visualHtmlToMarkdown('<a href="javascript:alert(1)">危险</a><script>坏</script>'))
      .toBe('危险');
  });

  it('处理行内代码中的反引号、浏览器段落与空白', () => {
    expect(visualHtmlToMarkdown('1. 普通正文')).toBe('1\\. 普通正文');
    expect(visualHtmlToMarkdown('<b> 加粗 </b><i> </i><strong></strong>'))
      .toBe(' **加粗**  ');
    expect(visualHtmlToMarkdown('<code>a`b</code> <code>`c`</code><div>下一段</div>'))
      .toBe('``a`b`` `` `c` ``\n\n下一段');
    expect(visualHtmlToMarkdown('<code>两字</code><b>粗</b><i>斜</i><del>删</del>'))
      .toBe('`两字`**粗***斜*~~删~~');
    expect(visualHtmlToMarkdown('<p># 标题</p><p>- 列表&nbsp;文字</p>'))
      .toBe('\\# 标题\n\n\\- 列表 文字');
  });

  it('识别普通块并保持标题、简单列表及引用的前缀', () => {
    expect(describeVisualBlock('正文')).toEqual({ prefix: '', kind: 'paragraph' });
    expect(describeVisualBlock('价格 $20')).toEqual({ prefix: '', kind: 'paragraph' });
    expect(describeVisualBlock('### 标题')).toEqual({ prefix: '### ', kind: 'heading' });
    const list = describeVisualBlock('3. 列表')!;
    expect(applyVisualPrefix('甲\n乙', list)).toBe('3. 甲\n   乙');
    expect(nextVisualPrefix(list)).toBe('4. ');
    const quote = describeVisualBlock('> 引用\n> 第二行')!;
    expect(applyVisualPrefix('甲\n乙', quote)).toBe('> 甲\n> 乙');
    expect(nextVisualPrefix(quote)).toBe('> ');
    expect(nextVisualPrefix({ prefix: '- ', kind: 'list' })).toBe('- ');
    expect(nextVisualPrefix({ prefix: '## ', kind: 'heading' })).toBe('');
    expect(applyVisualPrefix('新标题', { prefix: '# ', kind: 'heading' })).toBe('# 新标题');
  });

  it('拒绝可能损失语义的复杂块与伪造源映射', () => {
    for (const source of ['![图](a.png)', '<p data-source-start="1">内容</p>',
      '```js\ncode\n```', '    code', '$x$', '脚注[^1]', '- [ ] 任务',
      '- 父\n  - 子', '> > 引用', '> - 列表', '> ## 标题', '标题\n===', '[[toc]]', '---']) {
      expect(describeVisualBlock(source), source).toBeNull();
    }
  });
});
