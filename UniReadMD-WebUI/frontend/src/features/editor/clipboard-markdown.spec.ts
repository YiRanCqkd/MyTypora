import { describe, expect, it } from 'vitest';
import { htmlToMarkdown, markdownClipboardHtml } from './clipboard-markdown';

describe('富文本剪贴板', () => {
  it('保留标题、加粗、链接、嵌套列表和代码中的原始空白', () => {
    const markdown = htmlToMarkdown('<h2>标题</h2><p><b>粗体</b> '
      + '<a href="https://example.com/a">链接</a></p><ul><li>外层<ul><li>内层</li></ul>'
      + '</li></ul><pre><code class="language-js">a\n\n\n  b</code></pre>');
    expect(markdown).toContain('## 标题');
    expect(markdown).toContain('**粗体** [链接](https://example.com/a)');
    expect(markdown).toContain('  - 内层');
    expect(markdown).toContain('```js\na\n\n\n  b\n```');
  });

  it('表格单元格竖线转义，脚本与危险链接不可执行', () => {
    const markdown = htmlToMarkdown('<script>alert(1)</script><p>'
      + '<a href="javascript:alert(1)">文字</a></p><table><tr><th>甲</th><th>乙</th></tr>'
      + '<tr><td>a|b</td><td>c</td></tr></table>');
    expect(markdown).not.toContain('javascript:');
    expect(markdown).not.toContain('alert(1)');
    expect(markdown).toContain('| a\\|b | c |');
    expect(markdownClipboardHtml('**粗体**\n\n<script>')).toContain('<strong>粗体</strong>');
    expect(markdownClipboardHtml('<script>')).not.toContain('<script>');
  });

  it('行内代码含反引号时选择不会截断内容的分隔符', () => {
    expect(htmlToMarkdown('<code>a`b</code>')).toBe('``a`b``');
    expect(() => htmlToMarkdown('a'.repeat(5 * 1024 * 1024 + 1))).toThrow('5 MiB');
  });
  it('相邻段落不生成多余空段落，保留 Word 常见 span 样式', () => {
    expect(htmlToMarkdown('<p>甲</p><p>乙</p>')).toBe('甲\n\n乙');
    expect(htmlToMarkdown('<span style="font-weight:700">粗体</span>')).toBe('**粗体**');
    expect(htmlToMarkdown('<span style="font-style:italic">斜体</span>')).toBe('*斜体*');
  });
});
