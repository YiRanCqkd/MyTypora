import { describe, expect, it } from 'vitest';
import { highlightPreviewCode, syncPreviewCodeScroll } from './preview-code-presentation';

describe('代码块源码展示', () => {
  it('保留原始围栏、空行与内容，按语言高亮且不执行代码里的 HTML', () => {
    const source = '```js\nconst html = "<img src=x>";\n\n```';
    const html = highlightPreviewCode(source, 'fence');
    const node = document.createElement('pre');
    node.innerHTML = html;
    expect(node.textContent).toBe(source + '\u200b');
    expect(node.querySelector('.hljs-keyword')?.textContent).toBe('const');
    expect(node.querySelectorAll('.preview-code-fence')).toHaveLength(2);
    expect(node.querySelector('img')).toBeNull();
  });

  it('支持波浪围栏、语言大小写和附加参数', () => {
    const html = highlightPreviewCode('  ~~~~JS title\nlet x = 1;\n ~~~~~ ', 'fence');
    expect(html).toContain('hljs-keyword');
    expect(html.match(/preview-code-fence/g)).toHaveLength(2);
  });

  it('未知语言和缩进代码安全退回原文', () => {
    const node = document.createElement('pre');
    for (const [source, token] of [['```unknown\n<&>\n```', 'fence'],
      ['    <script>&</script>', 'code_block'], ['invalid <&>', 'fence']]) {
      node.innerHTML = highlightPreviewCode(source!, token!);
      expect(node.textContent).toBe(source + '\u200b');
      expect(node.querySelector('script')).toBeNull();
    }
    expect(highlightPreviewCode('正文', 'paragraph_open')).toBe('');
  });

  it('未闭合、过短或不同字符的围栏继续作为代码展示', () => {
    for (const source of ['```', '```js\nlet a = 1;', '~~~~\n~~~', '~~~\n```']) {
      const node = document.createElement('pre');
      node.innerHTML = highlightPreviewCode(source, 'fence');
      expect(node.textContent).toBe(source + '\u200b');
      expect(node.querySelectorAll('.preview-code-fence')).toHaveLength(1);
    }
  });

  it('滚动只同步相邻的高亮镜像，不影响其它源码输入框', () => {
    const host = document.createElement('section');
    const highlight = document.createElement('pre');
    const textarea = document.createElement('textarea');
    highlight.className = 'preview-code-highlight';
    host.append(highlight, textarea);
    textarea.addEventListener('scroll', syncPreviewCodeScroll);
    textarea.scrollLeft = 120;
    textarea.scrollTop = 30;
    textarea.dispatchEvent(new Event('scroll'));
    expect(highlight.scrollLeft).toBe(120);
    expect(highlight.scrollTop).toBe(30);
    highlight.className = '';
    textarea.scrollLeft = 300;
    textarea.dispatchEvent(new Event('scroll'));
    expect(highlight.scrollLeft).toBe(120);
    highlight.remove();
    expect(() => textarea.dispatchEvent(new Event('scroll'))).not.toThrow();
    expect(() => syncPreviewCodeScroll(new Event('scroll'))).not.toThrow();
  });
});
