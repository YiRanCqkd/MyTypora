import { afterEach, describe, expect, it } from 'vitest';
import { createPreviewSelection } from './preview-selection';
import { useEditorHosts } from './editor-hosts';

afterEach(() => { document.body.replaceChildren(); });

describe('图表与正文点击的源码映射', () => {
  it('代码文字和背景留白统一定位到完整代码容器', () => {
    const host = document.createElement('article');
    host.innerHTML = '<pre data-source-start="1" data-source-end="3" '
      + 'data-source-token="fence"><code data-source-start="1" data-source-end="3" '
      + 'data-source-token="fence"><span>const</span> a = 1;</code></pre>';
    const selection = createPreviewSelection(useEditorHosts());
    expect(selection.findSourceMappedNode(host.querySelector('span'))).toBe(host.firstChild);
    expect(selection.findSourceMappedNode(host.querySelector('pre'))).toBe(host.firstChild);
  });

  it('点击 SVG 节点文字可定位到外层 HTML 图表源码块', () => {
    const host = document.createElement('article');
    host.innerHTML = '<div data-source-start="1" data-source-end="5" '
      + 'data-source-token="fence"><svg><g><text><tspan>开始</tspan></text></g></svg></div>';
    document.body.append(host);
    const selection = createPreviewSelection(useEditorHosts());
    const mapped = selection.findSourceMappedNode(host.querySelector('tspan'));
    expect(mapped).toBe(host.firstElementChild);
    expect(selection.parseInlineEditSelection(mapped!))
      .toEqual({ startLine: 1, endLine: 5, token: 'fence' });
  });

  it('普通 HTML 子节点仍定位到正文块，非元素和无映射节点返回空值', () => {
    const host = document.createElement('article');
    host.innerHTML = '<p data-source-start="1" data-source-end="1" '
      + 'data-source-token="paragraph_open"><strong>正文</strong></p><p>无映射</p>';
    const selection = createPreviewSelection(useEditorHosts());
    expect(selection.findSourceMappedNode(host.querySelector('strong')))
      .toBe(host.firstElementChild);
    expect(selection.findSourceMappedNode(host.lastElementChild)).toBeNull();
    expect(selection.findSourceMappedNode(null)).toBeNull();
    expect(selection.findSourceMappedNode(document.createTextNode('文字'))).toBeNull();
  });

  it('SVG 本身的映射属性不能冒充 HTML 编辑块', () => {
    const host = document.createElement('article');
    host.innerHTML = '<svg data-source-start="1" data-source-end="5" '
      + 'data-source-token="fence"><text>伪造块</text></svg>';
    const selection = createPreviewSelection(useEditorHosts());
    expect(selection.findSourceMappedNode(host.querySelector('text'))).toBeNull();
  });
});
