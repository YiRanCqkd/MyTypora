import { describe, expect, it } from 'vitest';
import { sanitizeMarkdownHtml } from './safe-html';

describe('不可信 Markdown 渲染', () => {
  it('Markdown URL 编码只解码一次，支持中文与 Windows 图片路径', () => {
    const base = 'http://127.0.0.1:1234/__unireadmd_asset__/token?path=';
    const html = sanitizeMarkdownHtml(
      '<img src="%E5%9B%BE%E7%89%87/a%20b.png#片段"><img src="C:/图片/a.png">',
      { resourceBaseUrl: base },
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    const images = [...host.querySelectorAll('img')];
    expect(new URL(images[0].src).searchParams.get('path')).toBe('图片/a b.png');
    expect(new URL(images[1].src).searchParams.get('path')).toBe('C:/图片/a.png');
  });
  it('过滤事件、危险协议、嵌入框架与伪造源映射', () => {
    const html = sanitizeMarkdownHtml([
      '<img src="x" onerror="alert(1)">',
      '<a href="javascript:alert(1)">链接</a>',
      '<iframe src="https://example.com"></iframe>',
      '<img src="DATA:image/svg+xml;base64,PHN2Zz4=">',
      '<p data-source-start="1" data-diagram-code="伪造">正文</p>',
    ].join(''), { trustedAttributes: false });
    expect(html).not.toMatch(/onerror|javascript:|iframe|data-source|data-diagram|DATA:image/);
    expect(html).toContain('正文');
  });

  it('保留可信编辑映射和公式，按文档能力解析中文图片', () => {
    const html = sanitizeMarkdownHtml(
      '<p data-source-start="2"><img src="图片/a.png"><math><mi>x</mi></math></p>',
      { resourceBaseUrl: 'http://127.0.0.1:1234/__unireadmd_asset__/token?path=' },
    );
    expect(html).toContain('data-source-start="2"');
    expect(html).toContain('%E5%9B%BE%E7%89%87%2Fa.png');
    expect(html).toContain('<math>');
  });

  it('保留 SVG 图表，过滤外部对象和样式资源加载', () => {
    const html = sanitizeMarkdownHtml(
      '<svg><path d="M0 0 L1 1"></path><foreignObject>攻击</foreignObject></svg>'
      + '<p style="background:url(https://example.com)">文字</p>',
    );
    expect(html).toContain('<path');
    expect(html).not.toMatch(/foreignObject|url\(/);
  });

  it('拒绝图片危险协议，保留安全图片与任务复选框', () => {
    const html = sanitizeMarkdownHtml(
      '<img src="javascript:alert(1)"><img src="file:///C:/秘密.png">'
      + '<img src="ftp://example.com/x.png"><img src="data:image/png;base64,AA==">'
      + '<img src="坏%路径.png"><a href="https://example.com">链接</a>'
      + '<input type="text"><input type="checkbox">'
      + '<input type="checkbox" class="task-list-item-checkbox" checked>'
      + '<p style="position:fixed">危险样式</p><span style="height:1em">公式尺寸</span>',
      { resourceBaseUrl: 'http://127.0.0.1:1/assets?path=' },
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    const images = [...host.querySelectorAll('img')];
    expect(images.slice(0, 3).every((item) => !item.hasAttribute('src'))).toBe(true);
    expect(images[3].getAttribute('src')).toBe('data:image/png;base64,AA==');
    expect(new URL(images[4].src).searchParams.get('path')).toBe('坏%路径.png');
    expect(host.querySelectorAll('input')).toHaveLength(1);
    expect(host.querySelector('a')?.rel).toBe('noopener noreferrer');
    expect(host.querySelector('p')?.hasAttribute('style')).toBe(false);
    expect(host.querySelector('span')?.style.height).toBe('1em');
  });

  it('图表样式绑定 SVG，拒绝无根样式和外部资源', () => {
    const html = sanitizeMarkdownHtml(
      '<svg id="diagram"><style>.node{fill:red;background:url(https://example.com);'
      + 'font-family:\\72 oboto} #diagram .edge{stroke:blue}'
      + ' #diagram + .toolbar{color:red}'
      + ' @keyframes dash{to{stroke-dashoffset:0}}'
      + ' @media screen{.edge{stroke:blue}}</style><path class="node"></path></svg>'
      + '<style>.app{display:none}</style>'
      + '<svg id="9invalid"><style>.node{fill:red}</style></svg>'
      + '<svg><style>.toolbar{display:none}</style></svg>'
      + '<svg id="unsafe"><style>@import "https://example.com";</style></svg>'
      + '<svg id="invalid"><style>p{</style></svg>',
      { diagramStyles: true },
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    expect(host.querySelectorAll('style')).toHaveLength(1);
    expect(host.querySelector('style')?.textContent).toContain(':is(.node)');
    expect(host.querySelector('style')?.textContent).toContain(':is(#diagram .edge)');
    expect(host.querySelector('style')?.textContent).not.toMatch(/@keyframes|url\(|\\72/);
    const outside = document.createElement('div');
    outside.className = 'toolbar';
    host.append(outside);
    const selectors = host.querySelector('style')!.textContent!.match(/([^{}]+)\{/g)!;
    expect(host.querySelectorAll(selectors[0].slice(0, -1))).toHaveLength(1);
    expect(host.querySelectorAll(selectors[2].slice(0, -1))).toHaveLength(0);
  });
});
