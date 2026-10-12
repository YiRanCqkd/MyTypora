import MarkdownIt from 'markdown-it';
import { describe, expect, it } from 'vitest';
import { buildImageResizeSource, findPreviewImageSource, resizeImageSource } from './image-sizing';

const markdown = new MarkdownIt({ html: true });
function fixture(source: string) {
  const host = document.createElement('div');
  host.innerHTML = markdown.render(source);
  return { host, images: [...host.querySelectorAll('img')] };
}

describe('预览图片尺寸源码操作', () => {
  it('同段重复图片按实际位置定位，只调整选中的图片', () => {
    const source = '![图](a.png) 和 ![图](a.png)';
    const { host, images } = fixture(source);
    const selection = findPreviewImageSource(source, images[1], host, markdown);
    expect(selection?.from).toBe(14);
    const insert = resizeImageSource(selection!, 800, 50);
    expect(insert).toContain('width="400"');
    expect(insert).toContain('height: auto');
    expect(source.slice(0, selection!.from)).toBe('![图](a.png) 和 ');
  });

  it('引用图片解析原路径和标题，外层链接、引用定义保持原样', () => {
    const source = '[![中文][pic]](https://example.com)\n\n[pic]: <图 片.png> "说明"';
    const { host, images } = fixture(source);
    const selection = findPreviewImageSource(source, images[0], host, markdown)!;
    const insert = resizeImageSource(selection, 640, 75);
    const result = source.slice(0, selection.from) + insert + source.slice(selection.to);
    expect(insert).toContain('src="%E5%9B%BE%20%E7%89%87.png"');
    expect(insert).toContain('title="说明"');
    expect(result).toContain('](https://example.com)');
    expect(result).toContain('[pic]: <图 片.png> "说明"');
  });

  it('HTML 图片保留无关样式和属性，清除拉伸限制并恢复自动', () => {
    const source = '<img src="x.png" alt="图" title="标题" width="120" height="20" '
      + 'style="border-radius: 4px; width: 120px; max-height: 20px; zoom: 2">';
    const { host, images } = fixture(source);
    const selection = findPreviewImageSource(source, images[0], host, markdown)!;
    const scaled = resizeImageSource(selection, 800, 25);
    expect(scaled).toContain('border-radius: 4px');
    expect(scaled).toContain('width="200"');
    expect(scaled).not.toContain('height="20"');
    expect(scaled).not.toContain('max-height');
    expect(scaled).not.toContain('zoom');
    const automatic = resizeImageSource(selection, 0, null);
    expect(automatic).not.toMatch(/\b(?:width|height)=/);
    expect(automatic).toContain('border-radius: 4px');
  });

  it('排除代码、注释中的伪图片，并支持列表内图片', () => {
    const source = '```md\n![假](fake.png)\n```\n\n<!-- <img src="fake.png"> -->\n\n'
      + '- ![真](real.png)';
    const { host, images } = fixture(source);
    const selection = findPreviewImageSource(source, images[0], host, markdown)!;
    expect(source.slice(selection.from, selection.to)).toBe('![真](real.png)');
  });

  it('源映射块限制目标范围，资源代理地址不写回文档', () => {
    const source = '![同图](a.png)\n\n![同图](a.png)';
    const { host, images } = fixture(source);
    images[1].parentElement!.dataset.sourceStart = '3';
    images[1].parentElement!.dataset.sourceEnd = '3';
    images[1].setAttribute('src', 'https://assets.test/image?path=a.png');
    const selection = findPreviewImageSource(source, images[1], host, markdown,
      'https://assets.test/image?path=')!;
    expect(selection.from).toBe(source.lastIndexOf('!['));
    expect(resizeImageSource(selection, 800, 50)).toContain('src="a.png"');
  });

  it('拒绝过期或无法对应的图片，校验比例与原图尺寸', () => {
    const { host, images } = fixture('![图](a.png)');
    expect(findPreviewImageSource('![图](b.png)', images[0], host, markdown)).toBeNull();
    expect(findPreviewImageSource('无图片', images[0], host, markdown)).toBeNull();
    const selection = findPreviewImageSource('![图](a.png)', images[0], host, markdown)!;
    for (const percent of [0, -1, 401, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(() => resizeImageSource(selection, 800, percent)).toThrow();
    }
    expect(() => resizeImageSource(selection, 0, 50)).toThrow();
    expect(resizeImageSource(selection, 1, 1)).toContain('width="1"');
    expect(resizeImageSource(selection, 800, 12.5)).toContain('width="100"');
  });

  it('HTML 块忽略注释与脚本中的图片并保留真实标签，拒绝无图片和宿主外目标', () => {
    const source = '<div><!-- <img src="fake.png"> -->'
      + '<script>const text = \'<img src="fake.png">\';</script>'
      + '<style>.x { content: \'<img src="fake.png">\'; }</style>'
      + '<img src="real.png"></div>';
    const { host, images } = fixture(source);
    const selection = findPreviewImageSource(source, images[0], host, markdown)!;
    expect(selection.html).toBe('<img src="real.png">');
    expect(resizeImageSource(selection, 0, null)).toBe('<img src="real.png">');
    expect(findPreviewImageSource(source, document.createElement('img'), host, markdown))
      .toBeNull();
    expect(() => resizeImageSource({ from: 0, to: 0, html: '<p>正文</p>' }, 0, null)).toThrow();
    images[0].alt = '过期说明';
    expect(findPreviewImageSource(source, images[0], host, markdown)).toBeNull();
  });

  it('单行图片改为 HTML 后继续渲染后面的 Markdown，行内图片不添加空行', () => {
    for (const ending of ['\n', '\r\n']) {
      const source = `![图](a.png)${ending}后文 **加粗**`;
      const { host, images } = fixture(source);
      const range = findPreviewImageSource(source, images[0], host, markdown)!;
      const insert = buildImageResizeSource(source, range, 800, 50);
      const result = source.slice(0, range.from) + insert + source.slice(range.to);
      expect(markdown.render(result)).toContain('<strong>加粗</strong>');
    }
    const source = '前文 ![图](a.png) 后文';
    const { host, images } = fixture(source);
    const range = findPreviewImageSource(source, images[0], host, markdown)!;
    expect(buildImageResizeSource(source, range, 800, 50)).not.toContain('\n');
  });
});
