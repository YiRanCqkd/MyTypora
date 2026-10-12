const createDOMPurify = require('dompurify');
const { JSDOM } = require('jsdom');
const { sanitizeCss } = require('./safe-css.cjs');

const EXPORT_CSP = "default-src 'none'; style-src 'unsafe-inline'; "
  + "img-src data: https: http:; font-src data:; object-src 'none'; base-uri 'none'";

async function prepareExportHtml(html, event, assets) {
  // 不启用脚本和外部资源；导出输入也按不可信内容处理。
  const window = new JSDOM(html).window;
  try {
    const document = window.document;
    const title = document.title;
    const language = /^[a-z]{2,3}(?:-[a-zA-Z0-9]{2,8})*$/.test(document.documentElement.lang)
      ? document.documentElement.lang : 'zh-CN';
    const theme = document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light';
    const css = [...document.querySelectorAll('style')]
      .map((node) => sanitizeCss(node.textContent || '', { scoped: false })).join('\n');
    const purifier = createDOMPurify(window);
    const body = purifier.sanitize(document.body.innerHTML, {
      USE_PROFILES: { html: true, svg: true, mathMl: true },
      FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'style', 'link', 'form',
        'foreignObject', 'animate', 'set'],
      FORBID_ATTR: ['srcset', 'formaction'],
      ALLOW_DATA_ATTR: false,
    });
    document.body.innerHTML = body;
    for (const image of document.querySelectorAll('img')) {
      const source = image.getAttribute('src') || '';
      if (!source.includes('/__unireadmd_asset__/')) continue;
      if (!assets) throw new Error('无法读取导出图片');
      const url = new URL(source);
      const token = url.pathname.split('/').at(-1);
      const resource = await assets.readForOwner(event, token, url.searchParams.get('path'));
      image.src = `data:${resource.contentType};base64,${resource.content.toString('base64')}`;
    }
    document.head.innerHTML = '';
    document.title = title;
    document.documentElement.lang = language;
    document.documentElement.dataset.theme = theme;
    const encoding = document.createElement('meta');
    encoding.setAttribute('charset', 'UTF-8');
    const policy = document.createElement('meta');
    policy.httpEquiv = 'Content-Security-Policy';
    policy.content = EXPORT_CSP;
    const style = document.createElement('style');
    style.textContent = css;
    document.head.prepend(encoding, policy, style);
    return `<!doctype html>\n${document.documentElement.outerHTML}`;
  } finally {
    window.close();
  }
}

module.exports = { prepareExportHtml };
