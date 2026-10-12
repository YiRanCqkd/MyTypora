import DOMPurify from 'dompurify';
import postcss from 'postcss';

const FORBIDDEN_TAGS = [
  'script', 'iframe', 'object', 'embed', 'form', 'button', 'textarea', 'select',
  'style', 'link', 'meta', 'base', 'foreignObject', 'animate', 'set',
];

export function sanitizeMarkdownHtml(html: string, options?: {
  trustedAttributes?: boolean;
  resourceBaseUrl?: string;
  diagramStyles?: boolean;
}) {
  const safe = DOMPurify.sanitize(html, {
    USE_PROFILES: { html: true, svg: true, mathMl: true },
    FORBID_TAGS: options?.diagramStyles
      ? FORBIDDEN_TAGS.filter((tag) => tag !== 'style') : FORBIDDEN_TAGS,
    ADD_TAGS: options?.diagramStyles ? ['style'] : [],
    FORBID_ATTR: ['srcset', 'formaction', 'autofocus', 'name'],
    ALLOW_DATA_ATTR: false,
    ADD_URI_SAFE_ATTR: ['src'],
    ADD_ATTR: options?.trustedAttributes === false ? [] : [
      'data-source-start', 'data-source-end', 'data-source-token',
      'data-diagram-kind', 'data-diagram-code',
    ],
  });
  const host = document.createElement('template');
  host.innerHTML = safe;
  for (const style of host.content.querySelectorAll('style')) {
    const svg = style.closest('svg');
    const css = style.textContent || '';
    if (!svg?.id || !/^[a-z_][\w-]*$/i.test(svg.id)) {
      style.remove();
      continue;
    }
    try {
      const rules = postcss.parse(css);
      // 移除导入与动画规则，保留同一图表内的静态样式。
      rules.walkAtRules((rule) => {
        if (!['media', 'supports'].includes(rule.name.toLowerCase())) rule.remove();
      });
      rules.walkDecls((declaration) => {
        if (/url\(\s*[^#]|expression|javascript|behavior/i.test(declaration.toString())
          || declaration.value.includes('\\')) declaration.remove();
      });
      rules.walkRules((rule) => {
        rule.selectors = rule.selectors.map((selector) => {
          return `:where(#${svg.id}, #${svg.id} *):is(${selector.trim()})`;
        });
      });
      style.textContent = rules.toString();
      if (!style.textContent.trim()) style.remove();
    } catch {
      style.remove();
    }
  }
  for (const node of host.content.querySelectorAll<HTMLElement>('[style]')) {
    const style = node.getAttribute('style') || '';
    if (/url\s*\(|@import|expression\s*\(|behavior\s*:|position\s*:/i.test(style)) {
      node.removeAttribute('style');
    }
  }
  for (const input of host.content.querySelectorAll('input')) {
    if (input.type !== 'checkbox' || !input.classList.contains('task-list-item-checkbox')) {
      input.remove();
    }
  }
  for (const node of host.content.querySelectorAll<HTMLElement>('[href], [src]')) {
    for (const attribute of ['href', 'src']) {
      const value = node.getAttribute(attribute);
      if (!value) continue;
      const normalized = Array.from(value).filter((char) => char.charCodeAt(0) > 32).join('');
      if (/^(?:javascript|vbscript|file|blob):/i.test(normalized)
        || (/^data:/i.test(normalized) && !/^data:image\/(png|jpeg|gif|webp);/i.test(value))) {
        node.removeAttribute(attribute);
        continue;
      }
      const isLocal = !/^(?:[a-z][a-z\d+.-]*:|\/\/|#)/i.test(value)
        || /^[a-z]:[\\/]/i.test(value);
      if (attribute === 'src' && isLocal && options?.resourceBaseUrl) {
        let reference = value.split(/[?#]/)[0];
        try { reference = decodeURIComponent(reference); } catch { /* 保留字面百分号文件名。 */ }
        node.setAttribute('src', options.resourceBaseUrl + encodeURIComponent(reference));
      } else if (attribute === 'src' && !isLocal
        && !/^(https?:|data:image\/|\/\/|#)/i.test(value)) {
        node.removeAttribute('src');
      }
      if (attribute === 'href' && node.tagName.toLowerCase() === 'a') {
        if (/^(https?:|mailto:|tel:|\/\/)/i.test(value)) node.setAttribute('target', '_blank');
        else node.removeAttribute('target');
        node.setAttribute('rel', 'noopener noreferrer');
      }
    }
  }
  return host.innerHTML;
}
