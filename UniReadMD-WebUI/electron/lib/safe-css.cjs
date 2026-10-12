const postcss = require('postcss');
const selectorParser = require('postcss-selector-parser');
const { createContractError } = require('./errors.cjs');

function sanitizeCss(text, { scoped = true } = {}) {
  const maxBytes = (scoped ? 1 : 20) * 1024 * 1024;
  if (typeof text !== 'string' || Buffer.byteLength(text) > maxBytes) {
    throw createContractError('ERR_INVALID_CSS', `样式超过 ${scoped ? 1 : 20} MiB 限制`);
  }
  let root;
  try {
    root = postcss.parse(text);
  } catch {
    throw createContractError('ERR_INVALID_CSS', '样式文件语法错误');
  }
  root.walkAtRules((rule) => {
    if (!['media', 'supports', 'font-face'].includes(rule.name.toLowerCase())
      || (scoped && rule.name.toLowerCase() === 'font-face')) rule.remove();
  });
  root.walkDecls((declaration) => {
    if (/expression\s*\(|javascript:|vbscript:|behavior|binding/i.test(declaration.toString())
      || (scoped && (/url\s*\(/i.test(declaration.value)
        || declaration.value.includes('\\')))) declaration.remove();
  });
  if (scoped) root.walkRules((rule) => {
    try {
      rule.selector = selectorParser((selectors) => {
        selectors.each((selector) => {
          const first = selector.nodes.find((node) => node.type !== 'comment');
          if (!first || selector.nodes.some((node) => node.type === 'nesting')) {
            selector.remove();
            return;
          }
          if ((first.type === 'tag' && ['html', 'body'].includes(first.value))
            || (first.type === 'id' && first.value === 'write')
            || (first.type === 'pseudo' && first.value === ':root')) {
            first.replaceWith(selectorParser.className({ value: 'preview' }));
          }
          const head = selector.nodes.find((node) => node.type !== 'comment');
          const atRoot = head.type === 'class' && head.value === 'preview';
          const combinator = selector.nodes.find((node) => node.type === 'combinator');
          // 根节点的相邻选择器会越出预览；内部段落的相邻选择器仍可正常使用。
          if (atRoot && combinator && !['', '>'].includes(combinator.value.trim())) {
            selector.remove();
          } else if (!atRoot) {
            selector.prepend(selectorParser.combinator({ value: ' ' }));
            selector.prepend(selectorParser.className({ value: 'preview' }));
          }
        });
      }).processSync(rule.selector);
      if (!rule.selector.trim()) rule.remove();
    } catch {
      throw createContractError('ERR_INVALID_CSS', '样式选择器语法错误');
    }
  });
  return root.toString().replace(/</g, '\\3c ');
}

module.exports = { sanitizeCss };
