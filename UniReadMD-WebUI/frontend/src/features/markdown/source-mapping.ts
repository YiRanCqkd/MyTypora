import type { MarkdownSourceToken } from './renderer-contracts';
import type { createMarkdownRenderer } from './renderer';
type CreateMarkdownSourceMappingDependencies =
  ReturnType<typeof createMarkdownRenderer>;

export function createMarkdownSourceMapping(input: CreateMarkdownSourceMappingDependencies) {
  const { SOURCE_MAPPED_BLOCK_TOKENS,markdown,
  } = input;

  function annotateTokenSourceRange(token: MarkdownSourceToken) {
    if (!token.map || typeof token.attrSet !== 'function') {
      return;
    }

    const startLine = token.map[0] + 1;
    const endLine = Math.max(startLine, token.map[1]);
    token.attrSet('data-source-start', String(startLine));
    token.attrSet('data-source-end', String(endLine));
    token.attrSet('data-source-token', token.type);
  }

  function installSourceMapRendererRule(ruleName: string) {
    const fallbackRule = markdown.renderer.rules[ruleName];
    markdown.renderer.rules[ruleName] = (tokens, idx, options, env, self) => {
      const token = tokens[idx] as MarkdownSourceToken;
      annotateTokenSourceRange(token);

      if (fallbackRule) {
        return fallbackRule(tokens, idx, options, env, self);
      }

      return self.renderToken(tokens, idx, options);
    };
  }

  SOURCE_MAPPED_BLOCK_TOKENS.forEach((ruleName) => {
    installSourceMapRendererRule(ruleName);
  });

  const originalFenceRule = markdown.renderer.rules.fence;

  const diagramFenceLanguages = new Set(['mermaid', 'sequence', 'flow']);

  markdown.renderer.rules.fence = (tokens, idx, options, env, self) => {
    const token = tokens[idx] as MarkdownSourceToken;
    annotateTokenSourceRange(token);

    const sourceStart = token.attrGet?.('data-source-start') || '';
    const sourceEnd = token.attrGet?.('data-source-end') || '';
    const sourceToken = token.attrGet?.('data-source-token') || 'fence';
    const sourceAttrs = sourceStart && sourceEnd
      ? ` data-source-start='${sourceStart}' data-source-end='${sourceEnd}'`
        + ` data-source-token='${sourceToken}'`
      : '';

    const info = (token.info || '').trim().split(/\s+/)[0] || '';

    if (!diagramFenceLanguages.has(info)) {
      if (originalFenceRule) {
        return originalFenceRule(tokens, idx, options, env, self);
      }

      return self.renderToken(tokens, idx, options);
    }

    const escaped = markdown.utils.escapeHtml(token.content || '');
    const encoded = encodeURIComponent(token.content || '');
    const diagramAttrs = [
      ` data-diagram-kind='${info}'`,
      ` data-diagram-code='${encoded}'`,
    ].join('');

    return `
  <div class='mermaid-block'${sourceAttrs}${diagramAttrs}>
    <div class='mermaid-output'></div>
    <pre class='mermaid-fallback'>${escaped}</pre>
    <div class='mermaid-error'></div>
  </div>`;
  };

  const originalLinkOpenRule = markdown.renderer.rules.link_open;

  markdown.renderer.rules.link_open = (tokens, idx, options, env, self) => {
    const token = tokens[idx] as MarkdownSourceToken;
    token.attrSet?.('target', '_blank');
    token.attrSet?.('rel', 'noopener noreferrer');

    if (originalLinkOpenRule) {
      return originalLinkOpenRule(tokens, idx, options, env, self);
    }

    return self.renderToken(tokens, idx, options);
  };

  return { annotateTokenSourceRange,installSourceMapRendererRule,originalFenceRule,
    diagramFenceLanguages,originalLinkOpenRule,
  };
}
