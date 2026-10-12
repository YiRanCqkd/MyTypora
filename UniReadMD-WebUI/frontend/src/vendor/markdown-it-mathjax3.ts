import katex from 'katex';
import 'katex/contrib/mhchem';
import { installAcademicMath, substituteMathReferences, type AcademicMath }
  from '../features/markdown/academic-math';
import type MarkdownIt from 'markdown-it';
import type StateBlock from 'markdown-it/lib/rules_block/state_block.mjs';
import type StateInline from 'markdown-it/lib/rules_inline/state_inline.mjs';

const ESCAPE_RE = /[&<>"']/g;
const ESCAPE_MAP: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  '\'': '&#39;',
};

function escapeHtml(input: string): string {
  return input.replace(ESCAPE_RE, (ch) => ESCAPE_MAP[ch] || ch);
}

function renderMath(content: string, display: boolean, context?: AcademicMath): string {
  try {
    return katex.renderToString(substituteMathReferences(content, context), {
      displayMode: display,
      throwOnError: false,
      strict: 'ignore',
      output: 'htmlAndMathml',
      trust: false,
      macros: context?.macros,
      globalGroup: true,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const source = escapeHtml(content);
    const detail = escapeHtml(message);
    return `<span class='math-error' title='${detail}'>${source}</span>`;
  }
}

function isValidDelim(state: StateInline, pos: number): {
  can_open: boolean;
  can_close: boolean;
} {
  const max = state.posMax;
  const prevChar = pos > 0 ? state.src.charCodeAt(pos - 1) : -1;
  const nextChar = pos + 1 <= max ? state.src.charCodeAt(pos + 1) : -1;

  let canOpen = true;
  let canClose = true;

  if (prevChar === 0x20 || prevChar === 0x09 || (nextChar >= 0x30 && nextChar <= 0x39)) {
    canClose = false;
  }
  if (nextChar === 0x20 || nextChar === 0x09) {
    canOpen = false;
  }

  return {
    can_open: canOpen,
    can_close: canClose,
  };
}

function mathInline(state: StateInline, silent: boolean): boolean {
  if (state.src[state.pos] !== '$') {
    return false;
  }

  let res = isValidDelim(state, state.pos);
  if (!res.can_open) {
    if (!silent) {
      state.pending += '$';
    }
    state.pos += 1;
    return true;
  }

  const start = state.pos + 1;
  let match = start;
  while ((match = state.src.indexOf('$', match)) !== -1) {
    let pos = match - 1;
    while (state.src[pos] === '\\') {
      pos -= 1;
    }
    if ((match - pos) % 2 === 1) {
      break;
    }
    match += 1;
  }

  if (match === -1) {
    if (!silent) {
      state.pending += '$';
    }
    state.pos = start;
    return true;
  }

  if (match - start === 0) {
    if (!silent) {
      state.pending += '$$';
    }
    state.pos = start + 1;
    return true;
  }

  res = isValidDelim(state, match);
  if (!res.can_close) {
    if (!silent) {
      state.pending += '$';
    }
    state.pos = start;
    return true;
  }

  if (!silent) {
    const token = state.push('math_inline', 'math', 0);
    token.markup = '$';
    token.content = state.src.slice(start, match);
  }

  state.pos = match + 1;
  return true;
}

function mathBlock(
  state: StateBlock,
  start: number,
  end: number,
  silent: boolean,
): boolean {
  let next = start;
  let lastPos = 0;
  let found = false;
  let pos = state.bMarks[start] + state.tShift[start];
  let max = state.eMarks[start];
  let lastLine = '';

  if (pos + 2 > max) {
    return false;
  }
  if (state.src.slice(pos, pos + 2) !== '$$') {
    return false;
  }

  pos += 2;
  let firstLine = state.src.slice(pos, max);
  if (silent) {
    return true;
  }

  if (firstLine.trim().slice(-2) === '$$') {
    firstLine = firstLine.trim().slice(0, -2);
    found = true;
  }

  while (!found) {
    next += 1;
    if (next >= end) {
      break;
    }

    pos = state.bMarks[next] + state.tShift[next];
    max = state.eMarks[next];
    if (pos < max && state.tShift[next] < state.blkIndent) {
      break;
    }

    if (state.src.slice(pos, max).trim().slice(-2) === '$$') {
      lastPos = state.src.slice(0, max).lastIndexOf('$$');
      lastLine = state.src.slice(pos, lastPos);
      found = true;
    }
  }

  state.line = next + 1;
  const token = state.push('math_block', 'math', 0);
  token.block = true;
  token.content = `${firstLine && firstLine.trim() ? `${firstLine}\n` : ''}`
    + `${state.getLines(start + 1, next, state.tShift[start], true)}`
    + `${lastLine && lastLine.trim() ? lastLine : ''}`;
  token.map = [start, state.line];
  token.markup = '$$';
  return true;
}

const markdownItMathJax3 = (md: MarkdownIt) => {
  md.inline.ruler.after('escape', 'math_inline', mathInline);
  md.block.ruler.after('blockquote', 'math_block', mathBlock, {
    alt: ['paragraph', 'reference', 'blockquote', 'list'],
  });
  installAcademicMath(md);
  md.renderer.rules.math_inline = (tokens, idx, _options, env) => {
    const reference = /^\\(eqref|ref)\{([^{}]{1,200})\}$/.exec(tokens[idx].content.trim());
    const context = env.academicMath as AcademicMath | undefined;
    if (reference && context?.labels.has(reference[2])) {
      const label = md.utils.escapeHtml(substituteMathReferences(tokens[idx].content, context));
      const anchor = md.utils.escapeHtml(encodeURIComponent(reference[2]));
      return `<a class='math-reference' href='#equation-${anchor}'`
        + `>${label}</a>`;
    }
    return renderMath(tokens[idx].content, false, context);
  };
  md.renderer.rules.math_block = (tokens, idx, _options, env) => {
    const token = tokens[idx];
    const start = (token.map?.[0] || 0) + 1;
    const end = token.map?.[1] || start;
    const label = token.meta?.equationLabel;
    const id = label ? ` id='equation-${md.utils.escapeHtml(encodeURIComponent(label))}'` : '';
    return `<div class='math-block'${id} data-source-start='${start}'`
      + ` data-source-end='${end}' data-source-token='math_block'>`
      + renderMath(token.content, true, env.academicMath) + '</div>\n';
  };
};

export default markdownItMathJax3;
