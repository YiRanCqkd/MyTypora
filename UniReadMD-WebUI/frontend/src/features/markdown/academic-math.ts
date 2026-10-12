import type MarkdownIt from 'markdown-it';
import { readWritingPreferences } from '../settings/writing-preferences';

export interface AcademicMath {
  labels: Map<string, string>;
  macros: Record<string, string | object>;
}

export function installAcademicMath(markdown: MarkdownIt) {
  markdown.core.ruler.after('inline', 'academic_math', (state) => {
    const context: AcademicMath = { labels: new Map(), macros: {} };
    state.env.academicMath = context;
    let index = 0;
    const automatic = readWritingPreferences().equationNumbering;
    for (const token of state.tokens) {
      if (token.type !== 'math_block') continue;
      const labels = [...token.content.matchAll(/\\label\{([^{}]{1,200})\}/g)]
        .map((match) => match[1]);
      const manual = /\\tag\*?\{([^{}]{1,100})\}/.exec(token.content)?.[1];
      const number = manual || (automatic ? String(++index) : '');
      if (manual && automatic) index += 1;
      for (const label of labels) if (!context.labels.has(label)) {
        context.labels.set(label, number || label);
      }
      token.meta = { equationLabel: labels[0] || '' };
      token.content = token.content.replace(/\\label\{([^{}]{1,200})\}/g, '');
      if (automatic && !manual) token.content += `\\tag{${number}}`;
    }
  });
}

export function substituteMathReferences(content: string, context?: AcademicMath) {
  return content.replace(/\\(eqref|ref)\{([^{}]{1,200})\}/g, (_match, kind, key) => {
    const number = context?.labels.get(key) || '?';
    return kind === 'eqref' ? `(${number})` : number;
  });
}
