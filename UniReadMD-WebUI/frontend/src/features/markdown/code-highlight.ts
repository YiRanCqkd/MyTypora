import hljs from 'highlight.js/lib/core';
import javascript from 'highlight.js/lib/languages/javascript';
import typescript from 'highlight.js/lib/languages/typescript';
import python from 'highlight.js/lib/languages/python';
import cpp from 'highlight.js/lib/languages/cpp';
import java from 'highlight.js/lib/languages/java';
import json from 'highlight.js/lib/languages/json';
import bash from 'highlight.js/lib/languages/bash';
import xml from 'highlight.js/lib/languages/xml';
import css from 'highlight.js/lib/languages/css';
import sql from 'highlight.js/lib/languages/sql';
import markdown from 'highlight.js/lib/languages/markdown';
import yaml from 'highlight.js/lib/languages/yaml';
import rust from 'highlight.js/lib/languages/rust';
import go from 'highlight.js/lib/languages/go';

Object.entries({ javascript, typescript, python, cpp, java, json, bash, xml,
  css, sql, markdown, yaml, rust, go,
}).forEach(([name, grammar]) => hljs.registerLanguage(name, grammar));

export function highlightCode(source: string, language: string) {
  const normalized = ({ js: 'javascript', ts: 'typescript', html: 'xml',
    sh: 'bash', shell: 'bash', c: 'cpp', yml: 'yaml', py: 'python',
  } as Record<string, string>)[language] || language;
  if (!hljs.getLanguage(normalized)) return '';
  return hljs.highlight(source, { language: normalized, ignoreIllegals: true }).value;
}

export function headingSlug(text: string) {
  return `md-${text.trim().toLowerCase().replace(/\s+/g, '-')
    .replace(/[^\p{L}\p{N}_-]/gu, '') || 'heading'}`;
}
