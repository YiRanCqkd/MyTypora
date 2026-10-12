import type Token from 'markdown-it/lib/token.mjs';

let cachedContent: string | null = null;
let cachedWordCount = 0;

export function rememberDocumentWords(content: string, tokens: Token[]) {
  const text = tokens.flatMap((token) => {
    if (token.children) {
      return token.children.map((child) => ['text', 'code_inline'].includes(child.type)
        ? child.content : ' ');
    }
    return ['fence', 'code_block'].includes(token.type) ? [token.content] : [];
  }).join(' ');
  cachedContent = content;
  cachedWordCount = countPlainWords(text);
  return cachedWordCount;
}

export function countPlainWords(text: string) {
  const cjk = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/gu;
  let cjkCount = 0;
  const remaining = text.replace(cjk, () => { cjkCount += 1; return ' '; });
  const otherWords = remaining.match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu);
  return cjkCount + (otherWords?.length || 0);
}

export function getCachedDocumentWords(content: string) {
  return content === cachedContent ? cachedWordCount : null;
}

export function setCachedDocumentWords(content: string, words: number) {
  cachedContent = content;
  cachedWordCount = words;
}
