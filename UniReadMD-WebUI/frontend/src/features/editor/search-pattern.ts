export interface SearchOptions {
  caseSensitive?: boolean;
  wholeWord?: boolean;
  regularExpression?: boolean;
  global?: boolean;
}

export function advanceSearchOffset(content: string, offset: number) {
  return offset + ((content.codePointAt(offset) ?? 0) > 0xFFFF ? 2 : 1);
}

export function buildSearchExpression(query: string, options: SearchOptions = {}) {
  if (!query) return null;
  const escaped = options.regularExpression
    ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const boundary = '[\\p{L}\\p{N}_]';
  const pattern = options.wholeWord
    ? `(?<!${boundary})(?:${escaped})(?!${boundary})` : escaped;
  try {
    return new RegExp(pattern,
      `${options.global ? 'g' : ''}${options.caseSensitive ? '' : 'i'}u`);
  } catch {
    return null;
  }
}

export function replaceSearchMatches(
  content: string, expression: RegExp, replacement: string, regularExpression = false,
) {
  // 普通搜索始终按字面文本替换；只有明确开启正则时才展开捕获组。
  return regularExpression
    ? content.replace(expression, replacement.replace(/\$0/g, '$$&'))
    : content.replace(expression, () => replacement);
}

export function replaceCurrentSearchMatch(
  content: string, expression: RegExp, offset: number, replacement: string,
  regularExpression = false,
) {
  // 在全文的指定位置匹配，以保留前后查找和单词边界的上下文。
  const flags = expression.flags.replace(/[gy]/g, '') + 'y';
  const current = new RegExp(expression.source, flags);
  current.lastIndex = offset;
  return replaceSearchMatches(content, current, replacement, regularExpression);
}
