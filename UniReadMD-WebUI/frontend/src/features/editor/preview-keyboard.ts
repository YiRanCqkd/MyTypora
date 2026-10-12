export function previewKeyEdit(
  value: string, from: number, to: number, key: string, token: string, shift = false,
) {
  const pairs: Record<string, string> = { '(': ')', '[': ']', '{': '}', '"': '"', "'": "'" };
  const edit = (insert: string, start = from + insert.length, end = start) => ({
    value: value.slice(0, from) + insert + value.slice(to), start, end,
  });
  if (from === to && [')', ']', '}', '"', "'"].includes(key) && value[from] === key) {
    return { value, start: from + 1, end: from + 1 };
  }
  if (key === 'Backspace' && from === to && from > 0
    && pairs[value[from - 1]] && pairs[value[from - 1]] === value[from]) {
    return { value: value.slice(0, from - 1) + value.slice(from + 1),
      start: from - 1, end: from - 1 };
  }
  if (pairs[key]) {
    if (from === to && /[\p{L}\p{N}]/u.test(value[from] || '')) return null;
    if (from === to && ['"', "'"].includes(key)
      && /[\p{L}\p{N}]/u.test(value[from - 1] || '')) return null;
    return edit(key + value.slice(from, to) + pairs[key], from + 1, to + 1);
  }
  if (key === 'Enter' && !shift && !['fence', 'code_block'].includes(token)) {
    const lineStart = value.lastIndexOf('\n', from - 1) + 1;
    const line = value.slice(lineStart, from);
    const list = /^(\s*(?:>\s*)*)([-+*]|\d+[.)])\s+(?:\[[ xX]\]\s+)?/.exec(line);
    if (list) {
      if (line === list[0]) {
        return { value: value.slice(0, lineStart) + '\n' + value.slice(to),
          start: lineStart + 1, end: lineStart + 1 };
      }
      const ordered = /^\d/.test(list[2]);
      const marker = ordered ? `${parseInt(list[2]) + 1}${list[2].slice(-1)}` : list[2];
      const task = /\[[ xX]\]/.test(list[0]) ? '[ ] ' : '';
      return edit(`\n${list[1]}${marker} ${task}`);
    }
    const quote = /^(\s*(?:>\s*)+)/.exec(line);
    return edit(quote ? `\n${quote[1]}` : token === 'paragraph_open' ? '\n' : '\n\n');
  }
  return null;
}
