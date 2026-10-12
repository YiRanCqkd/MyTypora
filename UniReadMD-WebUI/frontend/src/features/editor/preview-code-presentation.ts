import { highlightCode } from '../markdown/code-highlight';

function escapeCode(value: string) {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function highlightPreviewCode(source: string, token: string) {
  if (token !== 'fence' && token !== 'code_block') return '';
  const lines = source.split('\n');
  const opening = token === 'fence' ? /^ {0,3}(`{3,}|~{3,})(.*)$/.exec(lines[0] || '') : null;
  if (!opening) return escapeCode(source) + '\u200b';
  const language = opening[2]!.trim().split(/\s+/)[0]!.toLowerCase();
  const last = lines[lines.length - 1] || '';
  const closing = /^ {0,3}(`{3,}|~{3,})\s*$/.exec(last);
  const closed = lines.length > 1 && closing?.[1]?.[0] === opening[1]![0]
    && closing[1].length >= opening[1]!.length;
  const code = lines.slice(1, closed ? -1 : undefined).join('\n');
  const body = highlightCode(code, language) || escapeCode(code);
  const marker = (line: string) => `<span class='preview-code-fence'>${escapeCode(line)}</span>`;
  return marker(lines[0]!) + (lines.length > 1 ? '\n' + body : '')
    + (closed ? '\n' + marker(last) : '') + '\u200b';
}

export function syncPreviewCodeScroll(event: Event) {
  const textarea = event.target;
  if (!(textarea instanceof HTMLTextAreaElement)) return;
  const highlight = textarea.previousElementSibling;
  if (!(highlight instanceof HTMLElement)
    || !highlight.classList.contains('preview-code-highlight')) return;
  highlight.scrollLeft = textarea.scrollLeft;
  highlight.scrollTop = textarea.scrollTop;
}
