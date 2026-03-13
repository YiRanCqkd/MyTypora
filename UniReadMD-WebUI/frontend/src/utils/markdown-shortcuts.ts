export type InlineShortcutKind = 'bold' | 'italic' | 'link';

export interface InlineShortcutPatch {
  insert: string;
  selectionStart: number;
  selectionEnd: number;
  changed: boolean;
}

export interface BlockIndentPatch {
  text: string;
  changed: boolean;
}

const INLINE_LINK_RE = /^\[([^\]]*)\]\(([^)]*)\)$/;
const LINK_TEXT_PLACEHOLDER = 'link';
const LINK_URL_PLACEHOLDER = 'https://';

function unwrapMarker(text: string, marker: string): string | null {
  if (text.length < marker.length * 2) {
    return null;
  }

  if (!text.startsWith(marker) || !text.endsWith(marker)) {
    return null;
  }

  return text.slice(marker.length, text.length - marker.length);
}

function buildMarkerShortcutPatch(selectionText: string, marker: string): InlineShortcutPatch {
  const unwrapped = unwrapMarker(selectionText, marker);
  if (unwrapped !== null) {
    return {
      insert: unwrapped,
      selectionStart: 0,
      selectionEnd: unwrapped.length,
      changed: true,
    };
  }

  const wrapped = `${marker}${selectionText}${marker}`;
  if (selectionText.length === 0) {
    return {
      insert: wrapped,
      selectionStart: marker.length,
      selectionEnd: marker.length,
      changed: true,
    };
  }

  return {
    insert: wrapped,
    selectionStart: marker.length,
    selectionEnd: marker.length + selectionText.length,
    changed: true,
  };
}

function buildLinkShortcutPatch(selectionText: string): InlineShortcutPatch {
  const trimmedSelection = selectionText.trim();
  const matchedLink = INLINE_LINK_RE.exec(trimmedSelection);
  if (matchedLink) {
    const linkText = matchedLink[1] || '';
    return {
      insert: linkText,
      selectionStart: 0,
      selectionEnd: linkText.length,
      changed: true,
    };
  }

  if (selectionText.length > 0) {
    const insert = `[${selectionText}](${LINK_URL_PLACEHOLDER})`;
    const urlStart = selectionText.length + 3;
    return {
      insert,
      selectionStart: urlStart,
      selectionEnd: urlStart + LINK_URL_PLACEHOLDER.length,
      changed: true,
    };
  }

  const insert = `[${LINK_TEXT_PLACEHOLDER}](${LINK_URL_PLACEHOLDER})`;
  return {
    insert,
    selectionStart: 1,
    selectionEnd: 1 + LINK_TEXT_PLACEHOLDER.length,
    changed: true,
  };
}

export function buildInlineShortcutPatch(
  selectionText: string,
  kind: InlineShortcutKind,
): InlineShortcutPatch {
  if (kind === 'bold') {
    return buildMarkerShortcutPatch(selectionText, '**');
  }

  if (kind === 'italic') {
    return buildMarkerShortcutPatch(selectionText, '*');
  }

  if (kind === 'link') {
    return buildLinkShortcutPatch(selectionText);
  }

  return {
    insert: selectionText,
    selectionStart: 0,
    selectionEnd: selectionText.length,
    changed: false,
  };
}

function outdentSingleLine(line: string): string {
  if (line.startsWith('\t')) {
    return line.slice(1);
  }

  if (line.startsWith('  ')) {
    return line.slice(2);
  }

  if (line.startsWith(' ')) {
    return line.slice(1);
  }

  return line;
}

export function indentMarkdownBlockLines(
  blockText: string,
  options?: {
    outdent?: boolean;
    indentUnit?: string;
  },
): BlockIndentPatch {
  const outdent = Boolean(options?.outdent);
  const indentUnit = options?.indentUnit || '  ';

  if (!blockText) {
    return {
      text: blockText,
      changed: false,
    };
  }

  const lines = blockText.split('\n');
  const nextLines = lines.map((line) => {
    if (outdent) {
      return outdentSingleLine(line);
    }

    return `${indentUnit}${line}`;
  });
  const nextText = nextLines.join('\n');
  return {
    text: nextText,
    changed: nextText !== blockText,
  };
}
