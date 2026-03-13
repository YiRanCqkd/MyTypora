export interface HeadingItem {
  level: number;
  text: string;
  line: number;
}

export interface SearchLineItem {
  lineNo: number;
  line: string;
}

export interface MarkdownSyntaxIssue {
  line: number;
  rule: string;
  message: string;
}

export function extractHeadings(content: string): HeadingItem[] {
  const lines = content.split(/\r?\n/);
  const output: HeadingItem[] = [];
  let inFence = false;
  let fenceMarker: '`' | '~' | '' = '';
  let fenceLength = 0;

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const trimmed = line.trim();
    const fenceOpenMatch = /^(?<fence>`{3,}|~{3,})(?<rest>.*)$/.exec(trimmed);

    if (!inFence && fenceOpenMatch?.groups?.fence) {
      const marker = fenceOpenMatch.groups.fence[0];
      if (marker === '`' || marker === '~') {
        inFence = true;
        fenceMarker = marker;
        fenceLength = fenceOpenMatch.groups.fence.length;
      }
      continue;
    }

    if (inFence) {
      const closeRe = fenceMarker === '`' ? /^`+\s*$/ : /^~+\s*$/;
      const closeMatch = closeRe.exec(trimmed);
      if (closeMatch && closeMatch[0].trim().length >= fenceLength) {
        inFence = false;
        fenceMarker = '';
        fenceLength = 0;
      }
      continue;
    }

    if (/^(?: {4,}|\t+)/.test(line)) {
      continue;
    }

    const atxMatch = /^ {0,3}(#{1,6})[ \t]+(.+?)\s*#*\s*$/.exec(line);
    if (atxMatch) {
      output.push({
        level: atxMatch[1].length,
        text: atxMatch[2].trim(),
        line: index + 1,
      });
      continue;
    }

    if (index + 1 >= lines.length) {
      continue;
    }

    const nextLine = lines[index + 1];
    const setextMatch = /^ {0,3}(=+|-+)\s*$/.exec(nextLine);
    if (!setextMatch) {
      continue;
    }

    const text = line.trim();
    if (!text) {
      continue;
    }

    output.push({
      level: setextMatch[1][0] === '=' ? 1 : 2,
      text,
      line: index + 1,
    });
  }

  return output;
}

export function findMatchedLines(
  content: string,
  query: string,
  maxResults = 200,
): SearchLineItem[] {
  const normalizedQuery = query.trim().toLowerCase();
  if (!normalizedQuery) {
    return [];
  }

  return content
    .split(/\r?\n/)
    .map((line, index) => ({
      lineNo: index + 1,
      line,
    }))
    .filter((item) => item.line.toLowerCase().includes(normalizedQuery))
    .slice(0, maxResults);
}

export function findActiveHeadingLine(headings: HeadingItem[], cursorLine: number): number {
  let matchedLine = 0;

  for (const item of headings) {
    if (item.line <= cursorLine) {
      matchedLine = item.line;
      continue;
    }

    break;
  }

  return matchedLine;
}

export function hasMathSyntax(content: string): boolean {
  return content.includes('$');
}

interface FenceState {
  char: '`' | '~';
  len: number;
  line: number;
  info: string;
  content: string[];
}

interface ParsedFence {
  line: number;
  info: string;
  content: string[];
}

const TABLE_DELIMITER_RE = /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/;
const FOOTNOTE_DEF_RE = /^\[\^([^\]]+)\]:\s+/;
const FOOTNOTE_REF_RE = /\[\^([^\]]+)\]/g;
const LINK_INLINE_RE = /(!?)\[[^\]]*]\(([^)]*)\)/g;
const TASK_LIST_PREFIX_RE = /^\s*(?:[-+*]|\d+[.)])\s+\[/;
const TASK_LIST_VALID_RE = /^\s*(?:[-+*]|\d+[.)])\s+\[(?: |x|X)\](?:\s+.*)?$/;
const WINDOWS_PATH_INVALID_CHAR_RE = /[<>"|?*]/;
const URI_SCHEME_RE = /^[A-Za-z][A-Za-z\d+.-]*:/;
const MERMAID_ROOT_TOKENS = [
  'graph',
  'flowchart',
  'sequencediagram',
  'classdiagram',
  'statediagram',
  'erdiagram',
  'journey',
  'gantt',
  'pie',
  'mindmap',
  'timeline',
  'quadrantchart',
  'gitgraph',
  'requirement',
  'xychart-beta',
  'sankey-beta',
  'block-beta',
  'packet-beta',
  'kanban',
];
const MERMAID_ROOT_RE = new RegExp(`^(${MERMAID_ROOT_TOKENS.join('|')})\\b`, 'i');

function countTableColumns(line: string): number {
  let normalized = line.trim();
  if (!normalized.includes('|')) {
    return 0;
  }

  if (normalized.startsWith('|')) {
    normalized = normalized.slice(1);
  }

  if (normalized.endsWith('|')) {
    normalized = normalized.slice(0, -1);
  }

  if (!normalized) {
    return 0;
  }

  return normalized.split('|').length;
}

function hasOddInlineMathDelimiter(line: string): boolean {
  let count = 0;
  let cursor = 0;

  while (cursor < line.length) {
    const ch = line[cursor];
    if (ch === '\\') {
      cursor += 2;
      continue;
    }

    if (ch !== '$') {
      cursor += 1;
      continue;
    }

    if (line[cursor + 1] === '$') {
      cursor += 2;
      continue;
    }

    count += 1;
    cursor += 1;
  }

  return count % 2 === 1;
}

function extractInlineLinkTarget(rawTarget: string): {
  target: string;
  wrapped: boolean;
} {
  const trimmed = rawTarget.trim();
  if (!trimmed) {
    return {
      target: '',
      wrapped: false,
    };
  }

  if (trimmed.startsWith('<')) {
    const closeIndex = trimmed.indexOf('>');
    if (closeIndex <= 0) {
      return {
        target: '',
        wrapped: true,
      };
    }

    return {
      target: trimmed.slice(1, closeIndex).trim(),
      wrapped: true,
    };
  }

  const firstToken = trimmed.split(/\s+/)[0] || '';
  return {
    target: firstToken.trim(),
    wrapped: false,
  };
}

function validateInlineLinkTarget(target: string, wrapped: boolean): string {
  if (!target) {
    return '链接或图片地址为空。';
  }

  if (!wrapped && /\s/.test(target)) {
    return '链接或图片地址包含空格，建议使用 <> 包裹或进行转义。';
  }

  if (WINDOWS_PATH_INVALID_CHAR_RE.test(target)) {
    return '链接或图片地址包含 Windows 非法路径字符。';
  }

  if (!URI_SCHEME_RE.test(target)) {
    return '';
  }

  try {
    const parsed = new URL(target);
    if (!parsed.protocol) {
      return '链接地址协议无效。';
    }
  } catch {
    return '链接地址格式非法。';
  }

  return '';
}

function validateDiagramFence(
  fence: ParsedFence,
  pushIssue: (issue: MarkdownSyntaxIssue) => void,
) {
  const kind = fence.info.toLowerCase();
  if (kind !== 'mermaid' && kind !== 'sequence' && kind !== 'flow') {
    return;
  }

  const nonEmptyLines = fence.content
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (nonEmptyLines.length === 0) {
    pushIssue({
      line: fence.line,
      rule: `diagram-${kind}`,
      message: `${kind} 图表代码块为空。`,
    });
    return;
  }

  if (kind === 'mermaid') {
    const firstLine = nonEmptyLines[0].toLowerCase();
    if (!MERMAID_ROOT_RE.test(firstLine)) {
      pushIssue({
        line: fence.line,
        rule: 'diagram-mermaid',
        message: 'Mermaid 代码块缺少合法图类型声明（如 graph/flowchart/sequenceDiagram）。',
      });
    }
    return;
  }

  if (kind === 'sequence') {
    const hasArrow = nonEmptyLines.some((line) => /-{1,2}>{1,2}|={1,2}>{1,2}/.test(line));
    if (!hasArrow) {
      pushIssue({
        line: fence.line,
        rule: 'diagram-sequence',
        message: 'Sequence 代码块缺少有效箭头语句（如 A->>B: msg）。',
      });
    }
    return;
  }

  const hasNodeDefine = nonEmptyLines.some((line) => line.includes('=>'));
  const hasFlowEdge = nonEmptyLines.some((line) => line.includes('->'));
  if (!hasNodeDefine || !hasFlowEdge) {
    pushIssue({
      line: fence.line,
      rule: 'diagram-flow',
      message: 'Flow 代码块需同时包含节点定义（=>）与连线（->）。',
    });
  }
}

export function validateMarkdownSyntax(
  content: string,
  maxIssues = 80,
): MarkdownSyntaxIssue[] {
  const lines = content.split(/\r?\n/);
  const issues: MarkdownSyntaxIssue[] = [];
  const isCodeLine = new Array(lines.length).fill(false);

  const pushIssue = (issue: MarkdownSyntaxIssue) => {
    if (issues.length >= maxIssues) {
      return;
    }

    issues.push(issue);
  };

  const parsedFences: ParsedFence[] = [];
  let fence: FenceState | null = null;

  lines.forEach((line, index) => {
    const lineNo = index + 1;
    const trimmed = line.trim();

    if (!fence) {
      const openMatch = /^(?<fence>`{3,}|~{3,})(?<rest>.*)$/.exec(trimmed);
      if (!openMatch || !openMatch.groups) {
        return;
      }

      const marker = openMatch.groups.fence;
      const firstChar = marker[0];
      if (firstChar !== '`' && firstChar !== '~') {
        return;
      }

      fence = {
        char: firstChar,
        len: marker.length,
        line: lineNo,
        info: openMatch.groups.rest.trim().split(/\s+/)[0]?.toLowerCase() || '',
        content: [],
      };
      isCodeLine[index] = true;
      return;
    }

    isCodeLine[index] = true;

    if (fence.char === '`') {
      const closeMatch = /^`+\s*$/.exec(trimmed);
      if (closeMatch && closeMatch[0].trim().length >= fence.len) {
        parsedFences.push({
          line: fence.line,
          info: fence.info,
          content: [...fence.content],
        });
        fence = null;
      } else {
        fence.content.push(line);
      }
      return;
    }

    const closeMatch = /^~+\s*$/.exec(trimmed);
    if (closeMatch && closeMatch[0].trim().length >= fence.len) {
      parsedFences.push({
        line: fence.line,
        info: fence.info,
        content: [...fence.content],
      });
      fence = null;
      return;
    }

    fence.content.push(line);
  });

  const unclosedFenceLine = (fence as FenceState | null)?.line ?? 0;
  if (unclosedFenceLine > 0) {
    pushIssue({
      line: unclosedFenceLine,
      rule: 'fence',
      message: '代码块未闭合：缺少结束围栏。',
    });
  }

  const footnoteDefs = new Map<string, number>();
  const footnoteRefs: Array<{ key: string; line: number }> = [];
  let mathBlockOpenLine = 0;

  lines.forEach((line, index) => {
    if (isCodeLine[index]) {
      return;
    }

    const lineNo = index + 1;
    const trimmed = line.trim();

    if (trimmed === '$$') {
      if (mathBlockOpenLine === 0) {
        mathBlockOpenLine = lineNo;
      } else {
        mathBlockOpenLine = 0;
      }
      return;
    }

    if (mathBlockOpenLine > 0) {
      return;
    }

    if (/^#{1,6}\S/.test(trimmed)) {
      pushIssue({
        line: lineNo,
        rule: 'heading-space',
        message: '标题 # 后缺少空格，建议写成 "# 标题"。',
      });
    }

    if (hasOddInlineMathDelimiter(line)) {
      pushIssue({
        line: lineNo,
        rule: 'inline-math',
        message: '行内数学公式分隔符 "$" 未成对。',
      });
    }

    if (TASK_LIST_PREFIX_RE.test(line) && !TASK_LIST_VALID_RE.test(line)) {
      pushIssue({
        line: lineNo,
        rule: 'task-list',
        message: '任务列表语法非法，需使用 "- [ ] item" 或 "- [x] item"。',
      });
    }

    LINK_INLINE_RE.lastIndex = 0;
    let inlineLinkMatch: RegExpExecArray | null = null;
    while ((inlineLinkMatch = LINK_INLINE_RE.exec(line)) !== null) {
      const { target, wrapped } = extractInlineLinkTarget(inlineLinkMatch[2]);
      const message = validateInlineLinkTarget(target, wrapped);
      if (!message) {
        continue;
      }

      pushIssue({
        line: lineNo,
        rule: 'link-target',
        message,
      });
    }

    const defMatch = FOOTNOTE_DEF_RE.exec(trimmed);
    if (defMatch) {
      footnoteDefs.set(defMatch[1], lineNo);
      return;
    }

    FOOTNOTE_REF_RE.lastIndex = 0;
    let currentRef: RegExpExecArray | null = null;
    while ((currentRef = FOOTNOTE_REF_RE.exec(line)) !== null) {
      footnoteRefs.push({
        key: currentRef[1],
        line: lineNo,
      });
    }
  });

  if (mathBlockOpenLine > 0) {
    pushIssue({
      line: mathBlockOpenLine,
      rule: 'block-math',
      message: '块级数学公式未闭合：缺少结束 "$$"。',
    });
  }

  footnoteRefs.forEach((ref) => {
    if (footnoteDefs.has(ref.key)) {
      return;
    }

    pushIssue({
      line: ref.line,
      rule: 'footnote-ref',
      message: `脚注引用 [^${ref.key}] 未找到对应定义。`,
    });
  });

  footnoteDefs.forEach((lineNo, key) => {
    const hasRef = footnoteRefs.some((item) => item.key === key);
    if (hasRef) {
      return;
    }

    pushIssue({
      line: lineNo,
      rule: 'footnote-def',
      message: `脚注定义 [^${key}] 未被引用。`,
    });
  });

  for (let i = 0; i < lines.length - 1; i += 1) {
    if (isCodeLine[i] || isCodeLine[i + 1]) {
      continue;
    }

    const headerLine = lines[i];
    const delimiterLine = lines[i + 1];

    if (!headerLine.includes('|') || !TABLE_DELIMITER_RE.test(delimiterLine)) {
      continue;
    }

    const headerCols = countTableColumns(headerLine);
    const delimiterCols = countTableColumns(delimiterLine);

    if (headerCols > 0 && delimiterCols > 0 && headerCols !== delimiterCols) {
      pushIssue({
        line: i + 2,
        rule: 'table-header',
        message: '表格表头列数与分隔线列数不一致。',
      });
    }

    for (let row = i + 2; row < lines.length; row += 1) {
      if (isCodeLine[row]) {
        break;
      }

      const rowLine = lines[row];
      if (!rowLine.trim() || !rowLine.includes('|')) {
        break;
      }

      const rowCols = countTableColumns(rowLine);
      if (headerCols > 0 && rowCols > 0 && rowCols !== headerCols) {
        pushIssue({
          line: row + 1,
          rule: 'table-row',
          message: '表格数据列数与表头列数不一致。',
        });
      }
    }
  }

  parsedFences.forEach((item) => {
    validateDiagramFence(item, pushIssue);
  });

  return issues;
}
