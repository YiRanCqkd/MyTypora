import { nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import { Transaction } from '@codemirror/state';
import { isolateHistory, redo, undo } from '@codemirror/commands';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import type { PreviewTableSelection } from '../editor/editor-contracts';
import { parseMarkdownTable, applyPreviewTableActionToModel,
  type PreviewTableAction } from '../../utils/preview-table';
import { navigateVisualTable, writeVisualTableCell } from '../editor/visual-table';
import { sanitizeMarkdownHtml } from '../markdown/safe-html';
import { applyVisualPrefix, describeVisualBlock, nextVisualPrefix, visualHtmlToMarkdown,
  type VisualBlock } from '../editor/visual-markdown';

interface VisualSession {
  node: HTMLElement;
  from: number;
  to: number;
  line: number;
  endLine: number;
  original: string;
  source: string;
  html: string;
  block: VisualBlock;
  table?: PreviewTableSelection;
}

export function useVisualPreviewEditing(domains: WorkbenchDomains) {
  const { runtime, store, previewHost, currentCursorLine } = domains.editor.foundation;
  const presentation = domains.editor.presentation();
  const { sourceMode } = domains.settings.foundation();
  const linkVisible = ref(false);
  const linkUrl = ref('');
  const linkError = ref('');
  let session: VisualSession | null = null;
  let composing = false;
  let writing = false;
  let timer = 0;
  let linkRange: Range | null = null;
  let plainPasteUntil = 0;

  function eligible(target: EventTarget | null) {
    const element = target instanceof Element ? target : null;
    const node = element?.closest<HTMLElement>('p, h1, h2, h3, h4, h5, h6, li, th, td');
    const cell = node?.matches('th, td');
    if (!node || !previewHost.value?.contains(node) || sourceMode.value
      || !(cell ? domains.editor.table().resolvePreviewTableCellContext(node)
        : node.matches('[data-source-start][data-source-end]'))
      || node.closest(cell ? 'pre, .mermaid-block, .footnotes'
        : 'pre, table, .mermaid-block, .footnotes')
      || node.querySelector('img, svg, input, ul, ol, .katex, mjx-container, .footnote-ref')
      || [...node.querySelectorAll('*')].some((child) => ![
        'BR', 'STRONG', 'B', 'EM', 'I', 'S', 'DEL', 'STRIKE', 'CODE', 'A', 'MARK', 'SUB', 'SUP',
      ].includes(child.tagName))) return null;
    return node;
  }

  function commit() {
    window.clearTimeout(timer);
    const view = runtime.editorView;
    const active = session;
    if (!active || !view || active.source !== view.state.doc.toString()) return;
    if (active.html === active.node.innerHTML) return;
    const value = visualHtmlToMarkdown(active.node.innerHTML,
      !active.table && active.block.kind === 'paragraph');
    const insert = active.table
      ? writeVisualTableCell(
        view.state.doc.sliceString(active.from, active.to), active.table, value,
      )
      : applyVisualPrefix(value, active.block);
    if (insert === null) return;
    if (view.state.doc.sliceString(active.from, active.to) === insert) return;
    writing = true;
    try {
      view.dispatch({ changes: { from: active.from, to: active.to, insert },
        selection: { anchor: active.from + insert.length },
        annotations: Transaction.userEvent.of('input.type') });
      active.to = active.from + insert.length;
      active.source = view.state.doc.toString();
      active.html = active.node.innerHTML;
    } finally { writing = false; }
  }

  function close(save = true) {
    window.clearTimeout(timer);
    if (!session) return;
    if (save) commit();
    session.node.removeAttribute('contenteditable');
    session.node.removeAttribute('role');
    session.node.removeAttribute('aria-label');
    session.node.classList.remove('preview-visual-editing');
    session = null;
    composing = false;
    linkVisible.value = false;
    linkRange = null;
    // 即使撤销回原源码，也强制更新 v-html，清除浏览器暂存的编辑 DOM。
    presentation.previewRenderRevision.value += 1;
    presentation.previewVisualEditing.value = false;
  }

  function placeCaret(node: HTMLElement, end = false) {
    node.focus({ preventScroll: true });
    const range = document.createRange();
    range.selectNodeContents(node);
    range.collapse(!end);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }

  function open(node: HTMLElement, point?: { x: number; y: number }, end = false) {
    const view = runtime.editorView;
    if (!view || !node.isConnected) return false;
    const table = node.matches('th, td')
      ? domains.editor.table().resolvePreviewTableCellContext(node) : null;
    const line = table?.startLine || Number(node.dataset.sourceStart);
    const endLine = table?.endLine || Number(node.dataset.sourceEnd);
    if (!line || !endLine || endLine > view.state.doc.lines) return false;
    const from = view.state.doc.line(line).from;
    const to = view.state.doc.line(endLine).to;
    const original = view.state.doc.sliceString(from, to);
    const parsed = table ? parseMarkdownTable(original) : null;
    const cellValue = table && parsed ? (table.rowIndex === 0 ? parsed.header[table.colIndex]
      : parsed.rows[table.rowIndex - 1]?.[table.colIndex]) : null;
    // 无法取得单元格值时交还源码入口，不能把缺失值当成可编辑内容。
    if (table && cellValue == null) return false;
    // 表格内换行只接受无属性的 br，其他 HTML 仍保留源码编辑入口。
    const described = describeVisualBlock(table && cellValue != null
      ? cellValue.replace(/<br\s*\/?>/gi, '\n') : original);
    const block: VisualBlock | null = table && described
      ? { prefix: '', kind: 'paragraph' } : described;
    if (!block) return false;
    domains.editor.preview().closePreviewInlineEditor();
    domains.editor.table().closePreviewTableCellEditorAndCommit();
    domains.editor.preview().closePreviewTableMenu();
    view.dispatch({ annotations: isolateHistory.of('before') });
    // 先求值，确保冻结的是刚提交的源码对应的预览。
    void presentation.renderedHtml.value;
    presentation.previewVisualEditing.value = true;
    session = { node, from, to, line, endLine, original, source: view.state.doc.toString(),
      html: node.innerHTML, block, ...(table ? { table: {
        startLine: line, endLine, rowIndex: table.rowIndex, colIndex: table.colIndex,
      } } : {}) };
    currentCursorLine.value = line;
    store.setCursorHint(line);
    node.setAttribute('contenteditable', 'true');
    node.setAttribute('role', 'textbox');
    node.setAttribute('aria-label', table ? '表格单元格原位编辑' : '原位可视编辑');
    node.classList.add('preview-visual-editing');
    const caret = point ? document.caretRangeFromPoint?.(point.x, point.y) : null;
    if (caret && node.contains(caret.startContainer)) {
      node.focus({ preventScroll: true });
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(caret);
    } else placeCaret(node, end);
    return true;
  }

  async function reopenTable(selection: PreviewTableSelection, point?: { x: number; y: number }) {
    await nextTick();
    const table = previewHost.value?.querySelector<HTMLElement>(
      `table[data-source-start='${selection.startLine}']`,
    );
    const node = table?.querySelectorAll('tr')[selection.rowIndex]
      ?.querySelectorAll<HTMLElement>('th, td')[selection.colIndex];
    if (!node) return;
    domains.editor.events().openPreviewTableCellEditor({ ...selection,
      endLine: Number(table!.dataset.sourceEnd) });
    if (session?.node === node && point) {
      const caret = document.caretRangeFromPoint?.(point.x, point.y);
      if (caret && node.contains(caret.startContainer)) {
        window.getSelection()?.removeAllRanges();
        window.getSelection()?.addRange(caret);
      }
    }
  }

  function navigateTable(direction: 'next' | 'previous' | 'down') {
    const active = session;
    const view = runtime.editorView;
    if (!active?.table || !view) return;
    commit();
    const result = navigateVisualTable(view.state.doc.sliceString(active.from, active.to),
      active.table, direction);
    close();
    if (!result) return;
    if (view.state.doc.sliceString(active.from, active.to) !== result.content) {
      view.dispatch({ changes: { from: active.from, to: active.to, insert: result.content },
        annotations: Transaction.userEvent.of('input.table-row') });
    }
    void reopenTable(result.selection);
  }

  function tableStructureAction(action: PreviewTableAction) {
    const active = session;
    if (!active?.table) return;
    commit();
    close();
    const result = domains.editor.table().mutatePreviewTable(active.table, (table, selection) => {
      const actionResult = applyPreviewTableActionToModel(table, selection, action);
      if (actionResult.errorMessage) {
        domains.editor.foundation.errorMessage.value = actionResult.errorMessage;
      }
      return actionResult.changed;
    });
    if (result) void reopenTable(domains.editor.tableFeature().getPreviewTableSelectionAfterAction(
      action, result.selection, result.table,
    ));
  }

  async function reopen(line: number, end = false) {
    await nextTick();
    let node = previewHost.value?.querySelector<HTMLElement>(
      `p[data-source-start='${line}'], h1[data-source-start='${line}'], `
      + `h2[data-source-start='${line}'], h3[data-source-start='${line}'], `
      + `h4[data-source-start='${line}'], h5[data-source-start='${line}'], `
      + `h6[data-source-start='${line}'], li[data-source-start='${line}']`,
    );
    const view = runtime.editorView;
    if (!node && previewHost.value && view && line <= view.state.doc.lines
      && !view.state.doc.line(line).text.trim()) {
      node = document.createElement('p');
      node.dataset.sourceStart = String(line);
      node.dataset.sourceEnd = String(line);
      node.dataset.sourceToken = 'paragraph_open';
      node.innerHTML = '<br>';
      const next = [...previewHost.value.children].find((child) =>
        Number((child as HTMLElement).dataset.sourceStart) > line);
      previewHost.value.insertBefore(node, next || null);
    }
    if (node && eligible(node)) open(node, undefined, end);
  }

  function click(event: MouseEvent) {
    if (event.altKey || event.ctrlKey || event.metaKey || event.button !== 0) {
      return;
    }
    const node = eligible(event.target);
    if (!node) return;
    if (session?.node === node) { event.stopImmediatePropagation(); return; }
    if (window.getSelection()?.isCollapsed === false) return;
    const point = { x: event.clientX, y: event.clientY };
    if (session) {
      const table = node.matches('th, td')
        ? domains.editor.table().resolvePreviewTableCellContext(node) : null;
      let line = table?.startLine || Number(node.dataset.sourceStart);
      commit();
      if (line > session.endLine && runtime.editorView) {
        line += runtime.editorView.state.doc.lineAt(session.to).number - session.endLine;
      }
      close();
      event.preventDefault();
      event.stopImmediatePropagation();
      if (table) void reopenTable({ ...table, startLine: line,
        endLine: table.endLine + line - table.startLine }, point);
      else void reopen(line);
    } else if (open(node, point)) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
  }

  function input(event: Event) {
    if (!session?.node.contains(event.target as Node) || composing) return;
    window.clearTimeout(timer);
    timer = window.setTimeout(commit, 100);
  }
  function compositionStart(event: Event) {
    if (session?.node.contains(event.target as Node)) composing = true;
  }
  function compositionEnd(event: Event) {
    if (!session?.node.contains(event.target as Node)) return;
    composing = false;
    commit();
  }

  function split() {
    const active = session;
    const view = runtime.editorView;
    const selection = window.getSelection();
    if (!active || !view || !selection?.rangeCount) return;
    const selected = selection.getRangeAt(0);
    if (!active.node.contains(selected.startContainer)
      || !active.node.contains(selected.endContainer)) return;
    const before = selected.cloneRange();
    before.selectNodeContents(active.node);
    before.setEnd(selected.startContainer, selected.startOffset);
    const after = selected.cloneRange();
    after.selectNodeContents(active.node);
    after.setStart(selected.endContainer, selected.endOffset);
    function serialize(range: Range) {
      const holder = document.createElement('div');
      holder.append(range.cloneContents());
      return visualHtmlToMarkdown(holder.innerHTML);
    }
    const left = serialize(before);
    const right = serialize(after);
    const prefix = nextVisualPrefix(active.block);
    const separator = active.block.kind === 'list' || active.block.kind === 'quote' ? '\n' : '\n\n';
    const emptyList = !left && !right && active.block.kind === 'list';
    const start = emptyList ? '\n' : applyVisualPrefix(left, active.block) + separator;
    const insert = emptyList ? start : start + prefix + right;
    const from = active.from;
    commit();
    const to = active.to;
    close(false);
    view.dispatch({ changes: { from, to, insert },
      selection: { anchor: Math.min(from + insert.length, from + start.length + prefix.length) },
      annotations: Transaction.userEvent.of('input.split') });
    const line = view.state.doc.lineAt(Math.min(view.state.doc.length,
      from + start.length)).number;
    void reopen(line);
  }

  function applyLink() {
    const url = linkUrl.value.trim();
    if (!session || !linkRange) return;
    if (!url || [...url].some((char) => char.charCodeAt(0) < 32 || char.charCodeAt(0) === 127)
      || /^[a-z][\w+.-]*:/i.test(url) && !/^(?:https?:|mailto:|tel:)/i.test(url)) {
      linkError.value = '请输入网页地址、邮件地址或文档相对路径。';
      return;
    }
    session.node.focus({ preventScroll: true });
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(linkRange);
    document.execCommand('createLink', false, url);
    linkVisible.value = false;
    commit();
  }

  function keydown(event: KeyboardEvent) {
    if (!session || composing || event.isComposing
      || !session.node.contains(event.target as Node)) return;
    const modifier = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();
    if (session.table) {
      let action: PreviewTableAction | undefined;
      if (modifier && event.shiftKey && key === 'enter') {
        action = event.altKey ? 'insert-col-right' : 'insert-row-below';
      } else if (modifier && event.shiftKey && ['delete', 'backspace'].includes(key)) {
        action = event.altKey ? 'delete-col' : 'delete-row';
      } else if (modifier && event.altKey) {
        action = ({ arrowup: 'move-row-up', arrowdown: 'move-row-down',
          arrowleft: 'move-col-left', arrowright: 'move-col-right',
        } as Record<string, PreviewTableAction>)[key];
      }
      if (action || key === 'tab' || key === 'enter') {
        event.preventDefault();
        event.stopImmediatePropagation();
        if (action) tableStructureAction(action);
        else if (key === 'tab') navigateTable(event.shiftKey ? 'previous' : 'next');
        else if (modifier) close();
        else if (event.shiftKey) { document.execCommand('insertLineBreak'); commit(); }
        else navigateTable('down');
        return;
      }
      // 单元格边界不能沿用正文的跨段合并和退格逻辑。
      if (!modifier && key !== 'escape') return;
    }
    if (modifier && key === 'v') plainPasteUntil = event.shiftKey ? Date.now() + 1000 : 0;
    if (!modifier && !event.altKey && !event.shiftKey
      && ['Backspace', 'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) {
      const selection = window.getSelection();
      if (selection?.isCollapsed && selection.rangeCount) {
        const caret = selection.getRangeAt(0);
        const before = caret.cloneRange();
        before.selectNodeContents(session.node);
        before.setEnd(caret.startContainer, caret.startOffset);
        // 换行可能嵌在加粗等格式节点内，空文本和格式容器不能遮住前方的 br。
        const walker = document.createTreeWalker(before.cloneContents(),
          NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
        let preceding: Node | null = null;
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          if (node instanceof HTMLBRElement
            || node.nodeType === Node.TEXT_NODE && node.textContent) preceding = node;
        }
        if (event.key === 'Backspace' && preceding instanceof HTMLBRElement) {
          // 删掉行间换行前先提交已有文字，让一次撤销只恢复这次合并。
          commit();
          runtime.editorView?.dispatch({ annotations: isolateHistory.of('before') });
          // 此时由浏览器删除段内换行，不能继续按纯文本长度判断跨段边界。
          return;
        }
        const atStart = preceding === null;
        const atEnd = before.toString().length === session.node.textContent?.length;
        const backwards = ['Backspace', 'ArrowUp', 'ArrowLeft'].includes(event.key);
        if ((backwards && atStart) || (!backwards && atEnd)) {
          const nodes = [...(previewHost.value?.querySelectorAll<HTMLElement>(
            'p[data-source-start], h1, h2, h3, h4, h5, h6, li[data-source-start]',
          ) || [])].filter((node) => eligible(node));
          const index = nodes.indexOf(session.node);
          const target = nodes[index + (backwards ? -1 : 1)];
          if (target) {
            if (event.key === 'Backspace') {
              if (session.block.kind !== 'paragraph' || target.tagName !== 'P'
                || target.parentElement !== session.node.parentElement) return;
              const view = runtime.editorView!;
              const line = Number(target.dataset.sourceStart);
              const from = view.state.doc.line(line).from;
              const previousEnd = view.state.doc.line(Number(target.dataset.sourceEnd)).to;
              if (view.state.doc.sliceString(previousEnd, session.from).trim()) return;
              const previous = view.state.doc.sliceString(from, previousEnd);
              if (describeVisualBlock(previous)?.kind !== 'paragraph') return;
              const textLength = target.textContent?.length || 0;
              commit();
              const insert = previous + view.state.doc.sliceString(session.from, session.to);
              const to = session.to;
              close(false);
              view.dispatch({ changes: { from, to, insert },
                selection: { anchor: from + previous.length },
                annotations: Transaction.userEvent.of('input.join') });
              void reopen(line).then(() => {
                if (!session) return;
                const walker = document.createTreeWalker(session.node, NodeFilter.SHOW_TEXT);
                let remaining = textLength;
                for (let text = walker.nextNode(); text; text = walker.nextNode()) {
                  if (remaining <= (text.textContent?.length || 0)) {
                    window.getSelection()?.collapse(text, remaining);
                    break;
                  }
                  remaining -= text.textContent?.length || 0;
                }
              });
            } else {
              const line = Number(target.dataset.sourceStart);
              close();
              void reopen(line, backwards);
            }
            event.preventDefault();
            event.stopImmediatePropagation();
            return;
          }
        }
      }
    }
    if (modifier && ['b', 'i', 'k', 'z', 'y'].includes(key) && !event.altKey) {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (key === 'z' || key === 'y') {
        const line = session.line;
        const table = session.table;
        close();
        const view = runtime.editorView;
        if (view) (key === 'y' || event.shiftKey ? redo : undo)(view);
        if (table) void reopenTable(table);
        else void reopen(Math.min(line, view?.state.doc.lines || 1), true);
      } else if (key === 'k') {
        const selection = window.getSelection();
        if (!selection?.rangeCount || selection.isCollapsed) return;
        linkRange = selection.getRangeAt(0).cloneRange();
        linkUrl.value = '';
        linkError.value = '';
        linkVisible.value = true;
        void nextTick(() => document.querySelector<HTMLInputElement>('.visual-link-dialog input')
          ?.focus());
      } else {
        document.execCommand(key === 'b' ? 'bold' : 'italic');
        commit();
      }
    } else if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (modifier) close();
      else if (session.block.kind === 'quote' || session.block.kind === 'paragraph') {
        document.execCommand('insertLineBreak');
        commit();
      } else split();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      event.stopImmediatePropagation();
      const { from, to, original } = session;
      close(false);
      runtime.editorView?.dispatch({ changes: { from, to, insert: original },
        annotations: Transaction.addToHistory.of(false) });
    }
  }

  function paste(event: ClipboardEvent) {
    if (!session?.node.contains(event.target as Node) || !event.clipboardData) return;
    if ([...event.clipboardData.items].some((item) => item.type.startsWith('image/'))) {
      close();
      return;
    }
    event.preventDefault();
    event.stopImmediatePropagation();
    const html = event.clipboardData.getData('text/html');
    const safe = html ? sanitizeMarkdownHtml(html, { trustedAttributes: false }) : '';
    const holder = document.createElement('div');
    holder.innerHTML = safe;
    const supported = safe && [...holder.querySelectorAll('*')].every((node) => [
      'P', 'DIV', 'BR', 'SPAN', 'STRONG', 'B', 'EM', 'I', 'S', 'DEL', 'CODE', 'A', 'MARK',
      'SUB', 'SUP',
    ].includes(node.tagName));
    if (supported && Date.now() >= plainPasteUntil) {
      document.execCommand('insertHTML', false, safe);
    }
    else document.execCommand('insertText', false,
      event.clipboardData.getData('text/plain') || holder.textContent || '');
    commit();
  }
  function pointerdown(event: Event) {
    const target = event.target;
    if (!(target instanceof Element) || target.closest('.visual-link-dialog')) return;
    if (!previewHost.value?.contains(target)) close();
  }

  let caretFrame = 0;
  function revealCaret() {
    cancelAnimationFrame(caretFrame);
    caretFrame = requestAnimationFrame(() => {
      const host = previewHost.value;
      const selection = window.getSelection();
      if (!session || !host || !selection?.rangeCount || composing
        || !session.node.contains(selection.focusNode)) return;
      const caret = selection.getRangeAt(0).cloneRange();
      caret.collapse(false);
      const rect = caret.getBoundingClientRect();
      const viewport = host.getBoundingClientRect();
      if (!rect.height) return;
      if (domains.settings.foundation().typewriterMode.value) {
        host.scrollTop += rect.top + rect.height / 2 - (viewport.top + viewport.height / 2);
      } else if (rect.bottom > viewport.bottom - 20) {
        host.scrollTop += rect.bottom - viewport.bottom + 20;
      } else if (rect.top < viewport.top + 12) host.scrollTop -= viewport.top + 12 - rect.top;
    });
  }

  runtime.flushVisualEdits = (selection) => {
    commit();
    const view = runtime.editorView;
    let mapped = selection;
    if (selection && session && view) {
      // 冻结预览仍使用会话开始时的行号；已防抖提交和待提交的改动都要换算。
      const endLine = view.state.doc.lineAt(session.to).number;
      const delta = endLine - session.endLine;
      if (selection.startLine > session.endLine) {
        mapped = { ...selection,
          startLine: selection.startLine + delta, endLine: selection.endLine + delta };
      } else if (session.table && selection.startLine === session.line) {
        mapped = { ...selection, endLine };
      }
    }
    close(false);
    return mapped;
  };
  runtime.openVisualTableCell = (selection) => {
    const node = domains.editor.table().findPreviewTableCellNode(selection);
    if (!node || !eligible(node)) return false;
    if (session?.node === node) return true;
    if (session) { close(); void reopenTable(selection); return true; }
    return open(node);
  };
  watch(() => store.documentRevision, () => close(false), { flush: 'sync' });
  watch(() => store.content, () => { if (!writing && session) close(false); }, { flush: 'sync' });
  watch(sourceMode, () => close(), { flush: 'sync' });
  onMounted(() => {
    document.addEventListener('click', click, true);
    document.addEventListener('input', input, true);
    document.addEventListener('compositionstart', compositionStart, true);
    document.addEventListener('compositionend', compositionEnd, true);
    document.addEventListener('keydown', keydown, true);
    document.addEventListener('paste', paste, true);
    document.addEventListener('pointerdown', pointerdown, true);
    document.addEventListener('input', revealCaret);
    document.addEventListener('keyup', revealCaret);
  });
  onBeforeUnmount(() => {
    close();
    cancelAnimationFrame(caretFrame);
    runtime.flushVisualEdits = undefined;
    runtime.openVisualTableCell = undefined;
    document.removeEventListener('click', click, true);
    document.removeEventListener('input', input, true);
    document.removeEventListener('compositionstart', compositionStart, true);
    document.removeEventListener('compositionend', compositionEnd, true);
    document.removeEventListener('keydown', keydown, true);
    document.removeEventListener('paste', paste, true);
    document.removeEventListener('pointerdown', pointerdown, true);
    document.removeEventListener('input', revealCaret);
    document.removeEventListener('keyup', revealCaret);
  });
  return { linkVisible, linkUrl, linkError, applyLink, cancelLink: () => {
    linkVisible.value = false;
    session?.node.focus({ preventScroll: true });
    if (linkRange) {
      window.getSelection()?.removeAllRanges();
      window.getSelection()?.addRange(linkRange);
    }
  } };
}
