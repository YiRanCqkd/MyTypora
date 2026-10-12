import { nextTick, watch } from 'vue';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import { previewKeyEdit } from '../editor/preview-keyboard';
import { formatPreviewRange, previewPointToSource } from '../editor/preview-source-offset';
import { Transaction } from '@codemirror/state';
import { getPreviewCaretOffset } from '../editor/preview-editor-layout';

export function createParityPreview(domains: WorkbenchDomains) {
  const editor = domains.editor.foundation;
  const { store } = editor;
  const settings = domains.settings.foundation();
  function updatePreviewMode() {
    const host = editor.previewHost.value;
    if (!host) return;
    const line = editor.currentCursorLine.value;
    const blocks = [...host.children] as HTMLElement[];
    let active = blocks.find((node) => {
      const start = Number(node.dataset.sourceStart);
      const end = Number(node.dataset.sourceEnd);
      return start <= line && end >= line;
    });
    if (!active) active = blocks.find((node) => Number(node.dataset.sourceStart) >= line);
    for (const node of blocks) node.classList.toggle('focus-active', node === active);
    const textarea = domains.editor.presentation().previewInlineEditTextarea.value;
    if (settings.typewriterMode.value && !settings.sourceMode.value && textarea
      && document.activeElement === textarea) {
      const rect = textarea.getBoundingClientRect();
      const viewport = host.getBoundingClientRect();
      const currentLine = textarea.value.slice(0, getPreviewCaretOffset(textarea))
        .split('\n').length - 1;
      const height = parseFloat(getComputedStyle(textarea).lineHeight) || 22;
      const cursorY = rect.top + currentLine * height - textarea.scrollTop + height / 2;
      host.scrollTop += cursorY - (viewport.top + viewport.height / 2);
    }
  }
  watch(() => [domains.editor.presentation().renderedHtml.value,
    editor.currentCursorLine.value, settings.focusMode.value, settings.typewriterMode.value],
  () => { void nextTick(updatePreviewMode); });

  function previewKeyboard(event: KeyboardEvent) {
    const selectionInPreview = window.getSelection();
    const host = editor.previewHost.value;
    if (!settings.sourceMode.value && host && selectionInPreview?.rangeCount
      && !selectionInPreview.isCollapsed && (event.ctrlKey || event.metaKey)
      && !event.altKey && !event.shiftKey && ['b', 'i', 'k'].includes(event.key.toLowerCase())) {
      const anchor = previewPointToSource(host, store.content, selectionInPreview.anchorNode,
        selectionInPreview.anchorOffset);
      const head = previewPointToSource(host, store.content, selectionInPreview.focusNode,
        selectionInPreview.focusOffset);
      if (anchor !== null && head !== null && editor.runtime.editorView) {
        event.preventDefault();
        event.stopImmediatePropagation();
        editor.runtime.editorView.dispatch({ selection: { anchor, head } });
        const kind = event.key.toLowerCase() === 'b' ? 'bold' : event.key.toLowerCase() === 'i'
          ? 'italic' : 'link';
        const from = Math.min(anchor, head);
        const to = Math.max(anchor, head);
        if (store.content.slice(from, to).includes('\n')) {
          const insert = formatPreviewRange(store.content, from, to, kind);
          editor.runtime.editorView.dispatch({ changes: { from, to, insert },
            selection: { anchor: from, head: from + insert.length },
            annotations: Transaction.userEvent.of('input.format') });
        } else domains.editor.source().applySourceInlineShortcut(kind);
        selectionInPreview.removeAllRanges();
        return;
      }
    }
    const presentation = domains.editor.presentation();
    const textarea = presentation.previewInlineEditTextarea.value;
    if (event.target !== textarea || !textarea || event.isComposing) return;
    const selection = presentation.previewInlineEditSelection.value;
    if (!selection) return;
    const atStart = textarea.selectionStart === 0 && textarea.selectionEnd === 0;
    const atEnd = textarea.selectionStart === textarea.value.length
      && textarea.selectionEnd === textarea.value.length;
    if ((event.key === 'ArrowUp' && atStart) || (event.key === 'ArrowDown' && atEnd)) {
      const backwards = event.key === 'ArrowUp';
      const nodes = [...(editor.previewHost.value?.querySelectorAll<HTMLElement>(
        '[data-source-start][data-source-end]',
      ) || [])];
      const candidates = nodes.filter((node) => backwards
        ? Number(node.dataset.sourceEnd) < selection.startLine
        : Number(node.dataset.sourceStart) > selection.endLine);
      candidates.sort((a, b) => backwards
        ? Number(b.dataset.sourceStart) - Number(a.dataset.sourceStart)
        : Number(a.dataset.sourceStart) - Number(b.dataset.sourceStart));
      const target = candidates[0];
      if (target) {
        event.preventDefault();
        event.stopImmediatePropagation();
        domains.editor.preview().closePreviewInlineEditor();
        void nextTick(() => {
          const currentTarget = [...(editor.previewHost.value?.querySelectorAll<HTMLElement>(
            '[data-source-start]',
          ) || [])].find((node) => node.dataset.sourceStart === target.dataset.sourceStart);
          currentTarget?.dispatchEvent(new MouseEvent('click', {
            bubbles: true, cancelable: true, altKey: true,
          }));
          void nextTick(() => {
            const next = presentation.previewInlineEditTextarea.value;
            const position = backwards ? next?.value.length || 0 : 0;
            next?.setSelectionRange(position, position);
            updatePreviewMode();
          });
        });
      }
    }
    if (!event.ctrlKey && !event.metaKey && !event.altKey) {
      const edit = previewKeyEdit(textarea.value, textarea.selectionStart,
        textarea.selectionEnd, event.key, selection.token, event.shiftKey);
      if (!edit) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      presentation.previewInlineEditValue.value = edit.value;
      domains.editor.preview().flushPreviewInlineEditSync();
      domains.editor.events().updatePreviewInlineEditorSelection(edit.start, edit.end);
    }
  }

  function syncPreviewCursor(event: Event) {
    const textarea = domains.editor.presentation().previewInlineEditTextarea.value;
    const session = domains.editor.presentation().previewInlineEditSession.value;
    if (!textarea || event.target !== textarea || !session) return;
    const line = session.startLine + textarea.value.slice(0, getPreviewCaretOffset(textarea))
      .split('\n').length - 1;
    editor.currentCursorLine.value = line;
    store.setCursorHint(line);
    // 高度由统一布局模块计算，按键后再次收缩会干扰浏览器的选区滚动。
    void nextTick(updatePreviewMode);
  }

  return { previewKeyboard, syncPreviewCursor };
}
