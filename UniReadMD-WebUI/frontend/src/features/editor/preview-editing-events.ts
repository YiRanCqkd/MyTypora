import type { PreviewTableSelection } from './editor-contracts';
import { sourceOffsetForRenderedText } from './preview-source-offset';
import { findInlineFragment } from './inline-fragment';
import { nextTick } from 'vue';
import { isolateHistory } from '@codemirror/commands';
import type { useTableCellEditing } from './table-cell-editing';
import type { usePreviewEditingSession } from './preview-editing-session';
import type { createTableMutations } from './table-mutations';
import type { usePreviewInteractions } from './preview-interactions';
import type { useCursorPosition } from './cursor-position';
import type { useApplicationStatus } from './application-status';
import type { createPreviewSelection } from './preview-selection';
import type { usePreviewEditingState } from './preview-editing-state';
import type { useViewOptions } from '../settings/view-options';
import type { useWorkbenchStore } from './editor-store';
import type { createWorkbenchRuntime } from './workbench-runtime';
import type { createPreviewEditingRanges } from './preview-editing-ranges';
type UsePreviewEditingEventsDependencies =
  ReturnType<typeof useTableCellEditing> &
  ReturnType<typeof usePreviewEditingSession> &
  ReturnType<typeof createTableMutations> &
  ReturnType<typeof usePreviewInteractions> &
  ReturnType<typeof useCursorPosition> &
  ReturnType<typeof useApplicationStatus> &
  ReturnType<typeof createPreviewSelection> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof useViewOptions> &
  ReturnType<typeof useWorkbenchStore> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof createPreviewEditingRanges>;

export function usePreviewEditingEvents(input: UsePreviewEditingEventsDependencies) {
  const { buildPreviewTableCellEditStyle,closePreviewInlineEditor,
    closePreviewTableCellEditorAndCommit,closePreviewTableMenu,closeSourceContextMenu,
    currentCursorLine,errorMessage,findPreviewTableCellNode,findSourceMappedNode,
    flushPreviewInlineEditSync,focusPreviewTableCellEditor,parseInlineEditSelection,
    previewInlineEditComposing,previewInlineEditSelection,previewInlineEditSession,
    previewInlineEditTextarea,previewInlineEditValue,previewInlineEditVisible,
    previewTableCellEditSelection,previewTableCellEditStyle,previewTableCellEditValue,
    previewTableCellEditVisible,previewTableSelection,queuePreviewInlineEditSync,
    readPreviewTableCellValue,resolvePreviewTableCellContext,resolvePreviewTaskItemContext,
    sourceMode,store,togglePreviewTaskItem,runtime,getEditorRangeByLines,
    buildPreviewInlineEditStyle,refreshPreviewInlineEditorPosition,revealPreviewInlineEditorCaret,
  } = input;

  function openPreviewTableCellEditor(selection: PreviewTableSelection) {
    closePreviewTableCellEditorAndCommit();

    const cell = findPreviewTableCellNode(selection);
    if (!cell) {
      return;
    }

    const nextStyle = buildPreviewTableCellEditStyle(cell);
    if (!nextStyle) {
      return;
    }

    closePreviewInlineEditor();
    closePreviewTableMenu();
    closeSourceContextMenu();
    previewTableSelection.value = {
      ...selection,
    };
    previewTableCellEditSelection.value = {
      ...selection,
    };
    previewTableCellEditStyle.value = nextStyle;
    previewTableCellEditValue.value = readPreviewTableCellValue(selection);
    previewTableCellEditVisible.value = true;
    errorMessage.value = '';
    focusPreviewTableCellEditor();
  }

  function handlePreviewClick(event: MouseEvent) {
    runtime.flushVisualEdits?.();
    if (sourceMode.value) {
      return;
    }
    // 鼠标拖选后 click 不能把跨块选区替换成编辑框。
    if (window.getSelection()?.isCollapsed === false) return;

    const target = event.target instanceof HTMLElement ? event.target : null;
    const taskItemContext = resolvePreviewTaskItemContext(target);
    if (taskItemContext) {
      event.preventDefault();
      event.stopPropagation();
      closePreviewTableMenu();
      closePreviewTableCellEditorAndCommit();
      closeSourceContextMenu();
      closePreviewInlineEditor();
      void togglePreviewTaskItem(taskItemContext);
      return;
    }

    const tableContext = resolvePreviewTableCellContext(target);
    if (tableContext) {
      event.preventDefault();
      openPreviewTableCellEditor({
        startLine: tableContext.startLine,
        endLine: tableContext.endLine,
        rowIndex: tableContext.rowIndex,
        colIndex: tableContext.colIndex,
      });
      return;
    }

    const anchor = target?.closest('a[href]');
    if (anchor && (event.ctrlKey || event.metaKey)) {
      return;
    }

    if (anchor) {
      event.preventDefault();
    }

    const node = findSourceMappedNode(event.target);
    if (!node) {
      if (previewInlineEditVisible.value) {
        closePreviewInlineEditor();
      }
      return;
    }

    const selection = parseInlineEditSelection(node);
    if (!selection) {
      return;
    }

    currentCursorLine.value = selection.startLine;
    store.setCursorHint(selection.startLine);

    const currentSelection = previewInlineEditSelection.value;
    const isSameInlineBlock = currentSelection
      && currentSelection.startLine === selection.startLine
      && currentSelection.endLine === selection.endLine
      && currentSelection.token === selection.token;
    if (previewInlineEditVisible.value && !isSameInlineBlock) {
      closePreviewInlineEditor();
    }
    if (isSameInlineBlock && previewInlineEditVisible.value) return;
    const range = getEditorRangeByLines(selection.startLine, selection.endLine);
    if (!range || !runtime.editorView) return;
    const originalText = runtime.editorView.state.doc.sliceString(range.from, range.to);
    const caret = document.caretRangeFromPoint?.(event.clientX, event.clientY);
    let offset = originalText.length;
    if (caret && node.contains(caret.startContainer)) {
      const prefix = document.createRange();
      prefix.selectNodeContents(node);
      prefix.setEnd(caret.startContainer, caret.startOffset);
      offset = sourceOffsetForRenderedText(originalText, prefix.toString());
    }
    const inline = target?.closest<HTMLElement>('strong, em, code, a');
    const tag = inline?.tagName.toLowerCase() || '';
    const fragment = inline && node.contains(inline) && !inline.closest('pre')
      ? findInlineFragment(originalText, offset, tag) : null;
    const anchorNode = fragment && inline ? inline : node;
    const style = buildPreviewInlineEditStyle(anchorNode);
    if (!style) return;
    const text = fragment ? originalText.slice(fragment.from, fragment.to) : originalText;
    runtime.editorView.dispatch({ annotations: isolateHistory.of('before') });
    previewInlineEditSelection.value = selection;
    previewInlineEditSession.value = {
      from: range.from + (fragment?.from || 0),
      to: fragment ? range.from + fragment.to : range.to, startLine: selection.startLine,
      originalText: text, appliedText: text,
      ...(fragment ? { fragment: {
        selector: tag, index: [...node.querySelectorAll(tag)].indexOf(inline!),
      } } : {}),
    };
    previewInlineEditValue.value = text;
    input.previewInlineEditStyle.value = style;
    previewInlineEditVisible.value = true;
    anchorNode.classList.add(fragment ? 'preview-editing-fragment' : 'preview-editing-block');
    offset = Math.max(0, Math.min(text.length, offset - (fragment?.from || 0)));
    void nextTick(() => {
      refreshPreviewInlineEditorPosition();
      previewInlineEditTextarea.value?.focus({ preventScroll: true });
      previewInlineEditTextarea.value?.setSelectionRange(offset, offset);
      revealPreviewInlineEditorCaret();
    });
  }

  function handlePreviewInlineEditInput() {
    if (!previewInlineEditVisible.value || !previewInlineEditSession.value) {
      return;
    }

    refreshPreviewInlineEditorPosition();
    void nextTick(revealPreviewInlineEditorCaret);
    if (previewInlineEditComposing.value) {
      return;
    }

    queuePreviewInlineEditSync();
  }

  function handlePreviewInlineEditCompositionStart() {
    previewInlineEditComposing.value = true;
  }

  function handlePreviewInlineEditCompositionEnd() {
    previewInlineEditComposing.value = false;
    flushPreviewInlineEditSync();
  }

  function handlePreviewInlineEditBlur() {
    if (!previewInlineEditVisible.value || previewInlineEditComposing.value) {
      return;
    }

    closePreviewInlineEditor();
  }

  function updatePreviewInlineEditorSelection(
    selectionStart: number,
    selectionEnd: number,
  ) {
    void nextTick(() => {
      const textarea = previewInlineEditTextarea.value;
      if (!textarea) {
        return;
      }

      const valueLength = previewInlineEditValue.value.length;
      const start = Math.min(Math.max(selectionStart, 0), valueLength);
      const end = Math.min(Math.max(selectionEnd, start), valueLength);
      textarea.focus({ preventScroll: true });
      textarea.setSelectionRange(start, end);
      revealPreviewInlineEditorCaret();
    });
  }

  return { openPreviewTableCellEditor,handlePreviewClick,handlePreviewInlineEditInput,
    handlePreviewInlineEditCompositionStart,handlePreviewInlineEditCompositionEnd,
    handlePreviewInlineEditBlur,updatePreviewInlineEditorSelection,
  };
}
