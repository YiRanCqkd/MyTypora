import type { PreviewTableSelection } from './editor-contracts';
import { nextTick } from 'vue';
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
  ReturnType<typeof useWorkbenchStore>;

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
    sourceMode,store,togglePreviewTaskItem,
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
    if (sourceMode.value) {
      return;
    }

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
  }

  function handlePreviewInlineEditInput() {
    if (!previewInlineEditVisible.value || !previewInlineEditSession.value) {
      return;
    }

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
      textarea.focus();
      textarea.setSelectionRange(start, end);
    });
  }

  return { openPreviewTableCellEditor,handlePreviewClick,handlePreviewInlineEditInput,
    handlePreviewInlineEditCompositionStart,handlePreviewInlineEditCompositionEnd,
    handlePreviewInlineEditBlur,updatePreviewInlineEditorSelection,
  };
}
