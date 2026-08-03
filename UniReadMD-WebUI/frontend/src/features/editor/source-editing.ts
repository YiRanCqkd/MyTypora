import type { InlineShortcutKind } from '../../utils/markdown-shortcuts';
import { EditorView } from '@codemirror/view';
import { Transaction } from '@codemirror/state';
import { buildInlineShortcutPatch } from '../../utils/markdown-shortcuts';
import { indentLess, indentMore } from '@codemirror/commands';
import { nextTick } from 'vue';
import type { usePreviewEditingSession } from './preview-editing-session';
import type { useTableCellEditing } from './table-cell-editing';
import type { usePreviewInteractions } from './preview-interactions';
import type { createTableMutations } from './table-mutations';
import type { useDocumentSearchState } from './document-search-state';
import type { useEditorHosts } from './editor-hosts';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createWorkbenchRuntime } from './workbench-runtime';
import type { createPreviewEditingRanges } from './preview-editing-ranges';
type UseSourceEditingDependencies =
  ReturnType<typeof usePreviewEditingSession> &
  ReturnType<typeof useTableCellEditing> &
  ReturnType<typeof usePreviewInteractions> &
  ReturnType<typeof createTableMutations> &
  ReturnType<typeof useDocumentSearchState> &
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof createPreviewEditingRanges>;

export function useSourceEditing(input: UseSourceEditingDependencies) {
  const { closePreviewInlineEditor,closePreviewTableCellEditor,closePreviewTableMenu,
    commitPreviewTableCellEditValue,docReplaceQuery,docSearchInput,docSearchQuery,previewHost,
    previewInlineEditVisible,previewTableCellEditVisible,runtime,schedulePreviewScrollSync,
    showDocReplacePanel,showDocSearchPanel,
  } = input;

  function handlePreviewInlineEditorWheel(event: WheelEvent) {
    const preview = previewHost.value;
    if (!preview) {
      return;
    }

    event.preventDefault();
    preview.scrollTop += event.deltaY;
  }

  function handlePreviewScroll() {
    schedulePreviewScrollSync();
    closePreviewTableMenu();
    if (previewTableCellEditVisible.value) {
      commitPreviewTableCellEditValue();
      closePreviewTableCellEditor();
    }

    if (!previewInlineEditVisible.value) {
      return;
    }

    closePreviewInlineEditor();
  }

  function selectAllSourceContent() {
    if (!runtime.editorView) {
      return;
    }

    const total = runtime.editorView.state.doc.length;
    runtime.editorView.dispatch({
      selection: {
        anchor: 0,
        head: total,
      },
    });
    runtime.editorView.focus();
  }

  function runLegacyEditCommand(command: 'cut' | 'copy' | 'paste' | 'selectAll') {
    try {
      return document.execCommand(command);
    } catch {
      return false;
    }
  }

  function applySourceInlineShortcut(
    kind: InlineShortcutKind,
    viewArg?: EditorView | null,
  ) {
    const targetView = viewArg || runtime.editorView;
    if (!targetView) {
      return false;
    }

    const selection = targetView.state.selection.main;
    const selectionText = targetView.state.doc.sliceString(selection.from, selection.to);
    const patch = buildInlineShortcutPatch(selectionText, kind);
    if (!patch.changed) {
      return false;
    }

    targetView.dispatch({
      changes: {
        from: selection.from,
        to: selection.to,
        insert: patch.insert,
      },
      selection: {
        anchor: selection.from + patch.selectionStart,
        head: selection.from + patch.selectionEnd,
      },
      annotations: Transaction.userEvent.of('input.type'),
    });
    return true;
  }

  function applySourceIndentShortcut(
    outdent: boolean,
    viewArg?: EditorView | null,
  ) {
    const targetView = viewArg || runtime.editorView;
    if (!targetView) {
      return false;
    }

    if (outdent) {
      return indentLess(targetView);
    }

    return indentMore(targetView);
  }

  function openDocSearchPanel(replaceMode = false) {
    showDocSearchPanel.value = true;
    showDocReplacePanel.value = replaceMode;

    if (!replaceMode && docReplaceQuery.value && !docSearchQuery.value) {
      docSearchQuery.value = docReplaceQuery.value;
    }

    void nextTick(() => {
      const input = docSearchInput.value;
      if (!input) {
        return;
      }

      input.focus();
      input.select();
    });
  }

  return { handlePreviewInlineEditorWheel,handlePreviewScroll,selectAllSourceContent,
    runLegacyEditCommand,applySourceInlineShortcut,applySourceIndentShortcut,openDocSearchPanel,
  };
}
