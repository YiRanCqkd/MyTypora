import { Transaction } from '@codemirror/state';
import type { createMarkdownRenderer } from '../markdown/renderer';
import type { createPreviewEditingRanges } from './preview-editing-ranges';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createWorkbenchRuntime } from './workbench-runtime';
type UsePreviewEditingSessionDependencies =
  ReturnType<typeof createMarkdownRenderer> &
  ReturnType<typeof createPreviewEditingRanges> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime>;

export function usePreviewEditingSession(input: UsePreviewEditingSessionDependencies) {
  const { PREVIEW_INLINE_EDIT_SYNC_DELAY,buildPreviewInlineEditStyle,
    findPreviewInlineEditAnchorNode,getPreviewInlineEditLineCount,
    normalizePreviewInlineEditContent,previewInlineEditComposing,previewInlineEditSelection,
    previewInlineEditSession,previewInlineEditStyle,previewInlineEditValue,
    previewInlineEditVisible,runtime,
  } = input;

  function refreshPreviewInlineEditorPosition() {
    if (!previewInlineEditVisible.value || !previewInlineEditSelection.value) {
      return;
    }

    const anchorNode = findPreviewInlineEditAnchorNode();
    if (!anchorNode) {
      return;
    }

    const nextStyle = buildPreviewInlineEditStyle(anchorNode);
    if (!nextStyle) {
      return;
    }

    previewInlineEditStyle.value = nextStyle;
    const nextToken = anchorNode.dataset.sourceToken || '';
    if (!nextToken) {
      return;
    }

    const selection = previewInlineEditSelection.value;
    if (!selection || selection.token === nextToken) {
      return;
    }

    previewInlineEditSelection.value = {
      ...selection,
      token: nextToken,
    };
  }

  function clearPreviewInlineEditSyncTimer() {
    if (!runtime.previewInlineEditSyncTimer) {
      return;
    }

    window.clearTimeout(runtime.previewInlineEditSyncTimer);
    runtime.previewInlineEditSyncTimer = 0;
  }

  function applyPreviewInlineEditToSource() {
    if (!previewInlineEditVisible.value || !runtime.editorView) {
      return;
    }

    const session = previewInlineEditSession.value;
    const selection = previewInlineEditSelection.value;
    if (!session || !selection) {
      return;
    }

    const insert = normalizePreviewInlineEditContent(previewInlineEditValue.value);
    if (insert === session.appliedText) {
      return;
    }

    runtime.editorView.dispatch({
      changes: {
        from: session.from,
        to: session.to,
        insert,
      },
      selection: {
        anchor: session.from + insert.length,
      },
      annotations: Transaction.userEvent.of('input.type'),
    });

    const lineCount = getPreviewInlineEditLineCount(insert);
    previewInlineEditSession.value = {
      ...session,
      to: session.from + insert.length,
      appliedText: insert,
    };
    previewInlineEditSelection.value = {
      ...selection,
      endLine: session.startLine + lineCount - 1,
    };
  }

  function flushPreviewInlineEditSync() {
    clearPreviewInlineEditSyncTimer();
    applyPreviewInlineEditToSource();
  }

  function queuePreviewInlineEditSync() {
    clearPreviewInlineEditSyncTimer();
    runtime.previewInlineEditSyncTimer = window.setTimeout(() => {
      runtime.previewInlineEditSyncTimer = 0;
      applyPreviewInlineEditToSource();
    }, PREVIEW_INLINE_EDIT_SYNC_DELAY);
  }

  function closePreviewInlineEditor(
    options?: {
      revert?: boolean;
    },
  ) {
    const currentSession = previewInlineEditSession.value;
    const shouldRevert = Boolean(options?.revert);
    if (shouldRevert) {
      clearPreviewInlineEditSyncTimer();
    } else {
      flushPreviewInlineEditSync();
    }

    previewInlineEditComposing.value = false;
    previewInlineEditVisible.value = false;
    previewInlineEditSelection.value = null;
    previewInlineEditStyle.value = {};
    previewInlineEditValue.value = '';

    if (currentSession && runtime.editorView) {
      if (shouldRevert && currentSession.appliedText !== currentSession.originalText) {
        runtime.editorView.dispatch({
          changes: {
            from: currentSession.from,
            to: currentSession.to,
            insert: currentSession.originalText,
          },
          selection: {
            anchor: currentSession.from + currentSession.originalText.length,
          },
          annotations: Transaction.addToHistory.of(false),
        });
      }
    }

    previewInlineEditSession.value = null;
  }

  function cancelPreviewInlineEdit() {
    closePreviewInlineEditor({
      revert: true,
    });
  }

  return { refreshPreviewInlineEditorPosition,clearPreviewInlineEditSyncTimer,
    applyPreviewInlineEditToSource,flushPreviewInlineEditSync,queuePreviewInlineEditSync,
    closePreviewInlineEditor,cancelPreviewInlineEdit,
  };
}
