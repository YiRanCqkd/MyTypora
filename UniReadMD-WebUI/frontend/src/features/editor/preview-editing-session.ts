import { Transaction } from '@codemirror/state';
import { isolateHistory } from '@codemirror/commands';
import { findInlineFragment } from './inline-fragment';
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
    markPreviewEditingBlocks,
  } = input;

  function refreshPreviewInlineEditorPosition() {
    if (!previewInlineEditVisible.value || !previewInlineEditSelection.value) {
      return;
    }

    const session = previewInlineEditSession.value;
    if (session?.fragment
      && !findInlineFragment(session.appliedText, 0, session.fragment.selector)) {
      closePreviewInlineEditor();
      return;
    }

    const anchorNode = findPreviewInlineEditAnchorNode();
    if (!anchorNode) {
      // 移除格式标记后结束片段编辑，已写回的普通文本仍保留。
      if (previewInlineEditSession.value?.fragment) closePreviewInlineEditor();
      return;
    }

    const nextStyle = buildPreviewInlineEditStyle(anchorNode);
    if (!nextStyle) {
      return;
    }

    previewInlineEditStyle.value = nextStyle;
    anchorNode.classList.add(previewInlineEditSession.value?.fragment
      ? 'preview-editing-fragment' : 'preview-editing-block');
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
        anchor: session.from + Math.min(insert.length,
          input.previewInlineEditTextarea.value?.selectionStart ?? insert.length),
        head: session.from + Math.min(insert.length,
          input.previewInlineEditTextarea.value?.selectionEnd ?? insert.length),
      },
      annotations: Transaction.userEvent.of('input.type'),
    });

    const lineDelta = getPreviewInlineEditLineCount(insert)
      - getPreviewInlineEditLineCount(session.appliedText);
    previewInlineEditSession.value = {
      ...session,
      to: session.from + insert.length,
      appliedText: insert,
    };
    previewInlineEditSelection.value = {
      ...selection,
      endLine: selection.endLine + lineDelta,
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
    const anchorNode = findPreviewInlineEditAnchorNode();
    markPreviewEditingBlocks?.(true);
    anchorNode?.classList.remove('preview-editing-block');
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
    runtime.editorView?.dispatch({ annotations: isolateHistory.of('after') });
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
