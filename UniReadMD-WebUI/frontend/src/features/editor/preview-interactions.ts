import type { PreviewTaskItemContext } from './editor-contracts';
import { Transaction } from '@codemirror/state';
import { toggleTaskLine } from './task-source';
import type { usePreviewEditingSession } from './preview-editing-session';
import type { useApplicationStatus } from './application-status';
import type { createPreviewEditingRanges } from './preview-editing-ranges';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createWorkbenchRuntime } from './workbench-runtime';
import type { useViewOptions } from '../settings/view-options';
type UsePreviewInteractionsDependencies =
  ReturnType<typeof usePreviewEditingSession> &
  ReturnType<typeof useApplicationStatus> &
  ReturnType<typeof createPreviewEditingRanges> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof useViewOptions>;

export function usePreviewInteractions(input: UsePreviewInteractionsDependencies) {
  const { closePreviewInlineEditor,errorMessage,getEditorRangeByLines,
    previewTableMenuStyle,previewTableMenuVisible,previewTableSelection,runtime,
    sourceContextMenuStyle,sourceContextMenuVisible,sourceMode,
  } = input;

  function togglePreviewTaskItem(context: PreviewTaskItemContext) {
    if (!runtime.editorView) {
      return false;
    }

    const range = getEditorRangeByLines(context.startLine, context.endLine);
    if (!range) {
      return false;
    }

    // 解析器已给出实际列表项的起始行，不能按全文任务序号猜测位置。
    const line = runtime.editorView.state.doc.line(context.startLine);
    const nextLineText = toggleTaskLine(line.text);
    if (nextLineText === null) {
      errorMessage.value = '任务源码映射已变化，请重试或切换源码模式。';
      return false;
    }

    runtime.editorView.dispatch({
      changes: {
        from: line.from,
        to: line.to,
        insert: nextLineText,
      },
      selection: {
        anchor: line.from + nextLineText.length,
      },
      annotations: Transaction.userEvent.of('input.type'),
    });
    errorMessage.value = '';
    return true;

  }

  function commitPreviewInlineEdit() {
    closePreviewInlineEditor();
  }

  function handlePreviewDoubleClick() {
    if (sourceMode.value) {
      return;
    }

    // 单击已建立编辑会话，双击保持会话以支持文本选择。
  }

  function closePreviewTableMenu() {
    previewTableMenuVisible.value = false;
    previewTableMenuStyle.value = {};
    previewTableSelection.value = null;
  }

  function closeSourceContextMenu() {
    sourceContextMenuVisible.value = false;
    sourceContextMenuStyle.value = {};
  }

  return { togglePreviewTaskItem,commitPreviewInlineEdit,handlePreviewDoubleClick,
    closePreviewTableMenu,closeSourceContextMenu,
  };
}
