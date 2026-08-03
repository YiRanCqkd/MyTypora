import type { PreviewTaskItemContext } from './editor-contracts';
import { Transaction } from '@codemirror/state';
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
  const { closePreviewInlineEditor,errorMessage,getEditorRangeByLines,previewInlineEditVisible,
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

    const doc = runtime.editorView.state.doc;
    const taskPrefixPattern = /^(\s*(?:[-+*]|\d+[.)])\s+\[)([ xX])(\]\s*)/;
    const allMatchedLineNumbers: number[] = [];
    for (let lineNo = 1; lineNo <= doc.lines; lineNo += 1) {
      const line = doc.line(lineNo);
      if (taskPrefixPattern.test(line.text)) {
        allMatchedLineNumbers.push(lineNo);
      }
    }
    if (allMatchedLineNumbers.length === 0) {
      errorMessage.value = '未定位到任务列表项源码行，请切换 Source 手动调整。';
      return false;
    }

    const globalLineNo = allMatchedLineNumbers[
      Math.min(context.globalTaskIndex, allMatchedLineNumbers.length - 1)
    ];
    const globalLine = doc.line(globalLineNo);
    const globalNextLineText = globalLine.text.replace(
      taskPrefixPattern,
      (_match, start, current, end) => {
        const marker = current === ' ' ? 'x' : ' ';
        return `${start}${marker}${end}`;
      },
    );
    if (globalNextLineText !== globalLine.text) {
      runtime.editorView.dispatch({
        changes: {
          from: globalLine.from,
          to: globalLine.to,
          insert: globalNextLineText,
        },
        selection: {
          anchor: globalLine.from + globalNextLineText.length,
        },
        annotations: Transaction.userEvent.of('input.type'),
      });
      errorMessage.value = '';
      return true;
    }

    const minLine = Math.max(1, context.startLine - 1);
    const maxLine = Math.min(doc.lines, context.endLine + 1);

    const matchedLineNumbers: number[] = [];
    for (let lineNo = minLine; lineNo <= maxLine; lineNo += 1) {
      const line = doc.line(lineNo);
      if (taskPrefixPattern.test(line.text)) {
        matchedLineNumbers.push(lineNo);
      }
    }
    if (matchedLineNumbers.length === 0) {
      errorMessage.value = '未定位到任务列表项源码行，请切换 Source 手动调整。';
      return false;
    }

    const lineNo = matchedLineNumbers[Math.min(context.taskIndex, matchedLineNumbers.length - 1)];
    const line = doc.line(lineNo);
    const nextLineText = line.text.replace(taskPrefixPattern, (_match, start, current, end) => {
      const marker = current === ' ' ? 'x' : ' ';
      return `${start}${marker}${end}`;
    });
    if (nextLineText === line.text) {
      return true;
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

    // 预览态禁用双击进入源码覆盖编辑，避免与渲染内容重叠。
    if (previewInlineEditVisible.value) {
      closePreviewInlineEditor();
    }
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
