import type { PreviewTableAction, ParsedMarkdownTable } from '../../utils/preview-table';
import type { PreviewTableSelection } from './editor-contracts';
import { applyPreviewTableActionToModel } from '../../utils/preview-table';
import { nextTick } from 'vue';
import { measuredContextMenuStyle } from '../workspace/context-menu-position';
import type { useContextMenuActions } from './context-menu-actions';
import type { usePreviewEditingSession } from './preview-editing-session';
import type { useTableCellEditing } from './table-cell-editing';
import type { createTableMutations } from './table-mutations';
import type { usePreviewInteractions } from './preview-interactions';
import type { useEditorHosts } from './editor-hosts';
import type { useApplicationStatus } from './application-status';
import type { usePreviewEditingEvents } from './preview-editing-events';
import type { usePreviewEditingState } from './preview-editing-state';
import type { useViewOptions } from '../settings/view-options';
import type { createWorkbenchRuntime } from './workbench-runtime';
type UseTableActionsDependencies =
  ReturnType<typeof useContextMenuActions> &
  ReturnType<typeof usePreviewEditingSession> &
  ReturnType<typeof useTableCellEditing> &
  ReturnType<typeof createTableMutations> &
  ReturnType<typeof usePreviewInteractions> &
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof useApplicationStatus> &
  ReturnType<typeof usePreviewEditingEvents> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof useViewOptions> &
  ReturnType<typeof createWorkbenchRuntime>;

export function useTableActions(input: UseTableActionsDependencies) {
  const { applyPreviewTableMutation,closePreviewInlineEditor,closePreviewTableCellEditor,
    closePreviewTableCellEditorAndCommit,closePreviewTableMenu,closeSourceContextMenu,
    commitPreviewTableCellEditValue,editorPreviewHost,errorMessage,findPreviewTableCellNode,
    mutatePreviewTable,
    openPreviewTableCellEditor,previewTableCellEditSelection,previewTableMenuStyle,
    previewTableMenuVisible,previewTableSelection,resolvePreviewTableCellContext,sourceMode,runtime,
  } = input;

  function applyPreviewTableAction(action: string) {
    const tableAction = action as PreviewTableAction;
    applyPreviewTableMutation((table, selection) => {
      const result = applyPreviewTableActionToModel(
        table,
        {
          rowIndex: selection.rowIndex,
          colIndex: selection.colIndex,
        },
        tableAction,
      );
      if (!result.changed && result.errorMessage) {
        errorMessage.value = result.errorMessage;
      }
      return result.changed;
    });
  }

  async function handlePreviewContextMenu(event: MouseEvent) {
    if (sourceMode.value) {
      return;
    }

    const target = event.target instanceof HTMLElement ? event.target : null;
    const context = resolvePreviewTableCellContext(target);
    const mapped = runtime.flushVisualEdits?.(context || undefined) || context;
    if (!context) {
      closePreviewTableMenu();
      return;
    }

    const container = editorPreviewHost.value;
    if (!container) {
      closePreviewTableMenu();
      return;
    }

    event.preventDefault();
    closePreviewTableCellEditorAndCommit();
    closeSourceContextMenu();
    closePreviewInlineEditor();
    // 刷新后按换算后的范围重新定位单元格，不能沿用冻结预览的旧源码行号。
    await nextTick();
    const cell = mapped ? findPreviewTableCellNode(mapped) : null;
    const refreshed = resolvePreviewTableCellContext(cell);
    if (!refreshed) {
      closePreviewTableMenu();
      return;
    }
    previewTableSelection.value = {
      startLine: refreshed.startLine,
      endLine: refreshed.endLine,
      rowIndex: refreshed.rowIndex,
      colIndex: refreshed.colIndex,
    };
    const containerRect = container.getBoundingClientRect();
    previewTableMenuStyle.value = {
      left: `${event.clientX}px`, top: `${event.clientY}px`, visibility: 'hidden',
    };
    previewTableMenuVisible.value = true;
    void nextTick(async () => {
      const menu = document.querySelector<HTMLElement>('.preview-table-menu');
      if (!menu || !previewTableMenuVisible.value) return;
      previewTableMenuStyle.value = measuredContextMenuStyle(menu, event, {
        left: containerRect.left + 8, right: containerRect.right - 8,
        top: containerRect.top + 8, bottom: containerRect.bottom - 8,
      });
      await nextTick();
      if (previewTableMenuVisible.value) {
        menu.querySelector('button')?.focus({ preventScroll: true });
      }
    });
  }

  function getPreviewTableSelectionAfterAction(
    action: PreviewTableAction,
    selection: PreviewTableSelection,
    table: ParsedMarkdownTable,
  ) {
    const rowCount = table.rows.length + 1;
    const colCount = table.header.length;
    let rowIndex = selection.rowIndex;
    let colIndex = selection.colIndex;

    if (action === 'insert-row-below') {
      rowIndex = Math.min(selection.rowIndex + 1, rowCount - 1);
    } else if (action === 'insert-row-above') {
      rowIndex = selection.rowIndex;
    } else if (action === 'delete-row') {
      rowIndex = Math.min(selection.rowIndex, rowCount - 1);
    } else if (action === 'move-row-up') {
      rowIndex = Math.max(0, selection.rowIndex - 1);
    } else if (action === 'move-row-down') {
      rowIndex = Math.min(rowCount - 1, selection.rowIndex + 1);
    } else if (action === 'insert-col-right') {
      colIndex = Math.min(selection.colIndex + 1, colCount - 1);
    } else if (action === 'insert-col-left') {
      colIndex = selection.colIndex;
    } else if (action === 'delete-col') {
      colIndex = Math.min(selection.colIndex, colCount - 1);
    } else if (action === 'move-col-left') {
      colIndex = Math.max(0, selection.colIndex - 1);
    } else if (action === 'move-col-right') {
      colIndex = Math.min(colCount - 1, selection.colIndex + 1);
    }

    return {
      ...selection,
      rowIndex,
      colIndex,
    };
  }

  async function applyPreviewTableActionFromCellEditor(action: PreviewTableAction) {
    const activeSelection = previewTableCellEditSelection.value;
    if (!activeSelection) {
      return;
    }

    const commitResult = commitPreviewTableCellEditValue();
    const baseSelection = commitResult?.selection || activeSelection;
    const result = mutatePreviewTable(baseSelection, (table, selection) => {
      const actionResult = applyPreviewTableActionToModel(
        table,
        {
          rowIndex: selection.rowIndex,
          colIndex: selection.colIndex,
        },
        action,
      );
      if (!actionResult.changed && actionResult.errorMessage) {
        errorMessage.value = actionResult.errorMessage;
      }
      return actionResult.changed;
    });
    if (!result) {
      closePreviewTableCellEditor();
      return;
    }

    const nextSelection = getPreviewTableSelectionAfterAction(
      action,
      result.selection,
      result.table,
    );
    closePreviewTableCellEditor();
    await nextTick();
    openPreviewTableCellEditor(nextSelection, true);
  }

  return { applyPreviewTableAction,handlePreviewContextMenu,
    getPreviewTableSelectionAfterAction,applyPreviewTableActionFromCellEditor,
  };
}
