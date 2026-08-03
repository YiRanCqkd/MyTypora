import type { PreviewTableAction, ParsedMarkdownTable } from '../../utils/preview-table';
import type { PreviewTableSelection } from './editor-contracts';
import { applyPreviewTableActionToModel } from '../../utils/preview-table';
import { nextTick } from 'vue';
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
  ReturnType<typeof useViewOptions>;

export function useTableActions(input: UseTableActionsDependencies) {
  const { applyPreviewTableMutation,closePreviewInlineEditor,closePreviewTableCellEditor,
    closePreviewTableCellEditorAndCommit,closePreviewTableMenu,closeSourceContextMenu,
    commitPreviewTableCellEditValue,editorPreviewHost,errorMessage,mutatePreviewTable,
    openPreviewTableCellEditor,previewTableCellEditSelection,previewTableMenuStyle,
    previewTableMenuVisible,previewTableSelection,resolvePreviewTableCellContext,sourceMode,
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

  function handlePreviewContextMenu(event: MouseEvent) {
    if (sourceMode.value) {
      return;
    }

    const target = event.target instanceof HTMLElement ? event.target : null;
    const context = resolvePreviewTableCellContext(target);
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
    previewTableSelection.value = {
      startLine: context.startLine,
      endLine: context.endLine,
      rowIndex: context.rowIndex,
      colIndex: context.colIndex,
    };
    const containerRect = container.getBoundingClientRect();
    previewTableMenuStyle.value = {
      left: `${event.clientX - containerRect.left}px`,
      top: `${event.clientY - containerRect.top}px`,
    };
    previewTableMenuVisible.value = true;
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
    openPreviewTableCellEditor(nextSelection);
  }

  return { applyPreviewTableAction,handlePreviewContextMenu,
    getPreviewTableSelectionAfterAction,applyPreviewTableActionFromCellEditor,
  };
}
