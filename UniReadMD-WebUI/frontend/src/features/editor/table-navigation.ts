import type { PreviewTableSelection } from './editor-contracts';
import { nextTick } from 'vue';
import { parseMarkdownTable, applyPreviewTableActionToModel } from '../../utils/preview-table';
import type { useTableCellEditing } from './table-cell-editing';
import type { createTableMutations } from './table-mutations';
import type { createPreviewEditingRanges } from './preview-editing-ranges';
import type { usePreviewEditingEvents } from './preview-editing-events';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createWorkbenchRuntime } from './workbench-runtime';
type UseTableNavigationDependencies =
  ReturnType<typeof useTableCellEditing> &
  ReturnType<typeof createTableMutations> &
  ReturnType<typeof createPreviewEditingRanges> &
  ReturnType<typeof usePreviewEditingEvents> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime>;

export function useTableNavigation(input: UseTableNavigationDependencies) {
  const { closePreviewTableCellEditor,closePreviewTableCellEditorAndCommit,
    commitPreviewTableCellEditValue,getEditorRangeByLines,mutatePreviewTable,
    openPreviewTableCellEditor,previewTableCellEditSelection,runtime,
  } = input;

  function getPreviewTableMetrics(selection: PreviewTableSelection) {
    if (!runtime.editorView) {
      return null;
    }

    const range = getEditorRangeByLines(selection.startLine, selection.endLine);
    if (!range) {
      return null;
    }

    const fragment = runtime.editorView.state.doc.sliceString(range.from, range.to);
    const table = parseMarkdownTable(fragment);
    if (!table) {
      return null;
    }

    return {
      rowCount: table.rows.length + 1,
      colCount: table.header.length,
    };
  }

  async function movePreviewTableCellEditor(direction: 'tab-next' | 'tab-prev' | 'enter-down') {
    const activeSelection = previewTableCellEditSelection.value;
    if (!activeSelection) {
      return;
    }

    const commitResult = commitPreviewTableCellEditValue();
    const committedSelection = commitResult?.selection || activeSelection;
    const metrics = getPreviewTableMetrics(committedSelection);
    if (!metrics) {
      closePreviewTableCellEditor();
      return;
    }

    let targetRow = committedSelection.rowIndex;
    let targetCol = committedSelection.colIndex;
    let needAppendRow = false;
    if (direction === 'tab-next') {
      if (targetCol < metrics.colCount - 1) {
        targetCol += 1;
      } else if (targetRow < metrics.rowCount - 1) {
        targetRow += 1;
        targetCol = 0;
      } else {
        needAppendRow = true;
        targetRow += 1;
        targetCol = 0;
      }
    } else if (direction === 'tab-prev') {
      if (targetCol > 0) {
        targetCol -= 1;
      } else if (targetRow > 0) {
        targetRow -= 1;
        targetCol = metrics.colCount - 1;
      }
    } else if (direction === 'enter-down') {
      if (targetRow < metrics.rowCount - 1) {
        targetRow += 1;
      } else {
        needAppendRow = true;
        targetRow += 1;
      }
    }

    let finalSelection: PreviewTableSelection = {
      ...committedSelection,
    };
    if (needAppendRow) {
      const appendResult = mutatePreviewTable(finalSelection, (table, currentSelection) => {
        const result = applyPreviewTableActionToModel(
          table,
          {
            rowIndex: currentSelection.rowIndex,
            colIndex: currentSelection.colIndex,
          },
          'insert-row-below',
        );
        return result.changed;
      });
      if (!appendResult) {
        closePreviewTableCellEditor();
        return;
      }
      finalSelection = appendResult.selection;
    }

    closePreviewTableCellEditor();
    await nextTick();
    openPreviewTableCellEditor({
      ...finalSelection,
      rowIndex: targetRow,
      colIndex: targetCol,
    });
  }

  function handlePreviewTableCellEditBlur() {
    closePreviewTableCellEditorAndCommit();
  }

  return { getPreviewTableMetrics,movePreviewTableCellEditor,
    handlePreviewTableCellEditBlur,
  };
}
