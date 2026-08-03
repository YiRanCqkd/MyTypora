import type { ParsedMarkdownTable } from '../../utils/preview-table';
import type { PreviewTableSelection } from './editor-contracts';
import { parseMarkdownTable, serializeMarkdownTable } from '../../utils/preview-table';
import type { useTableCellEditing } from './table-cell-editing';
import type { useApplicationStatus } from './application-status';
import type { createPreviewEditingRanges } from './preview-editing-ranges';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createWorkbenchRuntime } from './workbench-runtime';
type CreateTableMutationsDependencies =
  ReturnType<typeof useTableCellEditing> &
  ReturnType<typeof useApplicationStatus> &
  ReturnType<typeof createPreviewEditingRanges> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime>;

export function createTableMutations(input: CreateTableMutationsDependencies) {
  const { closePreviewTableCellEditor,errorMessage,getEditorRangeByLines,
    previewTableCellEditSelection,previewTableCellEditValue,previewTableCellEditVisible,
    previewTableSelection,runtime,
  } = input;

  function mutatePreviewTable(
    selection: PreviewTableSelection,
    mutator: (table: ParsedMarkdownTable, selection: PreviewTableSelection) => boolean,
  ) {
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
      errorMessage.value = '无法解析当前表格源码，请切换 Source 模式手动调整。';
      return null;
    }

    if (!mutator(table, selection)) {
      return {
        changed: false,
        selection,
        table,
      };
    }

    const insert = serializeMarkdownTable(table);
    const lineCount = insert.split('\n').length;
    const nextSelection = {
      ...selection,
      endLine: selection.startLine + lineCount - 1,
    };
    runtime.editorView.dispatch({
      changes: {
        from: range.from,
        to: range.to,
        insert,
      },
      selection: {
        anchor: range.from,
      },
    });

    previewTableSelection.value = {
      ...nextSelection,
    };
    if (previewTableCellEditSelection.value) {
      previewTableCellEditSelection.value = {
        ...previewTableCellEditSelection.value,
        endLine: nextSelection.endLine,
      };
    }

    return {
      changed: true,
      selection: nextSelection,
      table,
    };
  }

  function commitPreviewTableCellEditValue() {
    const selection = previewTableCellEditSelection.value;
    if (!selection) {
      return null;
    }

    const nextValue = previewTableCellEditValue.value.replace(/\r?\n/g, ' ').trim();
    const result = mutatePreviewTable(selection, (table, currentSelection) => {
      if (currentSelection.rowIndex <= 0) {
        const current = table.header[currentSelection.colIndex] || '';
        if (current === nextValue) {
          return false;
        }
        table.header[currentSelection.colIndex] = nextValue;
        return true;
      }

      const bodyRowIndex = currentSelection.rowIndex - 1;
      if (!table.rows[bodyRowIndex]) {
        return false;
      }

      const current = table.rows[bodyRowIndex][currentSelection.colIndex] || '';
      if (current === nextValue) {
        return false;
      }
      table.rows[bodyRowIndex][currentSelection.colIndex] = nextValue;
      return true;
    });

    if (!result) {
      return null;
    }

    if (result.changed) {
      previewTableCellEditSelection.value = {
        ...result.selection,
      };
      previewTableCellEditValue.value = nextValue;
    }

    return {
      selection: result.selection,
      table: result.table,
    };
  }

  function closePreviewTableCellEditorAndCommit() {
    if (previewTableCellEditVisible.value) {
      commitPreviewTableCellEditValue();
    }
    closePreviewTableCellEditor();
  }

  return { mutatePreviewTable,commitPreviewTableCellEditValue,
    closePreviewTableCellEditorAndCommit,
  };
}
