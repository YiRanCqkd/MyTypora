import type { PreviewTableCellContext, PreviewTableSelection } from './editor-contracts';
import { nextTick } from 'vue';
import { parseMarkdownTable } from '../../utils/preview-table';
import type { useEditorHosts } from './editor-hosts';
import type { createPreviewEditingRanges } from './preview-editing-ranges';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createWorkbenchRuntime } from './workbench-runtime';
type UseTableCellEditingDependencies =
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof createPreviewEditingRanges> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime>;

export function useTableCellEditing(input: UseTableCellEditingDependencies) {
  const { editorPreviewHost,getEditorRangeByLines,previewHost,previewTableCellEditInput,
    previewTableCellEditSelection,previewTableCellEditStyle,previewTableCellEditValue,
    previewTableCellEditVisible,runtime,
  } = input;

  function resolvePreviewTableCellContext(target: HTMLElement | null) {
    if (!target) {
      return null;
    }

    const cell = target.closest<HTMLElement>('th,td');
    const table = target.closest<HTMLElement>('table[data-source-start][data-source-end]');
    if (!cell || !table) {
      return null;
    }

    const row = cell.closest('tr');
    if (!row) {
      return null;
    }

    const rows = Array.from(table.querySelectorAll('tr'));
    const rowIndex = rows.indexOf(row);
    if (rowIndex < 0) {
      return null;
    }

    const cells = Array.from(row.querySelectorAll('th,td'));
    const colIndex = cells.indexOf(cell);
    if (colIndex < 0) {
      return null;
    }

    const startLine = Number(table.dataset.sourceStart || 0);
    const endLine = Number(table.dataset.sourceEnd || 0);
    if (
      !Number.isFinite(startLine) ||
      !Number.isFinite(endLine) ||
      startLine <= 0 ||
      endLine <= 0
    ) {
      return null;
    }

    return {
      table,
      cell,
      startLine: Math.round(startLine),
      endLine: Math.round(endLine),
      rowIndex,
      colIndex,
    } as PreviewTableCellContext;
  }

  function focusPreviewTableCellEditor() {
    void nextTick(() => {
      const input = previewTableCellEditInput.value;
      if (!input) {
        return;
      }

      input.focus();
      input.select();
    });
  }

  function buildPreviewTableCellEditStyle(cell: HTMLElement) {
    const container = editorPreviewHost.value;
    if (!container) {
      return null;
    }

    const containerRect = container.getBoundingClientRect();
    const cellRect = cell.getBoundingClientRect();
    return {
      left: `${cellRect.left - containerRect.left}px`,
      top: `${cellRect.top - containerRect.top}px`,
      width: `${Math.max(92, Math.round(cellRect.width))}px`,
      minHeight: `${Math.max(30, Math.round(cellRect.height))}px`,
    };
  }

  function findPreviewTableCellNode(selection: PreviewTableSelection) {
    const preview = previewHost.value;
    if (!preview) {
      return null;
    }

    const tableSelector = `table[data-source-start='${selection.startLine}']`
      + `[data-source-end='${selection.endLine}']`;
    const table = preview.querySelector<HTMLElement>(tableSelector);
    if (!table) {
      return null;
    }

    const rows = Array.from(table.querySelectorAll('tr'));
    const row = rows[selection.rowIndex];
    if (!row) {
      return null;
    }

    const cells = Array.from(row.querySelectorAll<HTMLElement>('th,td'));
    return cells[selection.colIndex] || null;
  }

  function readPreviewTableCellValue(selection: PreviewTableSelection) {
    if (!runtime.editorView) {
      return '';
    }

    const range = getEditorRangeByLines(selection.startLine, selection.endLine);
    if (!range) {
      return '';
    }

    const fragment = runtime.editorView.state.doc.sliceString(range.from, range.to);
    const table = parseMarkdownTable(fragment);
    if (!table) {
      return '';
    }

    if (selection.rowIndex <= 0) {
      return table.header[selection.colIndex] || '';
    }

    const bodyRow = table.rows[selection.rowIndex - 1];
    return bodyRow?.[selection.colIndex] || '';
  }

  function closePreviewTableCellEditor() {
    previewTableCellEditVisible.value = false;
    previewTableCellEditSelection.value = null;
    previewTableCellEditStyle.value = {};
    previewTableCellEditValue.value = '';
  }

  return { resolvePreviewTableCellContext,focusPreviewTableCellEditor,
    buildPreviewTableCellEditStyle,findPreviewTableCellNode,readPreviewTableCellValue,
    closePreviewTableCellEditor,
  };
}
