import type { ParsedMarkdownTable } from '../../utils/preview-table';
import { readMarkdownClipboard, writeMarkdownClipboard } from './clipboard-markdown';
import type { PreviewTableSelection } from './editor-contracts';
import type { useTableCellEditing } from './table-cell-editing';
import type { usePreviewInteractions } from './preview-interactions';
import type { createTableMutations } from './table-mutations';
import type { useSourceEditing } from './source-editing';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createWorkbenchRuntime } from './workbench-runtime';
import type { useViewOptions } from '../settings/view-options';
type UseContextMenuActionsDependencies =
  ReturnType<typeof useTableCellEditing> &
  ReturnType<typeof usePreviewInteractions> &
  ReturnType<typeof createTableMutations> &
  ReturnType<typeof useSourceEditing> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof useViewOptions>;

export function useContextMenuActions(input: UseContextMenuActionsDependencies) {
  const { buildPreviewTableCellEditStyle,closePreviewTableMenu,closeSourceContextMenu,
    findPreviewTableCellNode,mutatePreviewTable,openDocSearchPanel,
    previewTableCellEditSelection,previewTableCellEditStyle,previewTableCellEditVisible,
    previewTableSelection,runLegacyEditCommand,runtime,selectAllSourceContent,sourceMode,
  } = input;

  async function handleSourceContextAction(action: string) {
    if (!sourceMode.value || !runtime.editorView) {
      closeSourceContextMenu();
      return;
    }

    if (action === 'search') {
      openDocSearchPanel(false);
      closeSourceContextMenu();
      return;
    }

    runtime.editorView.focus();
    if (action === 'select-all') {
      selectAllSourceContent();
      closeSourceContextMenu();
      return;
    }

    const selection = runtime.editorView.state.selection.main;
    if (action === 'copy') {
      const text = runtime.editorView.state.doc.sliceString(selection.from, selection.to);
      if (text.length > 0) {
        try {
          await writeMarkdownClipboard(text);
        } catch {
          runLegacyEditCommand('copy');
        }
      } else {
        runLegacyEditCommand('copy');
      }
      closeSourceContextMenu();
      return;
    }

    if (action === 'cut') {
      const text = runtime.editorView.state.doc.sliceString(selection.from, selection.to);
      if (text.length > 0) {
        let manualCopied = false;
        try {
          await writeMarkdownClipboard(text);
          manualCopied = true;
        } catch {
          runLegacyEditCommand('cut');
        }

        if (manualCopied) {
          runtime.editorView.dispatch({
            changes: {
              from: selection.from,
              to: selection.to,
              insert: '',
            },
            selection: {
              anchor: selection.from,
            },
          });
        }
      } else {
        runLegacyEditCommand('cut');
      }

      closeSourceContextMenu();
      return;
    }

    if (action === 'paste') {
      try {
        const text = await readMarkdownClipboard();
        if (typeof text === 'string') {
          runtime.editorView.dispatch({
            changes: {
              from: selection.from,
              to: selection.to,
              insert: text,
            },
            selection: {
              anchor: selection.from + text.length,
            },
          });
        }
      } catch {
        runLegacyEditCommand('paste');
      }

      closeSourceContextMenu();
    }
  }

  function refreshPreviewTableCellEditorPosition() {
    if (!previewTableCellEditVisible.value || !previewTableCellEditSelection.value) {
      return;
    }

    const cell = findPreviewTableCellNode(previewTableCellEditSelection.value);
    if (!cell) {
      return;
    }

    const nextStyle = buildPreviewTableCellEditStyle(cell);
    if (!nextStyle) {
      return;
    }

    previewTableCellEditStyle.value = nextStyle;
  }

  function applyPreviewTableMutation(
    mutator: (table: ParsedMarkdownTable, selection: PreviewTableSelection) => boolean,
  ) {
    if (!previewTableSelection.value) {
      closePreviewTableMenu();
      return;
    }

    const result = mutatePreviewTable(previewTableSelection.value, mutator);
    if (!result) {
      closePreviewTableMenu();
      return;
    }

    closePreviewTableMenu();
  }

  return { handleSourceContextAction,refreshPreviewTableCellEditorPosition,
    applyPreviewTableMutation,
  };
}
