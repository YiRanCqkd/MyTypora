import type { usePreviewEditingSession } from './preview-editing-session';
import type { useDocumentSearchActions } from './document-search-actions';
import type { useTableCellEditing } from './table-cell-editing';
import type { usePreviewInteractions } from './preview-interactions';
import type { useDocumentExport } from '../export/export-document';
import type { useSourceEditing } from './source-editing';
import type { useDocumentSession } from '../workspace/document-session';
import type { usePreviewEditingState } from './preview-editing-state';
import type { useDocumentSearchState } from './document-search-state';
import type { useWorkbenchPreferences } from '../settings/workbench-preferences';
type UseKeyboardCommandsDependencies =
  ReturnType<typeof usePreviewEditingSession> &
  ReturnType<typeof useDocumentSearchActions> &
  ReturnType<typeof useTableCellEditing> &
  ReturnType<typeof usePreviewInteractions> &
  ReturnType<typeof useDocumentExport> &
  ReturnType<typeof useSourceEditing> &
  ReturnType<typeof useDocumentSession> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof useDocumentSearchState> &
  ReturnType<typeof useWorkbenchPreferences>;

export function useKeyboardCommands(input: UseKeyboardCommandsDependencies) {
  const { cancelPreviewInlineEdit,closeDocSearchPanel,closePreviewTableCellEditor,
    closePreviewTableMenu,closeSourceContextMenu,exportHtml,goToNextDocMatch,openDocSearchPanel,
    openFile,previewInlineEditVisible,previewTableCellEditVisible,previewTableMenuVisible,
    saveAsFile,saveFile,newFile,showDocSearchPanel,sourceContextMenuVisible,toggleFocusMode,
    toggleSpellPanel,toggleTypewriterMode,
  } = input;

  function handleGlobalKeydown(event: KeyboardEvent) {
    if (event.isComposing || input.previewInlineEditComposing.value) return;
    if (event.key === 'Escape') {
      if (previewTableMenuVisible.value) {
        event.preventDefault();
        closePreviewTableMenu();
        return;
      }

      if (sourceContextMenuVisible.value) {
        event.preventDefault();
        closeSourceContextMenu();
        return;
      }

      if (showDocSearchPanel.value) {
        event.preventDefault();
        closeDocSearchPanel();
        return;
      }
    }

    if (
      !sourceContextMenuVisible.value &&
      event.key === 'Escape' &&
      previewInlineEditVisible.value
    ) {
      event.preventDefault();
      cancelPreviewInlineEdit();
      return;
    }

    if (
      !sourceContextMenuVisible.value
      && event.key === 'Escape'
      && previewTableCellEditVisible.value
    ) {
      event.preventDefault();
      closePreviewTableCellEditor();
      return;
    }

    if (!(event.ctrlKey || event.metaKey)) {
      if (event.key === 'F3') {
        event.preventDefault();
        goToNextDocMatch();
      }
      return;
    }

    if (event.altKey && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      toggleFocusMode();
      return;
    }

    if (event.altKey && event.key.toLowerCase() === 't') {
      event.preventDefault();
      toggleTypewriterMode();
      return;
    }

    if (event.altKey && event.key.toLowerCase() === 'l') {
      event.preventDefault();
      toggleSpellPanel();
      return;
    }

    if (event.key.toLowerCase() === 'f') {
      event.preventDefault();
      openDocSearchPanel(false);
      return;
    }

    if (event.key.toLowerCase() === 'h') {
      event.preventDefault();
      openDocSearchPanel(true);
      return;
    }

    if (event.key.toLowerCase() === 'o') {
      event.preventDefault();
      void openFile();
      return;
    }
    if (event.key.toLowerCase() === 'n') {
      event.preventDefault();
      void newFile();
      return;
    }

    if (event.key.toLowerCase() === 's') {
      event.preventDefault();
      if (event.shiftKey) {
        void saveAsFile();
        return;
      }

      void saveFile();
      return;
    }

    if (event.key.toLowerCase() === 'e') {
      event.preventDefault();
      void exportHtml();
    }
  }

  return { handleGlobalKeydown,
  };
}
