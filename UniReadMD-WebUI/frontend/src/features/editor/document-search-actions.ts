import { EditorView } from '@codemirror/view';
import { nextTick } from 'vue';
import { replaceCurrentSearchMatch, replaceSearchMatches } from './search-pattern';
import type { useDocumentSearchMatches } from './document-search-matches';
import type { usePreviewInteractions } from './preview-interactions';
import type { useCursorPosition } from './cursor-position';
import type { useDocumentSearchState } from './document-search-state';
import type { createPreviewSearchDom } from './preview-search-dom';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createWorkbenchRuntime } from './workbench-runtime';
import type { usePreviewNavigation } from './preview-navigation';
import type { useViewOptions } from '../settings/view-options';
import type { useWorkbenchStore } from './editor-store';
type UseDocumentSearchActionsDependencies =
  ReturnType<typeof useDocumentSearchMatches> &
  ReturnType<typeof usePreviewInteractions> &
  ReturnType<typeof useCursorPosition> &
  ReturnType<typeof useDocumentSearchState> &
  ReturnType<typeof createPreviewSearchDom> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof usePreviewNavigation> &
  ReturnType<typeof useViewOptions> &
  ReturnType<typeof useWorkbenchStore>;

export function useDocumentSearchActions(input: UseDocumentSearchActionsDependencies) {
  const { buildSearchRegExp,closePreviewTableMenu,closeSourceContextMenu,currentCursorLine,
    docReplaceQuery,docSearchCaseSensitive,docSearchMatchIndex,docSearchMatches,docSearchQuery,
    docSearchWholeWord,docSearchRegex,normalizeSearchIndex,previewTableMenuVisible,runtime,
    scrollPreviewToLine,
    showDocReplacePanel,showDocSearchPanel,sourceContextMenuVisible,sourceMode,store,
  } = input;

  function closeDocSearchPanel() {
    showDocSearchPanel.value = false;
    showDocReplacePanel.value = false;
  }

  function resolveLineNumberByOffset(offset: number) {
    if (runtime.editorView) {
      return runtime.editorView.state.doc.lineAt(offset).number;
    }

    const safeOffset = Math.max(0, Math.min(offset, store.content.length));
    const prefix = store.content.slice(0, safeOffset);
    return prefix.split('\n').length;
  }

  function selectDocMatch(index: number, options?: { focusEditor?: boolean }) {
    const matches = docSearchMatches.value;
    if (!runtime.editorView || matches.length === 0) {
      return;
    }

    const normalized = normalizeSearchIndex(index, matches.length);
    docSearchMatchIndex.value = normalized;
    const target = matches[normalized];

    runtime.editorView.dispatch({
      selection: {
        anchor: target.from,
        head: target.to,
      },
      effects: EditorView.scrollIntoView(target.from, {
        y: 'center',
      }),
    });

    const lineNumber = resolveLineNumberByOffset(target.from);
    currentCursorLine.value = lineNumber;
    store.setCursorHint(lineNumber);
    scrollPreviewToLine(lineNumber);

    if (options?.focusEditor) {
      runtime.editorView.focus();
    }
  }

  function goToNextDocMatch() {
    selectDocMatch(docSearchMatchIndex.value + 1, {
      focusEditor: sourceMode.value && !showDocSearchPanel.value,
    });
  }

  function goToPrevDocMatch() {
    selectDocMatch(docSearchMatchIndex.value - 1, {
      focusEditor: sourceMode.value && !showDocSearchPanel.value,
    });
  }

  function replaceCurrentDocMatch() {
    const matches = docSearchMatches.value;
    if (matches.length === 0) {
      return;
    }

    const normalized = normalizeSearchIndex(docSearchMatchIndex.value, matches.length);
    const target = matches[normalized];
    const expression = buildSearchRegExp(docSearchQuery.value,
      docSearchCaseSensitive.value, docSearchWholeWord.value, false);
    if (!expression) return;
    const nextContent = replaceCurrentSearchMatch(
      store.content, expression, target.from, docReplaceQuery.value, docSearchRegex?.value,
    );
    store.setContent(nextContent);

    nextTick(() => {
      if (docSearchMatches.value.length > 0) {
        selectDocMatch(normalized, {
          focusEditor: false,
        });
      }
    });
  }

  function replaceAllDocMatch() {
    const expression = buildSearchRegExp(
      docSearchQuery.value,
      docSearchCaseSensitive.value,
      docSearchWholeWord.value,
      true,
    );

    if (!expression || docSearchMatches.value.length === 0) {
      return;
    }

    const nextContent = replaceSearchMatches(
      store.content, expression, docReplaceQuery.value, docSearchRegex?.value,
    );
    store.setContent(nextContent);

    nextTick(() => {
      docSearchMatchIndex.value = 0;
      if (docSearchMatches.value.length > 0) {
        selectDocMatch(0, {
          focusEditor: false,
        });
      }
    });
  }

  function handleDocSearchEnter(isShift: boolean) {
    if (isShift) {
      goToPrevDocMatch();
      return;
    }

    goToNextDocMatch();
  }

  function handleGlobalPointerDown(event: MouseEvent) {
    if (!previewTableMenuVisible.value && !sourceContextMenuVisible.value) {
      return;
    }

    const target = event.target as HTMLElement | null;
    if (!target) {
      closePreviewTableMenu();
      closeSourceContextMenu();
      return;
    }

    if (previewTableMenuVisible.value && target.closest('.preview-table-menu')) {
      return;
    }

    if (sourceContextMenuVisible.value && target.closest('.source-context-menu')) {
      return;
    }

    closePreviewTableMenu();
    closeSourceContextMenu();
  }

  return { closeDocSearchPanel,resolveLineNumberByOffset,selectDocMatch,
    goToNextDocMatch,goToPrevDocMatch,replaceCurrentDocMatch,replaceAllDocMatch,
    handleDocSearchEnter,handleGlobalPointerDown,
  };
}
