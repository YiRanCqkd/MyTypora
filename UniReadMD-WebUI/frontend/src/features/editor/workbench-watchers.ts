import { watch, nextTick } from 'vue';
import type { EditorDomain, SettingsDomain, WorkspaceDomain } from './workbench-domains';

export type WorkbenchWatcherFeatures = Readonly<{
  editor: Readonly<{
    commands: ReturnType<EditorDomain['commands']>;
    foundation: EditorDomain['foundation'];
    presentation: ReturnType<EditorDomain['presentation']>;
  }>;
  settings: Readonly<{
    foundation: ReturnType<SettingsDomain['foundation']>;
    session: ReturnType<SettingsDomain['session']>;
  }>;
  workspace: Readonly<{
    foundation: ReturnType<WorkspaceDomain['foundation']>;
    presentation: ReturnType<WorkspaceDomain['presentation']>;
  }>;
}>;

export function useWorkbenchWatchers(input: WorkbenchWatcherFeatures) {
  const {
    contentZoom,focusMode,pinOutline,sourceMode,typewriterMode,
  } = input.settings.foundation;
  const { docReplaceQuery,docSearchCaseSensitive,docSearchMatchIndex,docSearchQuery,
    docSearchWholeWord,showDocReplacePanel,showDocSearchPanel,store,
  } = input.editor.foundation;
  const { docSearchMatches } = input.editor.presentation;
  const { fileSortMode,fileTreeExpanded,fileViewMode,filesFilter,mainPreviewWidth,recentFiles,
    sidebarWidth } = input.workspace.foundation;
  const { persistSessionSnapshot } = input.settings.session;
  const { selectDocMatch } = input.editor.commands;
  const { windowTitle } = input.workspace.presentation;

  watch(
    () => [
      docSearchQuery.value,
      docSearchCaseSensitive.value,
      docSearchWholeWord.value,
      store.content,
    ],
    () => {
      docSearchMatchIndex.value = 0;

      if (docSearchMatches.value.length > 0 && showDocSearchPanel.value) {
        nextTick(() => {
          selectDocMatch(0, {
            focusEditor: false,
          });
        });
      }
    },
  );

  watch(
    () => showDocSearchPanel.value,
    (visible) => {
      if (visible && docSearchMatches.value.length > 0) {
        nextTick(() => {
          selectDocMatch(docSearchMatchIndex.value, {
            focusEditor: false,
          });
        });
      }
    },
  );

  watch(
    () => windowTitle.value,
    (nextTitle) => {
      document.title = nextTitle;
    },
    {
      immediate: true,
    },
  );

  watch(
    () => [
      store.filePath,
      store.content,
      store.isDirty,
      store.cursorHint,
      recentFiles.value.join('\n'),
      pinOutline.value,
      sourceMode.value,
      focusMode.value,
      typewriterMode.value,
      contentZoom.value,
      sidebarWidth.value,
      mainPreviewWidth.value,
      fileViewMode.value,
      fileSortMode.value,
      filesFilter.value,
      showDocSearchPanel.value,
      showDocReplacePanel.value,
      docSearchQuery.value,
      docReplaceQuery.value,
      docSearchCaseSensitive.value,
      docSearchWholeWord.value,
      docSearchMatchIndex.value,
      JSON.stringify(fileTreeExpanded.value),
    ],
    () => {
      persistSessionSnapshot();
    },
  );

}
