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
  const { docSearchCaseSensitive,docSearchMatchIndex,docSearchQuery,docSearchRegex,
    docSearchWholeWord,showDocSearchPanel,store,
  } = input.editor.foundation;
  const { docSearchMatches } = input.editor.presentation;
  const { selectDocMatch } = input.editor.commands;
  const { windowTitle } = input.workspace.presentation;

  watch(
    () => [
      docSearchQuery.value,
      docSearchCaseSensitive.value,
      docSearchWholeWord.value,
      docSearchRegex.value,
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

}
