import type { useCodeMirrorFeature } from '../editor/codemirror-feature';
import type { useEditorNavigation } from '../editor/editor-navigation';
import type { useWorkbenchFoundation } from '../editor/workbench-foundation';
import type { useWorkspaceFoundation } from '../workspace/workspace-foundation';
import type { useWorkspacePresentation } from '../workspace/workspace-presentation';
import type { useSettingsFoundation } from './settings-foundation';
import { useWorkbenchPreferences } from './workbench-preferences';

type SettingsActionsDependencies = Readonly<{
  codeMirror: ReturnType<typeof useCodeMirrorFeature>;
  editor: ReturnType<typeof useWorkbenchFoundation>;
  navigation: ReturnType<typeof useEditorNavigation>;
  settings: ReturnType<typeof useSettingsFoundation>;
  workspace: ReturnType<typeof useWorkspaceFoundation>;
  workspacePresentation: ReturnType<typeof useWorkspacePresentation>;
}>;

export function useSettingsActions(input: SettingsActionsDependencies) {
  return useWorkbenchPreferences({
    ...input.codeMirror,
    ...input.editor,
    ...input.navigation,
    ...input.settings,
    ...input.workspace,
    ...input.workspacePresentation,
  });
}
