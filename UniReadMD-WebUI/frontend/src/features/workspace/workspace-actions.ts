import type { useEditorNavigation } from '../editor/editor-navigation';
import type { useWorkbenchFoundation } from '../editor/workbench-foundation';
import type { useSettingsFoundation } from '../settings/settings-foundation';
import type { useSettingsSession } from '../settings/settings-session';
import { useDocumentSession } from './document-session';
import { useFileDrop } from './file-drop';
import { useFileTreeActions } from './file-tree-actions';
import { useOutlineNavigation } from './outline-navigation';
import type { useWorkspaceFoundation } from './workspace-foundation';
import type { useWorkspacePresentation } from './workspace-presentation';

type WorkspaceActionsDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  navigation: ReturnType<typeof useEditorNavigation>;
  settings: ReturnType<typeof useSettingsFoundation>;
  settingsSession: ReturnType<typeof useSettingsSession>;
  workspace: ReturnType<typeof useWorkspaceFoundation>;
  workspacePresentation: ReturnType<typeof useWorkspacePresentation>;
}>;

export function useWorkspaceActions(input: WorkspaceActionsDependencies) {
  const tree = useFileTreeActions({
    ...input.editor,
    ...input.settings,
    ...input.settingsSession,
    ...input.workspace,
    ...input.workspacePresentation,
  });
  const drop = useFileDrop({
    ...input.editor,
    ...input.workspace,
    ...input.workspacePresentation,
    ...tree,
  });
  const documents = useDocumentSession({
    ...input.editor,
    ...input.settingsSession,
    ...input.workspace,
    ...input.workspacePresentation,
    ...tree,
  });
  const outline = useOutlineNavigation({
    ...input.editor,
    ...input.navigation,
    ...input.settings,
    ...input.workspacePresentation,
    ...tree,
  });

  return {
    ...tree,
    ...drop,
    ...documents,
    ...outline,
  };
}
