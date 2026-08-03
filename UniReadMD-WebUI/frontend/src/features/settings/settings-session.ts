import type { useEditorPresentation } from '../editor/editor-presentation';
import type { useWorkbenchFoundation } from '../editor/workbench-foundation';
import type { useMarkdownFeature } from '../markdown/markdown-feature';
import type { useWorkspaceFoundation } from '../workspace/workspace-foundation';
import type { useWorkspacePresentation } from '../workspace/workspace-presentation';
import { useContentZoom } from './content-zoom';
import { useLayoutModel } from './layout-model';
import type { useSettingsFoundation } from './settings-foundation';
import { useSessionPersistence } from './session-persistence';

type SettingsSessionDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  editorPresentation: ReturnType<typeof useEditorPresentation>;
  markdown: ReturnType<typeof useMarkdownFeature>;
  settings: ReturnType<typeof useSettingsFoundation>;
  workspace: ReturnType<typeof useWorkspaceFoundation>;
  workspacePresentation: ReturnType<typeof useWorkspacePresentation>;
}>;

export function useSettingsSession(input: SettingsSessionDependencies) {
  const layout = useLayoutModel({ ...input.editor, ...input.settings, ...input.workspace });
  const session = useSessionPersistence({
    ...input.editor,
    ...input.editorPresentation,
    ...input.markdown,
    ...input.settings,
    ...input.workspace,
    ...input.workspacePresentation,
  });
  const zoom = useContentZoom({
    ...input.editor,
    ...input.markdown,
    ...input.settings,
    ...input.workspacePresentation,
    ...session,
  });

  return {
    ...layout,
    ...session,
    ...zoom,
  };
}
