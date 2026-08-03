import type { useEditorPresentation } from '../editor/editor-presentation';
import type { useWorkbenchFoundation } from '../editor/workbench-foundation';
import type { useMarkdownFeature } from '../markdown/markdown-feature';
import type { useSettingsFoundation } from '../settings/settings-foundation';
import { createFilePathModel } from './file-paths';
import { useFileTree } from './file-tree';
import { createLayoutMetrics } from './layout-metrics';
import { useSidebarModel } from './sidebar-model';
import { useTaskRunner } from './task-runner';
import type { useWorkspaceFoundation } from './workspace-foundation';

type WorkspacePresentationDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  editorPresentation: ReturnType<typeof useEditorPresentation>;
  markdown: ReturnType<typeof useMarkdownFeature>;
  settings: ReturnType<typeof useSettingsFoundation>;
  workspace: ReturnType<typeof useWorkspaceFoundation>;
}>;

export function useWorkspacePresentation(input: WorkspacePresentationDependencies) {
  const sidebar = useSidebarModel({
    ...input.editor,
    ...input.editorPresentation,
    ...input.workspace,
  });
  const paths = createFilePathModel({ ...input.editor, ...input.workspace });
  const tree = useFileTree({ ...paths, ...input.workspace });
  const metrics = createLayoutMetrics({
    ...input.editor,
    ...input.markdown,
    ...input.settings,
    ...input.workspace,
    ...paths,
  });
  const tasks = useTaskRunner(input.editor);

  return {
    ...sidebar,
    ...paths,
    ...tree,
    ...metrics,
    ...tasks,
  };
}
