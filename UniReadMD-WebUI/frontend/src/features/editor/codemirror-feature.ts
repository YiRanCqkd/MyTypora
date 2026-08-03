import type { useMarkdownFeature } from '../markdown/markdown-feature';
import type { useSettingsFoundation } from '../settings/settings-foundation';
import type { useWorkspacePresentation } from '../workspace/workspace-presentation';
import { useCodeMirrorEditor } from './codemirror-editor';
import type { useEditorNavigation } from './editor-navigation';
import type { useEditorPresentation } from './editor-presentation';
import type { useSourceEditingFeature } from './source-editing-feature';
import type { useWorkbenchFoundation } from './workbench-foundation';

type CodeMirrorFeatureDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  markdown: ReturnType<typeof useMarkdownFeature>;
  navigation: ReturnType<typeof useEditorNavigation>;
  presentation: ReturnType<typeof useEditorPresentation>;
  settings: ReturnType<typeof useSettingsFoundation>;
  source: ReturnType<typeof useSourceEditingFeature>;
  workspace: ReturnType<typeof useWorkspacePresentation>;
}>;

export function useCodeMirrorFeature(input: CodeMirrorFeatureDependencies) {
  return useCodeMirrorEditor({
    ...input.editor,
    ...input.markdown,
    ...input.navigation,
    ...input.presentation,
    ...input.settings,
    ...input.source,
    ...input.workspace,
  });
}
