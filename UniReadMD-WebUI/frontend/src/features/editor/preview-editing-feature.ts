import type { useMarkdownFeature } from '../markdown/markdown-feature';
import type { useSettingsFoundation } from '../settings/settings-foundation';
import type { useEditorNavigation } from './editor-navigation';
import type { useEditorPresentation } from './editor-presentation';
import { usePreviewEditingSession } from './preview-editing-session';
import { usePreviewInteractions } from './preview-interactions';
import { createPreviewSelection } from './preview-selection';
import type { useWorkbenchFoundation } from './workbench-foundation';

type PreviewEditingFeatureDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  markdown: ReturnType<typeof useMarkdownFeature>;
  navigation: ReturnType<typeof useEditorNavigation>;
  presentation: ReturnType<typeof useEditorPresentation>;
  settings: ReturnType<typeof useSettingsFoundation>;
}>;

export function usePreviewEditingFeature(input: PreviewEditingFeatureDependencies) {
  const session = usePreviewEditingSession({
    ...input.editor,
    ...input.markdown,
    ...input.navigation,
    ...input.presentation,
  });
  const selection = createPreviewSelection(input.editor);
  const interactions = usePreviewInteractions({
    ...input.editor,
    ...input.navigation,
    ...input.presentation,
    ...input.settings,
    ...session,
  });

  return {
    ...session,
    ...selection,
    ...interactions,
  };
}
