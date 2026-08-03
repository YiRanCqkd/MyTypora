import type { useEditorPresentation } from './editor-presentation';
import { useEditorSynchronization } from './editor-synchronization';
import { createPreviewEditingRanges } from './preview-editing-ranges';
import { usePreviewNavigation } from './preview-navigation';
import type { useWorkbenchFoundation } from './workbench-foundation';

type EditorNavigationDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  presentation: ReturnType<typeof useEditorPresentation>;
}>;

export function useEditorNavigation(input: EditorNavigationDependencies) {
  const synchronization = useEditorSynchronization(input.editor);
  const preview = usePreviewNavigation({ ...input.editor, ...synchronization });
  const ranges = createPreviewEditingRanges({
    ...input.editor,
    ...input.presentation,
    ...preview,
  });

  return {
    ...synchronization,
    ...preview,
    ...ranges,
  };
}
