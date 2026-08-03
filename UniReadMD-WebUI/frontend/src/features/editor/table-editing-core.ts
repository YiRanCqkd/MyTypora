import type { useEditorNavigation } from './editor-navigation';
import type { useEditorPresentation } from './editor-presentation';
import type { usePreviewEditingFeature } from './preview-editing-feature';
import { useTableCellEditing } from './table-cell-editing';
import { createTableMutations } from './table-mutations';
import type { useWorkbenchFoundation } from './workbench-foundation';

type TableEditingCoreDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  navigation: ReturnType<typeof useEditorNavigation>;
  presentation: ReturnType<typeof useEditorPresentation>;
  preview: ReturnType<typeof usePreviewEditingFeature>;
}>;

export function useTableEditingCore(input: TableEditingCoreDependencies) {
  const cells = useTableCellEditing({
    ...input.editor,
    ...input.navigation,
    ...input.presentation,
  });
  const mutations = createTableMutations({
    ...input.editor,
    ...input.navigation,
    ...input.presentation,
    ...cells,
  });

  return {
    ...cells,
    ...mutations,
  };
}
