import type { useSettingsFoundation } from '../settings/settings-foundation';
import type { useEditorNavigation } from './editor-navigation';
import type { useEditorPresentation } from './editor-presentation';
import type { usePreviewEditingFeature } from './preview-editing-feature';
import { useContextMenuActions } from './context-menu-actions';
import { useSourceEditing } from './source-editing';
import type { useTableEditingCore } from './table-editing-core';
import type { useWorkbenchFoundation } from './workbench-foundation';

type SourceEditingFeatureDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  navigation: ReturnType<typeof useEditorNavigation>;
  presentation: ReturnType<typeof useEditorPresentation>;
  preview: ReturnType<typeof usePreviewEditingFeature>;
  settings: ReturnType<typeof useSettingsFoundation>;
  table: ReturnType<typeof useTableEditingCore>;
}>;

export function useSourceEditingFeature(input: SourceEditingFeatureDependencies) {
  const source = useSourceEditing({
    ...input.editor,
    ...input.navigation,
    ...input.presentation,
    ...input.preview,
    ...input.table,
  });
  const menus = useContextMenuActions({
    ...input.editor,
    ...input.presentation,
    ...input.preview,
    ...input.settings,
    ...input.table,
    ...source,
  });

  return {
    ...source,
    ...menus,
  };
}
