import type { useSettingsFoundation } from '../settings/settings-foundation';
import type { useWorkspacePresentation } from '../workspace/workspace-presentation';
import type { useEditorNavigation } from './editor-navigation';
import type { useEditorPresentation } from './editor-presentation';
import type { usePreviewEditingEventsFeature } from './preview-editing-events-feature';
import type { usePreviewEditingFeature } from './preview-editing-feature';
import type { useSourceEditingFeature } from './source-editing-feature';
import { useTableActions } from './table-actions';
import type { useTableEditingCore } from './table-editing-core';
import { useTableEditingEvents } from './table-editing-events';
import { useTableNavigation } from './table-navigation';
import type { useWorkbenchFoundation } from './workbench-foundation';

type TableEditingFeatureDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  events: ReturnType<typeof usePreviewEditingEventsFeature>;
  navigation: ReturnType<typeof useEditorNavigation>;
  presentation: ReturnType<typeof useEditorPresentation>;
  preview: ReturnType<typeof usePreviewEditingFeature>;
  settings: ReturnType<typeof useSettingsFoundation>;
  source: ReturnType<typeof useSourceEditingFeature>;
  table: ReturnType<typeof useTableEditingCore>;
  workspace: ReturnType<typeof useWorkspacePresentation>;
}>;

export function useTableEditingFeature(input: TableEditingFeatureDependencies) {
  const actions = useTableActions({
    ...input.editor,
    ...input.events,
    ...input.presentation,
    ...input.preview,
    ...input.settings,
    ...input.source,
    ...input.table,
  });
  const navigation = useTableNavigation({
    ...input.editor,
    ...input.events,
    ...input.navigation,
    ...input.presentation,
    ...input.table,
  });
  const events = useTableEditingEvents({
    ...input.editor,
    ...input.presentation,
    ...input.preview,
    ...input.settings,
    ...input.source,
    ...input.table,
    ...input.workspace,
    ...actions,
    ...navigation,
  });

  return {
    ...actions,
    ...navigation,
    ...events,
  };
}
