import type { useSettingsFoundation } from '../settings/settings-foundation';
import type { useEditorPresentation } from './editor-presentation';
import { usePreviewEditingEvents } from './preview-editing-events';
import type { usePreviewEditingFeature } from './preview-editing-feature';
import { usePreviewEditingShortcuts } from './preview-editing-shortcuts';
import type { useTableEditingCore } from './table-editing-core';
import type { useWorkbenchFoundation } from './workbench-foundation';
import type { useEditorNavigation } from './editor-navigation';

type PreviewEditingEventsDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  navigation: ReturnType<typeof useEditorNavigation>;
  presentation: ReturnType<typeof useEditorPresentation>;
  preview: ReturnType<typeof usePreviewEditingFeature>;
  settings: ReturnType<typeof useSettingsFoundation>;
  table: ReturnType<typeof useTableEditingCore>;
}>;

export function usePreviewEditingEventsFeature(input: PreviewEditingEventsDependencies) {
  const events = usePreviewEditingEvents({
    ...input.editor,
    ...input.navigation,
    ...input.presentation,
    ...input.preview,
    ...input.settings,
    ...input.table,
  });
  const shortcuts = usePreviewEditingShortcuts({
    ...input.presentation,
    ...input.preview,
    ...events,
  });

  return {
    ...events,
    ...shortcuts,
  };
}
