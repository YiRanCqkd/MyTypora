import type { useExportFeature } from '../export/export-feature';
import type { useSettingsActions } from '../settings/settings-actions';
import type { useSettingsFoundation } from '../settings/settings-foundation';
import type { useSettingsSession } from '../settings/settings-session';
import type { useWorkspaceActions } from '../workspace/workspace-actions';
import type { useWorkspacePresentation } from '../workspace/workspace-presentation';
import type { useEditorNavigation } from './editor-navigation';
import type { useEditorPresentation } from './editor-presentation';
import { useDocumentSearchActions } from './document-search-actions';
import { useKeyboardCommands } from './keyboard-commands';
import type { usePreviewEditingFeature } from './preview-editing-feature';
import { usePreviewWheel } from './preview-wheel';
import type { useSourceEditingFeature } from './source-editing-feature';
import type { useTableEditingCore } from './table-editing-core';
import type { useWorkbenchFoundation } from './workbench-foundation';

type EditorCommandsDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  exporting: ReturnType<typeof useExportFeature>;
  navigation: ReturnType<typeof useEditorNavigation>;
  presentation: ReturnType<typeof useEditorPresentation>;
  preview: ReturnType<typeof usePreviewEditingFeature>;
  settings: ReturnType<typeof useSettingsFoundation>;
  settingsActions: ReturnType<typeof useSettingsActions>;
  settingsSession: ReturnType<typeof useSettingsSession>;
  source: ReturnType<typeof useSourceEditingFeature>;
  table: ReturnType<typeof useTableEditingCore>;
  workspaceActions: ReturnType<typeof useWorkspaceActions>;
  workspacePresentation: ReturnType<typeof useWorkspacePresentation>;
}>;

export function useEditorCommands(input: EditorCommandsDependencies) {
  const search = useDocumentSearchActions({
    ...input.editor,
    ...input.navigation,
    ...input.presentation,
    ...input.preview,
    ...input.settings,
  });
  const wheel = usePreviewWheel({
    ...input.editor,
    ...input.presentation,
    ...input.settingsSession,
  });
  const keyboard = useKeyboardCommands({
    ...input.editor,
    ...input.exporting,
    ...input.presentation,
    ...input.preview,
    ...input.settingsActions,
    ...input.source,
    ...input.table,
    ...input.workspaceActions,
    ...search,
  });

  return {
    ...search,
    ...wheel,
    ...keyboard,
  };
}
