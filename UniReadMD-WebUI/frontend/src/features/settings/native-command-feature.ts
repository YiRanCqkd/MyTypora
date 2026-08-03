import type { useSourceEditingFeature } from '../editor/source-editing-feature';
import type { useExportFeature } from '../export/export-feature';
import type { useDiagramFeature } from '../markdown/diagram-feature';
import type { useWorkspaceActions } from '../workspace/workspace-actions';
import type { useWorkspacePresentation } from '../workspace/workspace-presentation';
import { useNativeCommands } from './native-commands';
import type { useSettingsActions } from './settings-actions';
import type { useSettingsFoundation } from './settings-foundation';

type NativeCommandFeatureDependencies = Readonly<{
  diagrams: ReturnType<typeof useDiagramFeature>;
  exporting: ReturnType<typeof useExportFeature>;
  settings: ReturnType<typeof useSettingsFoundation>;
  settingsActions: ReturnType<typeof useSettingsActions>;
  source: ReturnType<typeof useSourceEditingFeature>;
  workspaceActions: ReturnType<typeof useWorkspaceActions>;
  workspacePresentation: ReturnType<typeof useWorkspacePresentation>;
}>;

export function useNativeCommandFeature(input: NativeCommandFeatureDependencies) {
  return useNativeCommands({
    ...input.diagrams,
    ...input.exporting,
    ...input.settings,
    ...input.settingsActions,
    ...input.source,
    ...input.workspaceActions,
    ...input.workspacePresentation,
  });
}
