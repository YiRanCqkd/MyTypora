import type { useEditorPresentation } from '../editor/editor-presentation';
import type { useWorkbenchFoundation } from '../editor/workbench-foundation';
import type { useWorkspaceActions } from '../workspace/workspace-actions';
import type { useWorkspacePresentation } from '../workspace/workspace-presentation';
import { useDocumentExport } from './export-document';
import type { useExportFoundation } from './export-foundation';
import { useExportCapabilities } from './export-capabilities';
import type { useDiagramFeature } from '../markdown/diagram-feature';

type ExportFeatureDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  diagrams: ReturnType<typeof useDiagramFeature>;
  exportState: ReturnType<typeof useExportFoundation>;
  presentation: ReturnType<typeof useEditorPresentation>;
  workspaceActions: ReturnType<typeof useWorkspaceActions>;
  workspacePresentation: ReturnType<typeof useWorkspacePresentation>;
}>;

export function useExportFeature(input: ExportFeatureDependencies) {
  const document = useDocumentExport({
    ...input.editor,
    ...input.diagrams,
    ...input.exportState,
    ...input.presentation,
    ...input.workspaceActions,
    ...input.workspacePresentation,
  });
  const capabilities = useExportCapabilities({
    ...input.exportState,
    ...input.workspacePresentation,
    ...document,
  });

  return {
    ...document,
    ...capabilities,
  };
}
