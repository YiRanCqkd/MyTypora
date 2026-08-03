import type { ResolveWorkbenchDomains } from '../editor/workbench-domains';
import { createDeferredFeature } from '../editor/workbench-domains';
import { useExportFeature } from './export-feature';
import { useExportFoundation } from './export-foundation';

export function useExportDomain(resolve: ResolveWorkbenchDomains) {
  const foundation = createDeferredFeature(() => useExportFoundation());
  const feature = createDeferredFeature(() => {
    const domains = resolve();
    return useExportFeature({
      editor: domains.editor.foundation,
      exportState: foundation(),
      presentation: domains.editor.presentation(),
      workspaceActions: domains.workspace.actions(),
      workspacePresentation: domains.workspace.presentation(),
    });
  });

  return { feature, foundation };
}
