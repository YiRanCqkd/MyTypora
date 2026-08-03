import type { ResolveWorkbenchDomains } from '../editor/workbench-domains';
import { createDeferredFeature } from '../editor/workbench-domains';
import { useDiagramFeature } from './diagram-feature';
import { useMarkdownFeature } from './markdown-feature';

export function useMarkdownDomain(resolve: ResolveWorkbenchDomains) {
  const feature = createDeferredFeature(() => {
    const domains = resolve();
    return useMarkdownFeature({ editor: domains.editor.foundation });
  });
  const diagrams = createDeferredFeature(() => {
    const domains = resolve();
    return useDiagramFeature({
      editor: domains.editor.foundation,
      editorPresentation: domains.editor.presentation(),
      markdown: feature(),
      settings: domains.settings.foundation(),
    });
  });

  return { diagrams, feature };
}
