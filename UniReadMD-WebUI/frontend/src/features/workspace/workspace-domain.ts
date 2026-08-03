import type { ResolveWorkbenchDomains } from '../editor/workbench-domains';
import { createDeferredFeature } from '../editor/workbench-domains';
import { useWorkspaceActions } from './workspace-actions';
import { useWorkspaceFoundation } from './workspace-foundation';
import { useWorkspacePresentation } from './workspace-presentation';

export function useWorkspaceDomain(resolve: ResolveWorkbenchDomains) {
  const foundation = createDeferredFeature(() => {
    return useWorkspaceFoundation({ markdown: resolve().markdown.feature() });
  });
  const presentation = createDeferredFeature(() => {
    const domains = resolve();
    return useWorkspacePresentation({
      editor: domains.editor.foundation,
      editorPresentation: domains.editor.presentation(),
      markdown: domains.markdown.feature(),
      settings: domains.settings.foundation(),
      workspace: foundation(),
    });
  });
  const actions = createDeferredFeature(() => {
    const domains = resolve();
    return useWorkspaceActions({
      editor: domains.editor.foundation,
      navigation: domains.editor.navigation(),
      settings: domains.settings.foundation(),
      settingsSession: domains.settings.session(),
      workspace: foundation(),
      workspacePresentation: presentation(),
    });
  });

  return { actions, foundation, presentation };
}
