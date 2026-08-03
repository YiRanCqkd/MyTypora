import type { ResolveWorkbenchDomains } from '../editor/workbench-domains';
import { createDeferredFeature } from '../editor/workbench-domains';
import { useNativeCommandFeature } from './native-command-feature';
import { useSettingsActions } from './settings-actions';
import { useSettingsFoundation } from './settings-foundation';
import { useSettingsSession } from './settings-session';

export function useSettingsDomain(resolve: ResolveWorkbenchDomains) {
  const foundation = createDeferredFeature(() => {
    return useSettingsFoundation({ markdown: resolve().markdown.feature() });
  });
  const session = createDeferredFeature(() => {
    const domains = resolve();
    return useSettingsSession({
      editor: domains.editor.foundation,
      editorPresentation: domains.editor.presentation(),
      markdown: domains.markdown.feature(),
      settings: foundation(),
      workspace: domains.workspace.foundation(),
      workspacePresentation: domains.workspace.presentation(),
    });
  });
  const actions = createDeferredFeature(() => {
    const domains = resolve();
    return useSettingsActions({
      codeMirror: domains.editor.codeMirror(),
      editor: domains.editor.foundation,
      navigation: domains.editor.navigation(),
      settings: foundation(),
      workspace: domains.workspace.foundation(),
      workspacePresentation: domains.workspace.presentation(),
    });
  });
  const nativeCommands = createDeferredFeature(() => {
    const domains = resolve();
    return useNativeCommandFeature({
      diagrams: domains.markdown.diagrams(),
      exporting: domains.export.feature(),
      settings: foundation(),
      settingsActions: actions(),
      source: domains.editor.source(),
      workspaceActions: domains.workspace.actions(),
      workspacePresentation: domains.workspace.presentation(),
    });
  });

  return { actions, foundation, nativeCommands, session };
}
