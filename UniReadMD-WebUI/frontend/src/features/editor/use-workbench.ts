import { useExportDomain } from '../export/export-domain';
import { useMarkdownDomain } from '../markdown/markdown-domain';
import { useSettingsDomain } from '../settings/settings-domain';
import { useWorkspaceDomain } from '../workspace/workspace-domain';
import { useEditorDomain } from './editor-domain';
import { createWorkbenchDomainResolver } from './workbench-domains';
import { createWorkbenchView } from './workbench-view';

export function useWorkbench() {
  const resolver = createWorkbenchDomainResolver();
  const editor = useEditorDomain(resolver.resolve);
  const markdown = useMarkdownDomain(resolver.resolve);
  const settings = useSettingsDomain(resolver.resolve);
  const workspace = useWorkspaceDomain(resolver.resolve);
  const exporting = useExportDomain(resolver.resolve);

  resolver.initialize({ editor, export: exporting, markdown, settings, workspace });
  return createWorkbenchView(resolver.resolve());
}
