import type { useMarkdownFeature } from '../markdown/markdown-feature';
import { useDocumentLoadingState } from './document-loading-state';
import { useRecentFiles } from './recent-files';
import { useSidebarTabs } from './sidebar-tabs';
import { useWorkspaceState } from './workspace-state';

type WorkspaceFoundationDependencies = Readonly<{
  markdown: ReturnType<typeof useMarkdownFeature>;
}>;

export function useWorkspaceFoundation(input: WorkspaceFoundationDependencies) {
  const tabs = useSidebarTabs();
  const recentFiles = useRecentFiles();
  const state = useWorkspaceState(input.markdown);
  const loading = useDocumentLoadingState();

  return {
    ...tabs,
    ...recentFiles,
    ...state,
    ...loading,
  };
}
