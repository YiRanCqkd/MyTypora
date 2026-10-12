import { bridgeService } from '../../services/bridge';
import type { useSidebarTabs } from './sidebar-tabs';
import type { useSessionPersistence } from '../settings/session-persistence';
import type { useFileTreeActions } from './file-tree-actions';
import type { useTaskRunner } from './task-runner';
import type { useWorkbenchStore } from '../editor/editor-store';
type UseDocumentSessionDependencies =
  ReturnType<typeof useSidebarTabs> &
  ReturnType<typeof useSessionPersistence> &
  ReturnType<typeof useFileTreeActions> &
  ReturnType<typeof useTaskRunner> &
  ReturnType<typeof useWorkbenchStore>;
import type { useDocumentIntegrity } from './document-integrity';
import type { useWorkspaceState } from './workspace-state';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';

export function useDocumentSession(input: UseDocumentSessionDependencies
  & ReturnType<typeof useDocumentIntegrity> & ReturnType<typeof useWorkspaceState>
  & ReturnType<typeof createWorkbenchRuntime>) {
  const { activeTab,pushRecentFile,runDocumentLoadingTask,store,saveCurrentDocument,
    ensureCanReplaceDocument,currentDirectoryPath,currentDirectoryMarkdownFiles,
    workspaceDirectory,workspaceQuery,workspaceMatches,runTask,fileModifiedTimes,
    workspaceDirectories,
    workspaceCaseSensitive,workspaceWholeWord,workspaceRegex,workspaceSearchTruncated,
  } = input;

  async function openFile() {
    const result = await runDocumentLoadingTask('open-file', (traceId) => {
      return bridgeService.openMarkdown(traceId);
    });

    if (!result) {
      return;
    }

    store.setFile(result.filePath, result.content, result);
    pushRecentFile(result.filePath);
    activeTab.value = 'outline';
  }

  async function saveFile() {
    if (await saveCurrentDocument()) pushRecentFile(store.filePath);
  }

  async function saveAsFile() {
    if (await saveCurrentDocument(true)) pushRecentFile(store.filePath);
  }

  async function newFile() {
    await runDocumentLoadingTask('新建文档', async () => {
      store.setFile('', '');
      return true;
    });
  }

  async function openFolder() {
    const result = await runTask('打开文件夹', async () => {
      return await window.nativeBridge.file.openDirectory?.({}) || null;
    });
    if (!result) return;
    currentDirectoryPath.value = result.directoryPath;
    workspaceDirectory.value = result.directoryPath;
    fileModifiedTimes.value = result.modifiedTimes || {};
    currentDirectoryMarkdownFiles.value = result.filePaths;
    workspaceDirectories.value = result.directories || [];
    activeTab.value = 'files';
  }

  async function searchWorkspace() {
    if (!workspaceQuery.value) {
      workspaceMatches.value = [];
      return;
    }
    const result = await runTask('搜索文件夹正文', async () => {
      return await window.nativeBridge.file.searchDirectory?.({ query: workspaceQuery.value,
        caseSensitive: workspaceCaseSensitive.value, wholeWord: workspaceWholeWord.value,
        regularExpression: workspaceRegex.value,
      })
        || { matches: [] };
    });
    workspaceMatches.value = result?.matches || [];
    workspaceSearchTruncated.value = Boolean(result && 'truncated' in result && result.truncated);
  }

  return { openFile,saveFile,saveAsFile,newFile,openFolder,ensureCanReplaceDocument,store,
    runtime: input.runtime,
    searchWorkspace,
  };
}
