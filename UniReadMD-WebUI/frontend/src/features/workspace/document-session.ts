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

export function useDocumentSession(input: UseDocumentSessionDependencies) {
  const { activeTab,pushRecentFile,runDocumentLoadingTask,runTask,store,
  } = input;

  async function openFile() {
    const result = await runDocumentLoadingTask('open-file', (traceId) => {
      return bridgeService.openMarkdown(traceId);
    });

    if (!result) {
      return;
    }

    store.setFile(result.filePath, result.content);
    pushRecentFile(result.filePath);
    activeTab.value = 'outline';
  }

  async function saveFile() {
    const result = await runTask('save-file', (traceId) => {
      return bridgeService.saveMarkdown({
        filePath: store.filePath || undefined,
        content: store.content,
        traceId,
      });
    });

    if (!result) {
      return;
    }

    store.markSaved(result.filePath);
    pushRecentFile(result.filePath);
  }

  async function saveAsFile() {
    const result = await runTask('save-as-file', (traceId) => {
      return bridgeService.saveMarkdown({
        content: store.content,
        traceId,
      });
    });

    if (!result) {
      return;
    }

    store.markSaved(result.filePath);
    pushRecentFile(result.filePath);
  }

  return { openFile,saveFile,saveAsFile,
  };
}
