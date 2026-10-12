import type { SessionSnapshot } from '../workspace/workspace-contracts';
import { bridgeService } from '../../services/bridge';
import { createTraceId, logBridge } from '../../utils/trace';
import type { createMarkdownRenderer } from '../markdown/renderer';
import type { usePreferences } from './preferences';
import type { useDocumentSearchState } from '../editor/document-search-state';
import type { useWorkspaceState } from '../workspace/workspace-state';
import type { useViewOptions } from './view-options';
import type { createLayoutMetrics } from '../workspace/layout-metrics';
import type { useRecentFiles } from '../workspace/recent-files';
import type { useTaskRunner } from '../workspace/task-runner';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';
import type { useWorkbenchStore } from '../editor/editor-store';
import type { useApplicationStatus } from '../editor/application-status';
type UseSessionPersistenceDependencies =
  ReturnType<typeof useApplicationStatus> &
  ReturnType<typeof createMarkdownRenderer> &
  ReturnType<typeof usePreferences> &
  ReturnType<typeof useDocumentSearchState> &
  ReturnType<typeof useWorkspaceState> &
  ReturnType<typeof useViewOptions> &
  ReturnType<typeof createLayoutMetrics> &
  ReturnType<typeof useRecentFiles> &
  ReturnType<typeof useTaskRunner> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof useWorkbenchStore>;

export function useSessionPersistence(input: UseSessionPersistenceDependencies) {
  let pendingSnapshot: SessionSnapshot | null = null;
  let pendingWrite: Promise<void> | null = null;
  const { SESSION_STORAGE_KEY,contentZoom,docReplaceQuery,docSearchCaseSensitive,
    docSearchMatchIndex,docSearchQuery,docSearchWholeWord,fileSortMode,fileTreeExpanded,
    fileViewMode,filesFilter,focusMode,isNativeWindow,mainPreviewWidth,normalizeRecentFiles,
    pinOutline,recentFiles,runtime,showDocReplacePanel,showDocSearchPanel,sidebarWidth,
    sourceMode,store,typewriterMode,
    workspaceDirectory,
    docSearchRegex,
    errorMessage,
  } = input;

  function buildSessionSnapshot(): SessionSnapshot {
    return {
      filePath: store.filePath,
      workspaceDirectory: workspaceDirectory.value,
      version: store.version,
      content: store.content,
      isDirty: store.isDirty,
      cursorHint: store.cursorHint,
      // IPC 只接收普通数据，同时避免排队快照随响应式状态继续变化。
      recentFiles: [...recentFiles.value],
      contentZoom: contentZoom.value,
      pinOutline: pinOutline.value,
      sourceMode: sourceMode.value,
      focusMode: focusMode.value,
      typewriterMode: typewriterMode.value,
      sidebarWidth: sidebarWidth.value,
      mainPreviewWidth: mainPreviewWidth.value,
      fileViewMode: fileViewMode.value,
      fileSortMode: fileSortMode.value,
      filesFilter: filesFilter.value,
      searchPanelVisible: showDocSearchPanel.value,
      replacePanelVisible: showDocReplacePanel.value,
      searchQuery: docSearchQuery.value,
      replaceQuery: docReplaceQuery.value,
      searchCaseSensitive: docSearchCaseSensitive.value,
      searchWholeWord: docSearchWholeWord.value,
      searchRegex: docSearchRegex?.value || false,
      searchMatchIndex: docSearchMatchIndex.value,
      fileTreeExpanded: { ...fileTreeExpanded.value },
    };
  }

  function persistSessionSnapshot() {
    const snapshot = buildSessionSnapshot();
    if (window.nativeBridge?.session) {
      pendingSnapshot = snapshot;
      if (!pendingWrite) {
        pendingWrite = (async () => {
          while (pendingSnapshot) {
            const current = pendingSnapshot;
            pendingSnapshot = null;
            await window.nativeBridge.session?.save(current);
          }
        })().catch((error) => {
          errorMessage.value = `恢复快照保存失败，请手动保存文档：${String(error)}`;
        }).finally(() => { pendingWrite = null; });
      }
      return pendingWrite;
    }
    try {
      localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(snapshot));
    } catch (error) {
      errorMessage.value = `恢复快照超过浏览器存储限制，请保存文档：${String(error)}`;
    }
  }

  function pushRecentFile(filePath: string) {
    if (!filePath) {
      return;
    }

    const filtered = recentFiles.value.filter((item) => item !== filePath);
    const nextRecentFiles = normalizeRecentFiles([filePath, ...filtered]);
    recentFiles.value = nextRecentFiles;

    if (isNativeWindow.value) {
      void bridgeService.setRecentFiles(nextRecentFiles, createTraceId('recent-files'))
        .then((result) => {
          recentFiles.value = normalizeRecentFiles(result.filePaths);
        })
        .catch(async (error) => {
          const message = error instanceof Error ? error.message : String(error);
          await logBridge('error', 'recent-files-sync:fail', createTraceId('recent-files'), {
            message,
            filePath,
          });
        });
      return;
    }

    persistSessionSnapshot();
  }

  async function loadSessionSnapshot() {
    if (window.nativeBridge?.session) {
      const snapshot = await window.nativeBridge.session.load();
      return snapshot as SessionSnapshot | null;
    }
    const raw = localStorage.getItem(SESSION_STORAGE_KEY);
    if (!raw) {
      return null;
    }

    try {
      const parsed = JSON.parse(raw) as SessionSnapshot;
      if (typeof parsed.content !== 'string') {
        return null;
      }

      return parsed;
    } catch {
      return null;
    }
  }

  async function loadRecentFilesFromNative(snapshotRecentFiles: string[]) {
    const fallbackRecentFiles = normalizeRecentFiles(snapshotRecentFiles);

    try {
      const traceId = createTraceId('recent-files-load');
      const result = await bridgeService.getRecentFiles(traceId);
      const nativeRecentFiles = normalizeRecentFiles(result.filePaths);
      if (nativeRecentFiles.length > 0) {
        recentFiles.value = nativeRecentFiles;
        return;
      }

      if (fallbackRecentFiles.length > 0) {
        recentFiles.value = fallbackRecentFiles;
        const saved = await bridgeService.setRecentFiles(fallbackRecentFiles, traceId);
        recentFiles.value = normalizeRecentFiles(saved.filePaths);
        await logBridge('info', 'recent-files-migrated', traceId, {
          count: recentFiles.value.length,
        });
        return;
      }

      recentFiles.value = [];
    } catch (error) {
      recentFiles.value = fallbackRecentFiles;
      const message = error instanceof Error ? error.message : String(error);
      await logBridge('error', 'recent-files-load:fail', createTraceId('recent-files-load'), {
        message,
      });
    }
  }

  function clearContentZoomPersistTimer() {
    if (!runtime.contentZoomPersistTimer) {
      return;
    }

    window.clearTimeout(runtime.contentZoomPersistTimer);
    runtime.contentZoomPersistTimer = 0;
  }

  return { buildSessionSnapshot,persistSessionSnapshot,
    pushRecentFile,loadSessionSnapshot,loadRecentFilesFromNative,clearContentZoomPersistTimer,
  };
}
