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
type UseSessionPersistenceDependencies =
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
  const { SESSION_STORAGE_KEY,contentZoom,docReplaceQuery,docSearchCaseSensitive,
    docSearchMatchIndex,docSearchQuery,docSearchWholeWord,fileSortMode,fileTreeExpanded,
    fileViewMode,filesFilter,focusMode,isNativeWindow,mainPreviewWidth,normalizeRecentFiles,
    pinOutline,recentFiles,runTask,runtime,showDocReplacePanel,showDocSearchPanel,sidebarWidth,
    sourceMode,store,typewriterMode,userCssEnabled,userCssPath,
  } = input;

  async function normalizeLegacyUserCssSettings() {
    if (!userCssEnabled.value && !userCssPath.value) {
      return;
    }

    await runTask('settings-usercss-disable', async (traceId) => {
      await bridgeService.setSetting('userCssEnabled', false, traceId);
      await bridgeService.setSetting('userCssPath', '', traceId);
      return {
        ok: true,
      };
    }, {
      fallbackHint: '若样式仍异常，请删除用户目录中的 settings.json 后重启。',
    });

    userCssEnabled.value = false;
    userCssPath.value = '';
  }

  function buildSessionSnapshot(): SessionSnapshot {
    return {
      filePath: store.filePath,
      content: store.content,
      isDirty: store.isDirty,
      cursorHint: store.cursorHint,
      recentFiles: recentFiles.value,
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
      searchMatchIndex: docSearchMatchIndex.value,
      fileTreeExpanded: fileTreeExpanded.value,
    };
  }

  function persistSessionSnapshot() {
    const snapshot = buildSessionSnapshot();
    localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(snapshot));
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

  function loadSessionSnapshot() {
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

  return { normalizeLegacyUserCssSettings,buildSessionSnapshot,persistSessionSnapshot,
    pushRecentFile,loadSessionSnapshot,loadRecentFilesFromNative,clearContentZoomPersistTimer,
  };
}
