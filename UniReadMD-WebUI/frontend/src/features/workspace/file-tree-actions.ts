import type { FileTreeRow } from './workspace-contracts';
import { bridgeService } from '../../services/bridge';
import { nextTick } from 'vue';
import type { useSidebarTabs } from './sidebar-tabs';
import type { createLayoutMetrics } from './layout-metrics';
import type { useWorkspaceState } from './workspace-state';
import type { useDocumentLoadingState } from './document-loading-state';
import type { useViewOptions } from '../settings/view-options';
import type { useSessionPersistence } from '../settings/session-persistence';
import type { useTaskRunner } from './task-runner';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';
import type { useWorkbenchStore } from '../editor/editor-store';
type UseFileTreeActionsDependencies =
  ReturnType<typeof useSidebarTabs> &
  ReturnType<typeof createLayoutMetrics> &
  ReturnType<typeof useWorkspaceState> &
  ReturnType<typeof useDocumentLoadingState> &
  ReturnType<typeof useViewOptions> &
  ReturnType<typeof useSessionPersistence> &
  ReturnType<typeof useTaskRunner> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof useWorkbenchStore>;
import type { useDocumentIntegrity } from './document-integrity';

export function useFileTreeActions(input: UseFileTreeActionsDependencies
  & ReturnType<typeof useDocumentIntegrity>) {
  const { activeTab,ensureMainPreviewWidth,fileTreeExpanded,isDocumentLoading,
    isResizingMainPreview,isResizingSidebar,mainPreviewWidth,normalizeMainPreviewWidth,
    normalizeSidebarWidth,pinOutline,pushRecentFile,resolveMainPreviewContainerWidth,runTask,
    runtime,sidebarWidth,sourceMode,store,ensureCanReplaceDocument,
  } = input;

  function startSidebarResize(event: MouseEvent) {
    if (!pinOutline.value) {
      return;
    }

    event.preventDefault();

    const startX = event.clientX;
    const startWidth = sidebarWidth.value;
    isResizingSidebar.value = true;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.clientX - startX;
      sidebarWidth.value = normalizeSidebarWidth(startWidth + deltaX);
    };

    const stopResize = () => {
      isResizingSidebar.value = false;
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', stopResize);
      runtime.sidebarResizeCleanup = null;
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', stopResize);

    runtime.sidebarResizeCleanup = stopResize;
  }

  function startMainPreviewResize(event: MouseEvent) {
    if (!sourceMode.value) {
      return;
    }

    event.preventDefault();

    const containerWidth = resolveMainPreviewContainerWidth();
    if (containerWidth <= 0) {
      return;
    }

    if (mainPreviewWidth.value <= 0) {
      ensureMainPreviewWidth();
    }

    const startX = event.clientX;
    const startWidth = mainPreviewWidth.value;
    isResizingMainPreview.value = true;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.clientX - startX;
      mainPreviewWidth.value = normalizeMainPreviewWidth(startWidth + deltaX, containerWidth);
    };

    const stopResize = () => {
      isResizingMainPreview.value = false;
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', stopResize);
      runtime.mainPreviewResizeCleanup = null;
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', stopResize);
    runtime.mainPreviewResizeCleanup = stopResize;
  }

  function toggleTreeRow(row: FileTreeRow) {
    if (!row.expandable) {
      return;
    }

    fileTreeExpanded.value = {
      ...fileTreeExpanded.value,
      [row.id]: !row.expanded,
    };
  }

  async function runDocumentLoadingTask<T>(
    taskName: string,
    executor: (traceId: string) => Promise<T>,
    options?: {
      fallbackHint?: string;
    },
  ): Promise<T | null> {
    if (isDocumentLoading.value) return null;
    isDocumentLoading.value = true;
    await nextTick();

    try {
      if (!await ensureCanReplaceDocument()) return null;
      return await runTask(taskName, executor, options);
    } finally {
      isDocumentLoading.value = false;
    }
  }

  async function openRecentFile(filePath: string) {
    const result = await runDocumentLoadingTask('open-recent-file', (traceId) => {
      return bridgeService.openFromPath(filePath, traceId);
    });

    if (!result) {
      return;
    }

    store.setFile(result.filePath, result.content, result);
    pushRecentFile(result.filePath);
    activeTab.value = 'outline';
  }

  function openFileFromList(path: string) {
    void openRecentFile(path);
  }

  function handleFileTreeContextMenu(event: MouseEvent, row: FileTreeRow) {
    if (row.type !== 'file' || !row.filePath) {
      return;
    }

    const filePath = row.filePath;
    const x = Number.isFinite(event.clientX) ? Math.round(event.clientX) : undefined;
    const y = Number.isFinite(event.clientY) ? Math.round(event.clientY) : undefined;

    void runTask('file-tree-context-menu', (traceId) => {
      return bridgeService.showFileTreeContextMenu({
        filePath,
        x,
        y,
        traceId,
      });
    });
  }

  function isFileDragEvent(event: DragEvent) {
    const types = event.dataTransfer?.types;
    if (!types) {
      return false;
    }

    return Array.from(types).includes('Files');
  }

  return { startSidebarResize,startMainPreviewResize,toggleTreeRow,
    runDocumentLoadingTask,openRecentFile,openFileFromList,handleFileTreeContextMenu,
    isFileDragEvent,
  };
}
