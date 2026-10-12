import { bridgeService } from '../../services/bridge';
import { createTraceId, logBridge } from '../../utils/trace';
import type { useWorkspaceState } from './workspace-state';
import type { createFilePathModel } from './file-paths';
import type { useApplicationStatus } from '../editor/application-status';
import type { useDocumentLoadingState } from './document-loading-state';
import type { useFileTreeActions } from './file-tree-actions';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';
import type { useWorkbenchStore } from '../editor/editor-store';
type UseFileDropDependencies =
  ReturnType<typeof useWorkspaceState> &
  ReturnType<typeof createFilePathModel> &
  ReturnType<typeof useApplicationStatus> &
  ReturnType<typeof useDocumentLoadingState> &
  ReturnType<typeof useFileTreeActions> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof useWorkbenchStore>;

export function useFileDrop(input: UseFileDropDependencies) {
  const { currentDirectoryMarkdownFiles,currentDirectoryPath,dedupeMarkdownPaths,errorMessage,
    getParentPath,isDragOverFile,isFileDragEvent,isMarkdownPath,openRecentFile,runtime,
    workspaceDirectory,fileModifiedTimes,
    workspaceDirectories,store,runDocumentLoadingTask,
  } = input;

  function parseFileUriToPath(value: string) {
    try {
      const parsed = new URL(value);
      if (parsed.protocol !== 'file:') {
        return '';
      }

      const decodedPath = decodeURIComponent(parsed.pathname || '');
      if (/^\/[A-Za-z]:\//.test(decodedPath)) {
        return decodedPath.slice(1).replace(/\//g, '\\');
      }

      return decodedPath.replace(/\//g, '\\');
    } catch {
      return '';
    }
  }

  function resolveDroppedMarkdownPath(event: DragEvent) {
    const transfer = event.dataTransfer;
    if (!transfer) {
      return '';
    }

    const droppedFiles = Array.from(transfer.files || []) as Array<File & { path?: string }>;
    for (const fileItem of droppedFiles) {
      let filePath = typeof fileItem.path === 'string' ? fileItem.path : '';
      if (!filePath) {
        filePath = bridgeService.getPathForFile(fileItem);
      }

      if (filePath && isMarkdownPath(filePath)) {
        return filePath;
      }
    }

    const uriList = transfer.getData('text/uri-list');
    if (!uriList) {
      return '';
    }

    const uriItems = uriList
      .split(/\r?\n/)
      .map((item) => item.trim())
      .filter((item) => item && !item.startsWith('#'));

    for (const uriItem of uriItems) {
      const filePath = parseFileUriToPath(uriItem);
      if (filePath && isMarkdownPath(filePath)) {
        return filePath;
      }
    }

    return '';
  }

  function handleWindowDragOver(event: DragEvent) {
    if (!isFileDragEvent(event)) {
      return;
    }

    event.preventDefault();
    isDragOverFile.value = true;
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = 'copy';
    }
  }

  function handleWindowDragEnter(event: DragEvent) {
    if (!isFileDragEvent(event)) {
      return;
    }

    event.preventDefault();
    runtime.fileDragDepth += 1;
    isDragOverFile.value = true;
  }

  function handleWindowDragLeave(event: DragEvent) {
    if (!isFileDragEvent(event)) {
      return;
    }

    event.preventDefault();
    runtime.fileDragDepth = Math.max(0, runtime.fileDragDepth - 1);
    if (runtime.fileDragDepth === 0) {
      isDragOverFile.value = false;
    }
  }

  function handleWindowDrop(event: DragEvent) {
    if (!isFileDragEvent(event)) {
      return;
    }

    event.preventDefault();
    runtime.fileDragDepth = 0;
    isDragOverFile.value = false;
    const dropped = [...(event.dataTransfer?.files || [])]
      .find((file) => /\.(md|markdown)$/i.test(file.name));
    if (dropped && window.nativeBridge.file.openDroppedFile) {
      // 先消费真实拖放能力，再等待脏文档对话框，避免能力在等待期间过期。
      const opening = window.nativeBridge.file.openDroppedFile(dropped);
      // 给读取 Promise 立即安装拒绝处理，用户在确认框停留时不产生未处理异常。
      void opening.catch(() => {});
      void runDocumentLoadingTask('拖放打开文档', () => opening).then((result) => {
        if (result) store.setFile(result.filePath, result.content, result);
      });
      return;
    }

    const filePath = resolveDroppedMarkdownPath(event);
    if (!filePath) {
      errorMessage.value = '仅支持拖拽打开 Markdown 文件（.md/.markdown）。';
      return;
    }

    void openRecentFile(filePath);
  }

  async function refreshCurrentDirectoryFiles(filePath: string) {
    if (workspaceDirectory.value && window.nativeBridge.file.openDirectory) {
      const result = await window.nativeBridge.file.openDirectory({
        directoryPath: workspaceDirectory.value,
      }).catch(() => null);
      if (result && result.directoryPath === workspaceDirectory.value) {
        currentDirectoryPath.value = result.directoryPath;
        currentDirectoryMarkdownFiles.value = result.filePaths;
        fileModifiedTimes.value = result.modifiedTimes || {};
        workspaceDirectories.value = result.directories || [];
      }
      return;
    }
    if (!filePath) {
      currentDirectoryPath.value = '';
      currentDirectoryMarkdownFiles.value = [];
      return;
    }

    const fallbackPath = dedupeMarkdownPaths([filePath]);

    try {
      const traceId = createTraceId('list-dir');
      const result = await bridgeService.listMarkdownInDirectory(filePath, traceId);
      currentDirectoryPath.value = result.directoryPath;
      fileModifiedTimes.value = result.modifiedTimes || {};
      workspaceDirectories.value = result.directories || [];
      currentDirectoryMarkdownFiles.value = dedupeMarkdownPaths([
        ...result.filePaths,
        ...fallbackPath,
      ]);
    } catch (error) {
      currentDirectoryPath.value = getParentPath(filePath);
      currentDirectoryMarkdownFiles.value = fallbackPath;
      const message = error instanceof Error ? error.message : String(error);
      await logBridge('error', 'list-directory-markdown:fail', createTraceId('list-dir'), {
        message,
        filePath,
      });
    }
  }

  return { parseFileUriToPath,resolveDroppedMarkdownPath,handleWindowDragOver,
    handleWindowDragEnter,handleWindowDragLeave,handleWindowDrop,refreshCurrentDirectoryFiles,
  };
}
