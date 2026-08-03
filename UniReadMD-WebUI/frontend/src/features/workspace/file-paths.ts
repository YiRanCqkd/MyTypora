import type { FileListItem } from './workspace-contracts';
import { computed } from 'vue';
import type { useWorkspaceState } from './workspace-state';
import type { useRecentFiles } from './recent-files';
import type { useWorkbenchStore } from '../editor/editor-store';
type CreateFilePathModelDependencies =
  ReturnType<typeof useWorkspaceState> &
  ReturnType<typeof useRecentFiles> &
  ReturnType<typeof useWorkbenchStore>;

export function createFilePathModel(input: CreateFilePathModelDependencies) {
  const { currentDirectoryPath,fileSortMode,recentFiles,store,
  } = input;

  function normalizeSlashes(inputPath: string) {
    return inputPath.replace(/\\/g, '/');
  }

  function splitPath(inputPath: string) {
    return normalizeSlashes(inputPath).split('/').filter((item) => item.length > 0);
  }

  function getFileName(inputPath: string) {
    const parts = splitPath(inputPath);
    return parts[parts.length - 1] || inputPath;
  }

  function getParentPath(inputPath: string) {
    const parts = splitPath(inputPath);
    if (parts.length <= 1) {
      return '/';
    }

    return parts.slice(0, -1).join('/');
  }

  function getFileExtension(inputPath: string) {
    const name = getFileName(inputPath);
    const index = name.lastIndexOf('.');
    if (index < 0 || index === name.length - 1) {
      return '';
    }

    return name.slice(index + 1).toLowerCase();
  }

  function isMarkdownPath(inputPath: string) {
    const ext = getFileExtension(inputPath);
    return ext === 'md' || ext === 'markdown';
  }

  const fileListItems = computed<FileListItem[]>(() => {
    const unique = new Map<string, FileListItem>();
    const paths = store.filePath
      ? [store.filePath, ...recentFiles.value]
      : [...recentFiles.value];

    paths.forEach((path, index) => {
      if (!path || unique.has(path) || !isMarkdownPath(path)) {
        return;
      }

      unique.set(path, {
        path,
        name: getFileName(path),
        parent: getParentPath(path),
        ext: getFileExtension(path),
        order: index,
      });
    });

    let items = Array.from(unique.values());
    const sortMode = fileSortMode.value;
    if (sortMode === 'name') {
      items = [...items].sort((left, right) => {
        return left.name.localeCompare(right.name, 'zh-CN');
      });
    } else if (sortMode === 'date') {
      items = [...items].sort((left, right) => {
        return left.order - right.order;
      });
    } else {
      items = [...items].sort((left, right) => left.order - right.order);
    }

    return items;
  });

  function getDirectoryDisplayName(inputPath: string) {
    const normalized = normalizeSlashes(inputPath).trim();
    if (!normalized) {
      return '当前目录';
    }

    if (normalized === '/') {
      return '/';
    }

    const withoutTailSlash = normalized.replace(/\/+$/g, '');
    if (/^[A-Za-z]:$/.test(withoutTailSlash)) {
      return `${withoutTailSlash}/`;
    }

    if (!withoutTailSlash) {
      return '/';
    }

    const parts = withoutTailSlash.split('/').filter((item) => item.length > 0);
    if (parts.length === 0) {
      return '/';
    }

    return parts[parts.length - 1];
  }

  const currentDirectoryLabel = computed(() => {
    const fallbackPath = store.filePath ? getParentPath(store.filePath) : '';
    const directoryPath = currentDirectoryPath.value || fallbackPath;
    if (!directoryPath) {
      return '当前目录';
    }

    return getDirectoryDisplayName(directoryPath);
  });

  function dedupeMarkdownPaths(paths: string[]) {
    const seen = new Set<string>();
    const result: string[] = [];
    paths.forEach((item) => {
      if (!item || !isMarkdownPath(item)) {
        return;
      }

      const key = normalizeSlashes(item).toLowerCase();
      if (seen.has(key)) {
        return;
      }

      seen.add(key);
      result.push(item);
    });
    return result;
  }

  return { normalizeSlashes,splitPath,getFileName,getParentPath,getFileExtension,
    isMarkdownPath,fileListItems,getDirectoryDisplayName,currentDirectoryLabel,
    dedupeMarkdownPaths,
  };
}
