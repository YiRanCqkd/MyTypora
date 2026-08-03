import type { FileSortMode, FileViewMode } from './workspace-contracts';
import { ref } from 'vue';
import type { createMarkdownRenderer } from '../markdown/renderer';
type UseWorkspaceStateDependencies =
  ReturnType<typeof createMarkdownRenderer>;

export function useWorkspaceState(input: UseWorkspaceStateDependencies) {
  const { SIDEBAR_WIDTH_DEFAULT,
  } = input;

  const sidebarWidth = ref(SIDEBAR_WIDTH_DEFAULT);

  const isResizingSidebar = ref(false);

  const mainPreviewWidth = ref(0);

  const isResizingMainPreview = ref(false);

  const fileViewMode = ref<FileViewMode>('tree');

  const fileSortMode = ref<FileSortMode>('natural');

  const filesFilter = ref('');

  const searchText = ref('');

  const fileTreeExpanded = ref<Record<string, boolean>>({});

  const currentDirectoryPath = ref('');

  const currentDirectoryMarkdownFiles = ref<string[]>([]);

  return { sidebarWidth,isResizingSidebar,mainPreviewWidth,isResizingMainPreview,
    fileViewMode,fileSortMode,filesFilter,searchText,fileTreeExpanded,currentDirectoryPath,
    currentDirectoryMarkdownFiles,
  };
}
