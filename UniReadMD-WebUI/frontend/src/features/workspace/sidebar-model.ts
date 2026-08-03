import { computed } from 'vue';
import type { useSidebarTabs } from './sidebar-tabs';
import type { useWorkspaceState } from './workspace-state';
import type { usePreviewEditingState } from '../editor/preview-editing-state';
import type { useWorkbenchStore } from '../editor/editor-store';
type UseSidebarModelDependencies =
  ReturnType<typeof useSidebarTabs> &
  ReturnType<typeof useWorkspaceState> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof useWorkbenchStore>;

export function useSidebarModel(input: UseSidebarModelDependencies) {
  const { activeTab,filesFilter,headings,searchText,store,
  } = input;

  const sidebarSearchPlaceholder = computed(() => {
    if (activeTab.value === 'files') {
      return '匹配文档名（当前目录 / 最近打开）';
    }

    return '匹配当前文档标题';
  });

  const sidebarSearchText = computed({
    get() {
      if (activeTab.value === 'files') {
        return filesFilter.value;
      }

      return searchText.value;
    },
    set(value: string) {
      if (activeTab.value === 'files') {
        filesFilter.value = value;
        return;
      }

      searchText.value = value;
    },
  });

  const outlineDisplayItems = computed(() => {
    const items = headings.value;
    const keyword = searchText.value.trim().toLowerCase();
    if (!keyword) {
      return items;
    }

    const includedLines = new Set<number>();
    items.forEach((item, index) => {
      if (!item.text.toLowerCase().includes(keyword)) {
        return;
      }

      includedLines.add(item.line);
      let parentLevel = item.level;
      for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
        const candidate = items[cursor];
        if (candidate.level < parentLevel) {
          includedLines.add(candidate.line);
          parentLevel = candidate.level;
        }

        if (parentLevel <= 1) {
          break;
        }
      }
    });

    return items.filter((item) => includedLines.has(item.line));
  });

  const currentFileName = computed(() => {
    if (!store.filePath) {
      return 'Untitled.md';
    }

    const normalized = store.filePath.replace(/\\/g, '/');
    const parts = normalized.split('/');
    return parts[parts.length - 1] || store.filePath;
  });

  const windowTitle = computed(() => {
    return `${currentFileName.value} - UniReadMD`;
  });

  return { sidebarSearchPlaceholder,sidebarSearchText,outlineDisplayItems,
    currentFileName,windowTitle,
  };
}
