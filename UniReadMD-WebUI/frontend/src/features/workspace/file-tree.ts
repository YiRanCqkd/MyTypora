import type { FileTreeRow } from './workspace-contracts';
import { computed } from 'vue';
import type { createFilePathModel } from './file-paths';
import type { useWorkspaceState } from './workspace-state';
type UseFileTreeDependencies =
  ReturnType<typeof createFilePathModel> &
  ReturnType<typeof useWorkspaceState>;

export function useFileTree(input: UseFileTreeDependencies) {
  const { currentDirectoryLabel,currentDirectoryMarkdownFiles,dedupeMarkdownPaths,fileListItems,
    fileTreeExpanded,filesFilter,getFileName,
  } = input;

  const fileTreeRows = computed<FileTreeRow[]>(() => {
    const output: FileTreeRow[] = [];
    const currentId = 'dir:group-current';
    const recentId = 'dir:group-recent';
    const fileKeyword = filesFilter.value.trim().toLowerCase();
    const filterByName = (paths: string[]) => {
      if (!fileKeyword) {
        return paths;
      }

      return paths.filter((item) => getFileName(item).toLowerCase().includes(fileKeyword));
    };
    const currentFiles = filterByName(dedupeMarkdownPaths(currentDirectoryMarkdownFiles.value));
    const recentFilesInSidebar = filterByName(fileListItems.value.map((item) => item.path));

    const pushGroupRows = (groupId: string, groupLabel: string, paths: string[]) => {
      if (paths.length === 0) {
        return;
      }

      const expanded = fileTreeExpanded.value[groupId] !== false;
      output.push({
        id: groupId,
        type: 'dir',
        depth: 0,
        label: groupLabel,
        path: groupId,
        expandable: true,
        expanded,
      });

      if (!expanded) {
        return;
      }

      paths.forEach((filePath) => {
        output.push({
          id: `file:${filePath}`,
          type: 'file',
          depth: 1,
          label: getFileName(filePath),
          path: filePath,
          expandable: false,
          expanded: false,
          filePath,
        });
      });
    };

    pushGroupRows(currentId, currentDirectoryLabel.value, currentFiles);
    pushGroupRows(recentId, '最近打开', recentFilesInSidebar);
    return output;
  });

  return { fileTreeRows,
  };
}
