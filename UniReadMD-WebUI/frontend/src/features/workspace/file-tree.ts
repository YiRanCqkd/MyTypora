import type { FileTreeRow } from './workspace-contracts';
import { computed } from 'vue';
import type { createFilePathModel } from './file-paths';
import type { useWorkspaceState } from './workspace-state';
type UseFileTreeDependencies =
  ReturnType<typeof createFilePathModel> &
  ReturnType<typeof useWorkspaceState>;

export function useFileTree(input: UseFileTreeDependencies) {
  const { currentDirectoryLabel,currentDirectoryMarkdownFiles,dedupeMarkdownPaths,fileListItems,
    fileTreeExpanded,filesFilter,getFileName,currentDirectoryPath,fileSortMode,fileViewMode,
    fileModifiedTimes,
    workspaceDirectories,
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
    currentFiles.sort((left, right) => {
      if (fileSortMode.value === 'date') {
        const difference = (fileModifiedTimes.value[right] || 0)
          - (fileModifiedTimes.value[left] || 0);
        if (difference) return difference;
      }
      return getFileName(left).localeCompare(getFileName(right),
        'zh-CN', { numeric: fileSortMode.value === 'natural' });
    });
    const recentFilesInSidebar = filterByName(fileListItems.value.map((item) => item.path));

    const pushGroupRows = (groupId: string, groupLabel: string, paths: string[]) => {
      if (paths.length === 0 && !(groupId === currentId && currentDirectoryPath.value)) {
        return;
      }

      const expanded = fileTreeExpanded.value[groupId] !== false;
      output.push({
        id: groupId,
        type: 'dir',
        depth: 0,
        label: groupLabel,
        path: groupId === currentId ? currentDirectoryPath.value : groupId,
        expandable: true,
        expanded,
      });

      if (!expanded) {
        return;
      }

      const children = new Map<string, Array<FileTreeRow>>();
      const directories = new Set<string>();
      const add = (parent: string, row: FileTreeRow) => {
        const rows = children.get(parent) || [];
        rows.push(row);
        children.set(parent, rows);
      };
      const treePaths = groupId === currentId && !fileKeyword && fileViewMode.value === 'tree'
        ? [...paths, ...workspaceDirectories.value.map((directory) => `${directory}/`)] : paths;
      treePaths.forEach((filePath) => {
        let parentId = groupId;
        let depth = 1;
        if (groupId === currentId && fileViewMode.value === 'tree') {
          const relative = filePath.replace(/\\/g, '/')
            .slice(currentDirectoryPath.value.replace(/\\/g, '/').length + 1);
          const parts = relative.split('/').slice(0, -1);
          let directoryPath = currentDirectoryPath.value;
          for (const label of parts) {
            directoryPath += `/${label}`;
            const id = `dir:${directoryPath}`;
            if (!directories.has(id)) {
              directories.add(id);
              add(parentId, {
                id, type: 'dir', depth, label, path: directoryPath,
                expandable: true, expanded: fileTreeExpanded.value[id] !== false,
              });
            }
            parentId = id;
            depth += 1;
          }
        }
        if (filePath.endsWith('/')) return;
        add(parentId, {
          id: `file:${filePath}`,
          type: 'file',
          depth,
          label: getFileName(filePath),
          path: filePath,
          expandable: false,
          expanded: false,
          filePath,
        });
      });
      const visit = (parent: string) => {
        const rows = children.get(parent) || [];
        rows.sort((left, right) => {
          if (left.type !== right.type) return left.type === 'dir' ? -1 : 1;
          if (left.type === 'dir') return left.label.localeCompare(right.label,
            'zh-CN', { numeric: true });
          return 0;
        });
        for (const row of rows) {
          output.push(row);
          if (row.type === 'dir' && row.expanded) visit(row.id);
        }
      };
      visit(groupId);
    };

    pushGroupRows(currentId, currentDirectoryLabel.value, currentFiles);
    pushGroupRows(recentId, '最近打开', recentFilesInSidebar);
    return output;
  });

  return { fileTreeRows,
  };
}
