export type Tab = 'files' | 'outline';

export type FileViewMode = 'list' | 'tree';

export type FileSortMode = 'natural' | 'name' | 'date';

export interface SessionSnapshot {
  workspaceDirectory?: string;
  version?: string;
  filePath: string;
  content: string;
  isDirty: boolean;
  cursorHint: number;
  recentFiles: string[];
  contentZoom?: number;
  pinOutline?: boolean;
  sourceMode?: boolean;
  focusMode?: boolean;
  typewriterMode?: boolean;
  sidebarWidth?: number;
  mainPreviewWidth?: number;
  fileViewMode?: FileViewMode;
  fileSortMode?: FileSortMode;
  filesFilter?: string;
  searchPanelVisible?: boolean;
  replacePanelVisible?: boolean;
  searchQuery?: string;
  replaceQuery?: string;
  searchCaseSensitive?: boolean;
  searchWholeWord?: boolean;
  searchRegex?: boolean;
  searchMatchIndex?: number;
  fileTreeExpanded?: Record<string, boolean>;
}

export interface FileListItem {
  path: string;
  name: string;
  parent: string;
  ext: string;
  order: number;
}

export interface FileTreeRow {
  id: string;
  type: 'dir' | 'file';
  depth: number;
  label: string;
  path: string;
  expandable: boolean;
  expanded: boolean;
  filePath?: string;
}
