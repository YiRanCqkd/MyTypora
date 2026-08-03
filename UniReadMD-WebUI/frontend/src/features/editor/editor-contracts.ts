import type { EditorView } from '@codemirror/view';

export interface DocMatchRange {
  from: number;
  to: number;
  text: string;
}

export interface PreviewInlineEditSelection {
  startLine: number;
  endLine: number;
  token: string;
}

export interface PreviewInlineEditSession {
  from: number;
  to: number;
  startLine: number;
  originalText: string;
  appliedText: string;
}

export interface PreviewTableSelection {
  startLine: number;
  endLine: number;
  rowIndex: number;
  colIndex: number;
}

export interface PreviewTableCellContext extends PreviewTableSelection {
  table: HTMLElement;
  cell: HTMLElement;
}

export interface PreviewTaskItemContext {
  startLine: number;
  endLine: number;
  taskIndex: number;
  globalTaskIndex: number;
}

export interface WorkbenchRuntime {
  latestFrontMatter: string;
  editorView: EditorView | null;
  isSyncingFromStore: boolean;
  mermaidRenderSeed: number;
  mermaidThemeKey: string;
  mermaidLoadPromise: Promise<unknown> | null;
  mermaidApi: {
    initialize: (config: Record<string, unknown>) => void;
    render: (id: string, source: string) => Promise<{ svg: string }>;
  } | null;
  flowchartLoadPromise: Promise<unknown> | null;
  flowchartApi: {
    parse: (code: string) => {
      drawSVG: (
        container: HTMLElement | string,
        options?: Record<string, unknown>,
      ) => void;
    };
  } | null;
  mathPluginLoadPromise: Promise<void> | null;
  sidebarResizeCleanup: (() => void) | null;
  mainPreviewResizeCleanup: (() => void) | null;
  menuCommandCleanup: (() => void) | null;
  launchFileCleanup: (() => void) | null;
  typewriterScrollFrame: number;
  previewScrollSyncFrame: number;
  previewInlineEditSyncTimer: number;
  contentZoomPersistTimer: number;
  previewHeadingNodes: Array<{ line: number; node: HTMLElement }>;
  previewSourceNodes: Array<{ line: number; node: HTMLElement }>;
  fileDragDepth: number;
}
