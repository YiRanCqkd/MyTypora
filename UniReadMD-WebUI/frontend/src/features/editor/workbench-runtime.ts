import type { WorkbenchRuntime } from './editor-contracts';

export function createWorkbenchRuntime() {
  const runtime: WorkbenchRuntime = {
    latestFrontMatter: '',
    editorView: null,
    isSyncingFromStore: false,
    mermaidRenderSeed: 0,
    mermaidThemeKey: '',
    mermaidLoadPromise: null,
    mermaidApi: null,
    flowchartLoadPromise: null,
    flowchartApi: null,
    mathPluginLoadPromise: null,
    sidebarResizeCleanup: null,
    mainPreviewResizeCleanup: null,
    menuCommandCleanup: null,
    launchFileCleanup: null,
    typewriterScrollFrame: 0,
    previewScrollSyncFrame: 0,
    previewInlineEditSyncTimer: 0,
    contentZoomPersistTimer: 0,
    previewHeadingNodes: [],
    previewSourceNodes: [],
    fileDragDepth: 0,
  };

  return { runtime };
}
