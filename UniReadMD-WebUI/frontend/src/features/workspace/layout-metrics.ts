import type { createMarkdownRenderer } from '../markdown/renderer';
import type { usePreferences } from '../settings/preferences';
import type { createFilePathModel } from './file-paths';
import type { useEditorHosts } from '../editor/editor-hosts';
import type { useWorkspaceState } from './workspace-state';
import type { useViewOptions } from '../settings/view-options';
type CreateLayoutMetricsDependencies =
  ReturnType<typeof createMarkdownRenderer> &
  ReturnType<typeof usePreferences> &
  ReturnType<typeof createFilePathModel> &
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof useWorkspaceState> &
  ReturnType<typeof useViewOptions>;

export function createLayoutMetrics(input: CreateLayoutMetricsDependencies) {
  const { CONTENT_ZOOM_DEFAULT,CONTENT_ZOOM_MAX,CONTENT_ZOOM_MIN,CONTENT_ZOOM_STEP,
    MAIN_PREVIEW_WIDTH_DEFAULT_RATIO,MAIN_PREVIEW_WIDTH_MIN,MAIN_SOURCE_WIDTH_MIN,
    RECENT_FILES_LIMIT,SIDEBAR_WIDTH_DEFAULT,SIDEBAR_WIDTH_MAX,SIDEBAR_WIDTH_MIN,contentZoom,
    dedupeMarkdownPaths,editorPreviewHost,fontSize,mainPreviewWidth,sourceMode,
  } = input;

  function normalizeSidebarWidth(width: number) {
    if (!Number.isFinite(width)) {
      return SIDEBAR_WIDTH_DEFAULT;
    }

    const rounded = Math.round(width);
    return Math.min(SIDEBAR_WIDTH_MAX, Math.max(SIDEBAR_WIDTH_MIN, rounded));
  }

  function normalizeContentZoom(value: number) {
    if (!Number.isFinite(value)) {
      return CONTENT_ZOOM_DEFAULT;
    }

    const rounded = Math.round(value / CONTENT_ZOOM_STEP) * CONTENT_ZOOM_STEP;
    return Math.min(CONTENT_ZOOM_MAX, Math.max(CONTENT_ZOOM_MIN, rounded));
  }

  function normalizeMainPreviewWidth(width: number, containerWidth: number) {
    if (!Number.isFinite(width) || !Number.isFinite(containerWidth) || containerWidth <= 0) {
      return 0;
    }

    const minWidth = Math.min(MAIN_PREVIEW_WIDTH_MIN, Math.floor(containerWidth / 2));
    const maxWidth = Math.max(
      minWidth,
      Math.floor(containerWidth - MAIN_SOURCE_WIDTH_MIN),
    );
    const rounded = Math.round(width);
    return Math.min(maxWidth, Math.max(minWidth, rounded));
  }

  function resolveMainPreviewContainerWidth() {
    return editorPreviewHost.value?.clientWidth ?? 0;
  }

  function ensureMainPreviewWidth() {
    if (!sourceMode.value) {
      return;
    }

    const containerWidth = resolveMainPreviewContainerWidth();
    if (containerWidth <= 0) {
      return;
    }

    if (mainPreviewWidth.value <= 0) {
      const initialWidth = containerWidth * MAIN_PREVIEW_WIDTH_DEFAULT_RATIO;
      mainPreviewWidth.value = normalizeMainPreviewWidth(initialWidth, containerWidth);
      return;
    }

    mainPreviewWidth.value = normalizeMainPreviewWidth(
      mainPreviewWidth.value,
      containerWidth,
    );
  }

  function normalizeRecentFiles(paths: string[]) {
    return dedupeMarkdownPaths(paths).slice(0, RECENT_FILES_LIMIT);
  }

  function getEffectiveEditorFontSize() {
    const scaledSize = (fontSize.value * contentZoom.value) / 100;
    return Number(scaledSize.toFixed(2));
  }

  return { normalizeSidebarWidth,normalizeContentZoom,normalizeMainPreviewWidth,
    resolveMainPreviewContainerWidth,ensureMainPreviewWidth,normalizeRecentFiles,
    getEffectiveEditorFontSize,
  };
}
