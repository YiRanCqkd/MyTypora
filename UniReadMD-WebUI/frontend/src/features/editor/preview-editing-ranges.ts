import { EditorState } from '@codemirror/state';
import type { useCursorPosition } from './cursor-position';
import type { useEditorHosts } from './editor-hosts';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createWorkbenchRuntime } from './workbench-runtime';
import type { usePreviewNavigation } from './preview-navigation';
type CreatePreviewEditingRangesDependencies =
  ReturnType<typeof useCursorPosition> &
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof usePreviewNavigation>;

export function createPreviewEditingRanges(input: CreatePreviewEditingRangesDependencies) {
  const { currentCursorLine,editorPreviewHost,previewHost,previewInlineEditSelection,runtime,
    syncCursorLineFromPreviewViewport,
  } = input;

  function schedulePreviewScrollSync() {
    if (runtime.previewScrollSyncFrame) {
      return;
    }

    runtime.previewScrollSyncFrame = requestAnimationFrame(() => {
      runtime.previewScrollSyncFrame = 0;
      syncCursorLineFromPreviewViewport();
    });
  }

  function syncCursorLineFromState(state: EditorState) {
    currentCursorLine.value = state.doc.lineAt(state.selection.main.head).number;
  }

  function normalizeLineRange(startLine: number, endLine: number) {
    if (!runtime.editorView) {
      return null;
    }

    const totalLines = runtime.editorView.state.doc.lines;
    if (totalLines <= 0) {
      return null;
    }

    const safeStart = Math.min(Math.max(startLine, 1), totalLines);
    const safeEnd = Math.min(Math.max(endLine, safeStart), totalLines);
    return {
      startLine: safeStart,
      endLine: safeEnd,
    };
  }

  function getEditorRangeByLines(startLine: number, endLine: number) {
    if (!runtime.editorView) {
      return null;
    }

    const lineRange = normalizeLineRange(startLine, endLine);
    if (!lineRange) {
      return null;
    }

    const doc = runtime.editorView.state.doc;
    const from = doc.line(lineRange.startLine).from;
    const to = doc.line(lineRange.endLine).to;
    return {
      ...lineRange,
      from,
      to,
    };
  }

  function normalizePreviewInlineEditContent(content: string) {
    return content.replace(/\r\n/g, '\n');
  }

  function getPreviewInlineEditLineCount(content: string) {
    const normalized = normalizePreviewInlineEditContent(content);
    if (normalized.length === 0) {
      return 1;
    }

    return normalized.split('\n').length;
  }

  function buildPreviewInlineEditStyle(node: HTMLElement) {
    const container = editorPreviewHost.value;
    const preview = previewHost.value;
    if (!container || !preview) {
      return null;
    }

    const containerRect = container.getBoundingClientRect();
    const previewRect = preview.getBoundingClientRect();
    const targetRect = node.getBoundingClientRect();
    const left = Math.max(targetRect.left, previewRect.left) - containerRect.left;
    const top = targetRect.top - containerRect.top;
    const availableWidth = Math.max(120, previewRect.right - containerRect.left - left - 2);
    const targetWidth = Math.max(120, Math.round(targetRect.width) + 4);
    const width = Math.min(targetWidth, availableWidth);
    const minHeight = Math.max(24, Math.round(targetRect.height));

    return {
      left: `${left}px`,
      top: `${top}px`,
      width: `${width}px`,
      minHeight: `${minHeight}px`,
    };
  }

  function findPreviewInlineEditAnchorNode() {
    const preview = previewHost.value;
    const selection = previewInlineEditSelection.value;
    if (!preview || !selection) {
      return null;
    }

    const startLineSelector = `[data-source-start='${selection.startLine}']`;
    const tokenSelector = `${startLineSelector}[data-source-token='${selection.token}']`;
    const tokenMatched = preview.querySelector<HTMLElement>(tokenSelector);
    if (tokenMatched) {
      return tokenMatched;
    }

    const startLineMatched = preview.querySelector<HTMLElement>(startLineSelector);
    if (startLineMatched) {
      return startLineMatched;
    }

    const mappedNodes = Array.from(
      preview.querySelectorAll<HTMLElement>(
        '[data-source-start][data-source-end][data-source-token]',
      ),
    );
    return mappedNodes.find((node) => {
      const start = Number(node.dataset.sourceStart || 0);
      const end = Number(node.dataset.sourceEnd || 0);
      return Number.isFinite(start)
        && Number.isFinite(end)
        && start <= selection.startLine
        && end >= selection.startLine;
    }) || null;
  }


  return {
    schedulePreviewScrollSync,
    syncCursorLineFromState,
    normalizeLineRange,
    getEditorRangeByLines,
    normalizePreviewInlineEditContent,
    getPreviewInlineEditLineCount,
    buildPreviewInlineEditStyle,
    findPreviewInlineEditAnchorNode,
  };
}
