import type { useCursorPosition } from './cursor-position';
import type { useEditorHosts } from './editor-hosts';
import type { useEditorSynchronization } from './editor-synchronization';
import type { createWorkbenchRuntime } from './workbench-runtime';
type UsePreviewNavigationDependencies =
  ReturnType<typeof useCursorPosition> &
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof useEditorSynchronization> &
  ReturnType<typeof createWorkbenchRuntime>;

export function usePreviewNavigation(input: UsePreviewNavigationDependencies) {
  const { currentCursorLine,previewHost,resolvePreviewHeadingNode,runtime,
  } = input;

  function refreshPreviewSourceNodes() {
    const preview = previewHost.value;
    if (!preview) {
      runtime.previewSourceNodes = [];
      return;
    }

    const nodes = Array.from(
      preview.querySelectorAll<HTMLElement>('[data-source-start]'),
    );
    runtime.previewSourceNodes = nodes
      .map((node) => {
        const line = Number(node.dataset.sourceStart || 0);
        if (!Number.isFinite(line) || line <= 0) {
          return null;
        }

        return {
          line,
          node,
        };
      })
      .filter((item): item is {
        line: number;
        node: HTMLElement;
      } => item !== null);
  }

  function resolvePreviewNodeByLine(lineNumber: number) {
    const headingNode = resolvePreviewHeadingNode(lineNumber);
    if (headingNode) {
      return headingNode;
    }

    if (runtime.previewSourceNodes.length === 0) {
      refreshPreviewSourceNodes();
    }

    if (runtime.previewSourceNodes.length === 0) {
      return null;
    }

    const exact = runtime.previewSourceNodes.find((item) => item.line === lineNumber);
    if (exact) {
      return exact.node;
    }

    let nearestNext: {
      line: number;
      node: HTMLElement;
    } | null = null;
    let nearestPrev: {
      line: number;
      node: HTMLElement;
    } | null = null;

    for (const item of runtime.previewSourceNodes) {
      if (item.line >= lineNumber) {
        if (!nearestNext || item.line < nearestNext.line) {
          nearestNext = item;
        }
        continue;
      }

      if (!nearestPrev || item.line > nearestPrev.line) {
        nearestPrev = item;
      }
    }

    if (nearestNext) {
      return nearestNext.node;
    }

    if (nearestPrev) {
      return nearestPrev.node;
    }

    return runtime.previewSourceNodes[0].node;
  }

  function resolvePreviewViewportAnchorLine() {
    const preview = previewHost.value;
    if (!preview) {
      return 0;
    }

    if (runtime.previewSourceNodes.length === 0) {
      refreshPreviewSourceNodes();
    }

    if (runtime.previewSourceNodes.length === 0) {
      return 0;
    }

    const anchorTop = preview.scrollTop + 18;
    let activeLine = runtime.previewSourceNodes[0].line;
    for (const item of runtime.previewSourceNodes) {
      if (item.node.offsetTop <= anchorTop) {
        activeLine = item.line;
        continue;
      }

      break;
    }

    return activeLine;
  }

  function scrollPreviewToLine(lineNumber: number) {
    const preview = previewHost.value;
    if (!preview) {
      return;
    }

    const target = resolvePreviewNodeByLine(lineNumber);
    if (!target) {
      return;
    }

    const paddingTop = 16;
    const targetTop = Math.max(0, target.offsetTop - paddingTop);
    preview.scrollTo({
      top: targetTop,
      behavior: 'auto',
    });
  }

  function syncCursorLineFromPreviewViewport() {
    const activeLine = resolvePreviewViewportAnchorLine();

    if (activeLine > 0 && currentCursorLine.value !== activeLine) {
      currentCursorLine.value = activeLine;
    }
  }

  return { refreshPreviewSourceNodes,resolvePreviewNodeByLine,
    resolvePreviewViewportAnchorLine,scrollPreviewToLine,syncCursorLineFromPreviewViewport,
  };
}
