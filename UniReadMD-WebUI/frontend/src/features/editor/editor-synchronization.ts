import { EditorView } from '@codemirror/view';
import type { useEditorHosts } from './editor-hosts';
import type { createWorkbenchRuntime } from './workbench-runtime';
type UseEditorSynchronizationDependencies =
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof createWorkbenchRuntime>;

export function useEditorSynchronization(input: UseEditorSynchronizationDependencies) {
  const { previewHost,runtime,
  } = input;

  function createFontSizeExtension(size: number) {
    return EditorView.theme({
      '&': {
        fontSize: `${size}px`,
      },
    });
  }

  function applyStoreContentToEditor(nextContent: string) {
    if (!runtime.editorView) {
      return;
    }

    const currentContent = runtime.editorView.state.doc.toString();
    if (currentContent === nextContent) {
      return;
    }

    runtime.isSyncingFromStore = true;
    runtime.editorView.dispatch({
      changes: {
        from: 0,
        to: runtime.editorView.state.doc.length,
        insert: nextContent,
      },
    });
    runtime.isSyncingFromStore = false;
  }

  function scrollEditorToLine(lineNumber: number) {
    if (!runtime.editorView) {
      return;
    }

    const totalLines = runtime.editorView.state.doc.lines;
    const safeLine = Math.min(Math.max(lineNumber, 1), totalLines);
    const line = runtime.editorView.state.doc.line(safeLine);

    runtime.editorView.dispatch({
      selection: {
        anchor: line.from,
      },
      effects: EditorView.scrollIntoView(line.from, {
        y: 'center',
      }),
    });

    runtime.editorView.focus();
  }

  function refreshPreviewHeadingNodes() {
    const preview = previewHost.value;
    if (!preview) {
      runtime.previewHeadingNodes = [];
      return;
    }

    const nodes = Array.from(
      preview.querySelectorAll<HTMLElement>(
        "[data-source-token='heading_open'][data-source-start]",
      ),
    );
    runtime.previewHeadingNodes = nodes
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
      } => item !== null)
      .sort((left, right) => left.line - right.line);
  }

  function resolvePreviewHeadingNode(lineNumber: number) {
    if (runtime.previewHeadingNodes.length === 0) {
      refreshPreviewHeadingNodes();
    }

    if (runtime.previewHeadingNodes.length === 0) {
      return null;
    }

    const exactHeading = runtime.previewHeadingNodes.find((item) => item.line === lineNumber);
    if (exactHeading) {
      return exactHeading.node;
    }

    let nearestNext: {
      line: number;
      node: HTMLElement;
    } | null = null;
    let nearestPrev: {
      line: number;
      node: HTMLElement;
    } | null = null;

    for (const item of runtime.previewHeadingNodes) {
      const line = item.line;
      if (line >= lineNumber) {
        if (!nearestNext || line < nearestNext.line) {
          nearestNext = {
            line,
            node: item.node,
          };
        }
        continue;
      }

      if (!nearestPrev || line > nearestPrev.line) {
        nearestPrev = {
          line,
          node: item.node,
        };
      }
    }

    if (nearestNext) {
      return nearestNext.node;
    }

    if (nearestPrev) {
      return nearestPrev.node;
    }

    return null;
  }

  return { createFontSizeExtension,applyStoreContentToEditor,scrollEditorToLine,
    refreshPreviewHeadingNodes,resolvePreviewHeadingNode,
  };
}
