import type { useEditorHosts } from './editor-hosts';
type CreatePreviewSearchDomDependencies =
  ReturnType<typeof useEditorHosts>;

export function createPreviewSearchDom(input: CreatePreviewSearchDomDependencies) {
  const { previewHost,
  } = input;

  function clearPreviewSearchHighlights() {
    const preview = previewHost.value;
    if (!preview) {
      return;
    }

    const marks = Array.from(
      preview.querySelectorAll<HTMLElement>("mark[data-doc-search-highlight='true']"),
    );
    marks.forEach((mark) => {
      const parent = mark.parentNode;
      if (!parent) {
        return;
      }

      while (mark.firstChild) {
        parent.insertBefore(mark.firstChild, mark);
      }
      parent.removeChild(mark);
    });
    preview.normalize();
  }

  function shouldSkipPreviewSearchHighlight(node: Text) {
    const parent = node.parentElement;
    const value = node.nodeValue || '';
    if (!parent || !value.trim()) {
      return true;
    }

    if (parent.closest("mark[data-doc-search-highlight='true']")) {
      return true;
    }

    const tagName = parent.tagName;
    if (['SCRIPT', 'STYLE', 'TEXTAREA', 'INPUT', 'OPTION'].includes(tagName)) {
      return true;
    }

    return Boolean(
      parent.closest(
        '.mermaid-block, .preview-inline-editor, .preview-table-cell-editor, .preview-table-menu',
      ),
    );
  }

  function normalizeSearchIndex(index: number, length: number) {
    if (length <= 0) {
      return 0;
    }

    if (index < 0) {
      return length - 1;
    }

    if (index >= length) {
      return 0;
    }

    return index;
  }

  return { clearPreviewSearchHighlights,shouldSkipPreviewSearchHighlight,
    normalizeSearchIndex,
  };
}
