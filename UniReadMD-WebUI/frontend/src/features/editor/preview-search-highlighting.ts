import type { useDocumentSearchMatches } from './document-search-matches';
import { advanceSearchOffset } from './search-pattern';
import type { createPreviewSearchDom } from './preview-search-dom';
import type { useDocumentSearchState } from './document-search-state';
import type { useEditorHosts } from './editor-hosts';
import type { useViewOptions } from '../settings/view-options';
type UsePreviewSearchHighlightingDependencies =
  ReturnType<typeof useDocumentSearchMatches> &
  ReturnType<typeof createPreviewSearchDom> &
  ReturnType<typeof useDocumentSearchState> &
  ReturnType<typeof useEditorHosts> & Pick<ReturnType<typeof useViewOptions>, 'sourceMode'>;

export function usePreviewSearchHighlighting(input: UsePreviewSearchHighlightingDependencies) {
  const { buildSearchRegExp,clearPreviewSearchHighlights,docSearchCaseSensitive,
    docSearchMatchIndex,docSearchQuery,docSearchWholeWord,normalizeSearchIndex,previewHost,
    shouldSkipPreviewSearchHighlight,showDocSearchPanel,sourceMode,
  } = input;

  function applyPreviewSearchHighlights() {
    const preview = previewHost.value;
    clearPreviewSearchHighlights();
    if (!preview || !showDocSearchPanel.value || sourceMode.value) {
      return;
    }

    const expression = buildSearchRegExp(
      docSearchQuery.value,
      docSearchCaseSensitive.value,
      docSearchWholeWord.value,
      true,
    );
    if (!expression) {
      return;
    }

    const textNodes: Text[] = [];
    const walker = document.createTreeWalker(preview, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        return shouldSkipPreviewSearchHighlight(node as Text)
          ? NodeFilter.FILTER_REJECT
          : NodeFilter.FILTER_ACCEPT;
      },
    });

    let currentNode = walker.nextNode();
    while (currentNode) {
      textNodes.push(currentNode as Text);
      currentNode = walker.nextNode();
    }

    const marks: HTMLElement[] = [];
    textNodes.forEach((textNode) => {
      const value = textNode.nodeValue || '';
      const localExpression = new RegExp(expression.source, expression.flags);
      let match = localExpression.exec(value);
      if (!match) {
        return;
      }

      const fragment = document.createDocumentFragment();
      let lastIndex = 0;
      while (match) {
        const [matchedText] = match;
        const matchIndex = match.index;
        if (matchIndex > lastIndex) {
          fragment.appendChild(document.createTextNode(value.slice(lastIndex, matchIndex)));
        }

        const mark = document.createElement('mark');
        mark.className = 'doc-search-highlight';
        mark.dataset.docSearchHighlight = 'true';
        mark.textContent = matchedText;
        fragment.appendChild(mark);
        marks.push(mark);
        lastIndex = matchIndex + matchedText.length;

        if (matchedText.length === 0) {
          localExpression.lastIndex = advanceSearchOffset(value, localExpression.lastIndex);
        }
        match = localExpression.exec(value);
      }

      if (lastIndex < value.length) {
        fragment.appendChild(document.createTextNode(value.slice(lastIndex)));
      }

      textNode.parentNode?.replaceChild(fragment, textNode);
    });

    if (marks.length === 0) {
      return;
    }

    const activeIndex = normalizeSearchIndex(docSearchMatchIndex.value, marks.length);
    const activeMark = marks[activeIndex];
    if (activeMark) {
      activeMark.classList.add('is-active');
    }
  }

  return { applyPreviewSearchHighlights,
  };
}
