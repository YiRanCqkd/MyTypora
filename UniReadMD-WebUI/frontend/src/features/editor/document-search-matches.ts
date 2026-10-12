import type { DocMatchRange } from './editor-contracts';
import { computed } from 'vue';
import { advanceSearchOffset, buildSearchExpression } from './search-pattern';
import type { useDocumentSearchState } from './document-search-state';
import type { useWorkbenchStore } from './editor-store';
type UseDocumentSearchMatchesDependencies =
  ReturnType<typeof useDocumentSearchState> &
  ReturnType<typeof useWorkbenchStore>;

export function useDocumentSearchMatches(input: UseDocumentSearchMatchesDependencies) {
  const { docSearchCaseSensitive,docSearchMatchIndex,docSearchQuery,docSearchWholeWord,store,
    docSearchRegex,
  } = input;

  function escapeRegExp(input: string) {
    return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function buildSearchRegExp(
    query: string,
    isCaseSensitive: boolean,
    isWholeWord: boolean,
    isGlobal: boolean,
  ) {
    return buildSearchExpression(query, {
      caseSensitive: isCaseSensitive, wholeWord: isWholeWord, global: isGlobal,
      regularExpression: docSearchRegex?.value,
    });
  }

  const docSearchMatches = computed<DocMatchRange[]>(() => {
    const expression = buildSearchRegExp(
      docSearchQuery.value,
      docSearchCaseSensitive.value,
      docSearchWholeWord.value,
      true,
    );

    if (!expression) {
      return [];
    }

    const matches: DocMatchRange[] = [];
    const content = store.content;
    let current: RegExpExecArray | null = null;
    while ((current = expression.exec(content)) !== null) {
      matches.push({
        from: current.index,
        to: current.index + current[0].length,
        text: current[0],
      });

      if (current[0].length === 0) {
        expression.lastIndex = advanceSearchOffset(content, expression.lastIndex);
      }
    }

    return matches;
  });

  const docSearchSummary = computed(() => {
    const total = docSearchMatches.value.length;
    if (total === 0) {
      if (docSearchQuery.value && !buildSearchRegExp(docSearchQuery.value,
        docSearchCaseSensitive.value, docSearchWholeWord.value, false)) {
        return '正则表达式无效';
      }
      return 'No Results';
    }

    return `${docSearchMatchIndex.value + 1}/${total}`;
  });

  return { escapeRegExp,buildSearchRegExp,docSearchMatches,docSearchSummary,
  };
}
