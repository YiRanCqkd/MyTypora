import { ref } from 'vue';

export function useDocumentSearchState() {
  const showDocSearchPanel = ref(false);

  const showDocReplacePanel = ref(false);

  const docSearchInput = ref<HTMLInputElement | null>(null);

  const docSearchQuery = ref('');

  const docReplaceQuery = ref('');

  const docSearchCaseSensitive = ref(false);

  const docSearchWholeWord = ref(false);
  const docSearchRegex = ref(false);

  const docSearchMatchIndex = ref(0);

  return { showDocSearchPanel,showDocReplacePanel,docSearchInput,docSearchQuery,
    docReplaceQuery,docSearchCaseSensitive,docSearchWholeWord,docSearchRegex,docSearchMatchIndex,
  };
}
