import { computed } from 'vue';
import type { useWorkbenchStore } from './editor-store';
type UseDocumentStatisticsDependencies =
  ReturnType<typeof useWorkbenchStore>;

export function useDocumentStatistics(input: UseDocumentStatisticsDependencies) {
  const { store,
  } = input;

  const contentCharCount = computed(() => {
    return store.content.replace(/\r?\n/g, '').length;
  });

  const contentWordCount = computed(() => {
    const tokens = store.content
      .replace(/\r?\n/g, ' ')
      .split(/\s+/)
      .map((item) => item.trim())
      .filter((item) => item.length > 0);
    return tokens.length;
  });

  const contentReadMinutes = computed(() => {
    return Math.max(1, Math.ceil(contentWordCount.value / 220));
  });

  const statusWordLabel = computed(() => {
    return `${contentWordCount.value} Words`;
  });

  return { contentCharCount,contentWordCount,contentReadMinutes,statusWordLabel,
  };
}
