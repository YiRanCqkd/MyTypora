import { computed } from 'vue';
import MarkdownIt from 'markdown-it';
import { getCachedDocumentWords, rememberDocumentWords } from '../markdown/document-text';
import type { useWorkbenchStore } from './editor-store';
type UseDocumentStatisticsDependencies =
  ReturnType<typeof useWorkbenchStore>;

const statisticsParser = new MarkdownIt({ html: true });

export function countDocumentWords(content: string) {
  const cached = getCachedDocumentWords(content);
  if (cached !== null) return cached;
  const tokens = statisticsParser.parse(content.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, ''), {});
  return rememberDocumentWords(content, tokens);
}

export function useDocumentStatistics(input: UseDocumentStatisticsDependencies) {
  const { store,
  } = input;

  const contentCharCount = computed(() => {
    return Array.from(store.content.replace(/\r?\n/g, '')).length;
  });

  const contentWordCount = computed(() => {
    return countDocumentWords(store.content);
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
