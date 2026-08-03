import { computed } from 'vue';
import type { useViewOptions } from './view-options';
type UseStatusLabelsDependencies =
  ReturnType<typeof useViewOptions>;

export function useStatusLabels(input: UseStatusLabelsDependencies) {
  const { sourceMode,spellDictionaryStatus,
  } = input;

  const sourceToggleLabel = computed(() => {
    if (sourceMode.value) {
      return '退出源码模式';
    }

    return '源码模式';
  });

  const spellDictionaryLabel = computed(() => {
    return spellDictionaryStatus.value === 'ready' ? 'Dictionary Ready' : 'Dictionary Missing';
  });

  return { sourceToggleLabel,spellDictionaryLabel,
  };
}
