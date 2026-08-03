import { computed } from 'vue';
import type { usePreferences } from './preferences';
type UseZoomLabelDependencies =
  ReturnType<typeof usePreferences>;

export function useZoomLabel(input: UseZoomLabelDependencies) {
  const { contentZoom,
  } = input;

  const contentZoomLabel = computed(() => {
    return `${contentZoom.value}%`;
  });

  return { contentZoomLabel,
  };
}
