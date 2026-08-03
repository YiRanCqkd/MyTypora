import { computed } from 'vue';
import type { useCursorPosition } from './cursor-position';
import type { useWorkbenchStore } from './editor-store';
type UseStatusLineDependencies =
  ReturnType<typeof useCursorPosition> &
  ReturnType<typeof useWorkbenchStore>;

export function useStatusLine(input: UseStatusLineDependencies) {
  const { currentCursorLine,store,
  } = input;

  const statusLineLabel = computed(() => {
    if (store.cursorHint > 0) {
      return `L${store.cursorHint}`;
    }

    return `L${currentCursorLine.value}`;
  });

  return { statusLineLabel,
  };
}
