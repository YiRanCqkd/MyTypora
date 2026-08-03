import { ref } from 'vue';

export function useCursorPosition() {
  const currentCursorLine = ref(1);

  return { currentCursorLine,
  };
}
