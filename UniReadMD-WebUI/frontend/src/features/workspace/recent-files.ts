import { ref } from 'vue';

export function useRecentFiles() {
  const recentFiles = ref<string[]>([]);

  return { recentFiles,
  };
}
