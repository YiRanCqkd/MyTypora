import { ref } from 'vue';

export function useDocumentLoadingState() {
  const isDocumentLoading = ref(false);

  const isDragOverFile = ref(false);

  return { isDocumentLoading,isDragOverFile,
  };
}
