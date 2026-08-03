import { ref } from 'vue';

export function useEditorHosts() {
  const editorPreviewHost = ref<HTMLElement | null>(null);

  const editorHost = ref<HTMLDivElement | null>(null);

  const previewHost = ref<HTMLElement | null>(null);

  return { editorPreviewHost,editorHost,previewHost,
  };
}
