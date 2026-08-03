import { useEditorStore } from '../../stores/editor';

export function useWorkbenchStore() {
  const store = useEditorStore();

  return { store,
  };
}
