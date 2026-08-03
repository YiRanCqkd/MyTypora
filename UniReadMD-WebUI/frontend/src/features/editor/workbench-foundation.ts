import { useApplicationStatus } from './application-status';
import { useCursorPosition } from './cursor-position';
import { useDocumentSearchState } from './document-search-state';
import { useEditorHosts } from './editor-hosts';
import { useWorkbenchStore } from './editor-store';
import { createWorkbenchRuntime } from './workbench-runtime';

export function useWorkbenchFoundation() {
  const runtime = createWorkbenchRuntime();
  const store = useWorkbenchStore();
  const application = useApplicationStatus();
  const cursor = useCursorPosition();
  const hosts = useEditorHosts();
  const search = useDocumentSearchState();

  return {
    ...runtime,
    ...store,
    ...application,
    ...cursor,
    ...hosts,
    ...search,
  };
}
