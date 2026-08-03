import {
  useViewportLifecycle,
  type ViewportLifecycleFeatures,
} from './viewport-lifecycle';
import {
  useWorkbenchLifecycle,
  type WorkbenchLifecycleFeatures,
} from './workbench-lifecycle';
import {
  useWorkbenchWatchers,
  type WorkbenchWatcherFeatures,
} from './workbench-watchers';

export type WorkbenchRuntimeFeatures = Readonly<{
  lifecycle: WorkbenchLifecycleFeatures;
  viewport: ViewportLifecycleFeatures;
  watchers: WorkbenchWatcherFeatures;
}>;

export function useWorkbenchRuntimeFeature(input: WorkbenchRuntimeFeatures) {
  const viewport = useViewportLifecycle(input.viewport);
  useWorkbenchWatchers(input.watchers);
  useWorkbenchLifecycle(input.lifecycle, viewport);
}
