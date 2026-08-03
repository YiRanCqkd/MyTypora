import { bridgeService } from '../../services/bridge';
import { logBridge, createTraceId } from '../../utils/trace';
import type { createMarkdownRenderer } from '../markdown/renderer';
import type { useSessionPersistence } from './session-persistence';
import type { usePreferences } from './preferences';
import type { createLayoutMetrics } from '../workspace/layout-metrics';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';
type UseContentZoomDependencies =
  ReturnType<typeof createMarkdownRenderer> &
  ReturnType<typeof useSessionPersistence> &
  ReturnType<typeof usePreferences> &
  ReturnType<typeof createLayoutMetrics> &
  ReturnType<typeof createWorkbenchRuntime>;

export function useContentZoom(input: UseContentZoomDependencies) {
  const { CONTENT_ZOOM_DEFAULT,CONTENT_ZOOM_STEP,clearContentZoomPersistTimer,contentZoom,
    normalizeContentZoom,persistSessionSnapshot,runtime,
  } = input;

  function schedulePersistContentZoom() {
    clearContentZoomPersistTimer();
    runtime.contentZoomPersistTimer = window.setTimeout(() => {
      runtime.contentZoomPersistTimer = 0;
      void bridgeService.setSetting('contentZoom', contentZoom.value)
        .catch(async (error) => {
          const message = error instanceof Error ? error.message : String(error);
          await logBridge('error', 'content-zoom-sync:fail', createTraceId('content-zoom'), {
            message,
            contentZoom: contentZoom.value,
          });
        });
    }, 180);
  }

  function setContentZoom(nextZoom: number) {
    const normalizedZoom = normalizeContentZoom(nextZoom);
    if (normalizedZoom === contentZoom.value) {
      return;
    }

    contentZoom.value = normalizedZoom;
    schedulePersistContentZoom();
    persistSessionSnapshot();
  }

  function stepContentZoom(direction: 'in' | 'out') {
    const delta = direction === 'in' ? CONTENT_ZOOM_STEP : -CONTENT_ZOOM_STEP;
    setContentZoom(contentZoom.value + delta);
  }

  function resetContentZoom() {
    setContentZoom(CONTENT_ZOOM_DEFAULT);
  }

  return { schedulePersistContentZoom,setContentZoom,stepContentZoom,resetContentZoom,
  };
}
