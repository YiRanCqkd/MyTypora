import type { useEditorHosts } from './editor-hosts';
import type { usePreviewEditingState } from './preview-editing-state';
import type { useContentZoom } from '../settings/content-zoom';
type UsePreviewWheelDependencies =
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof useContentZoom>;

export function usePreviewWheel(input: UsePreviewWheelDependencies) {
  const { editorPreviewHost,previewHost,previewInlineEditVisible,previewTableCellEditVisible,
    stepContentZoom,
  } = input;

  function handleGlobalWheel(event: WheelEvent) {
    if (event.defaultPrevented || event.deltaY === 0) return;
    const target = event.target instanceof HTMLElement ? event.target : null;
    const isZoomGesture = event.ctrlKey || event.metaKey;
    if (isZoomGesture && target && editorPreviewHost.value?.contains(target)) {
      event.preventDefault();
      stepContentZoom(event.deltaY < 0 ? 'in' : 'out');
      return;
    }

    if (!previewInlineEditVisible.value && !previewTableCellEditVisible.value) {
      return;
    }

    if (
      !target?.closest('.preview-inline-editor')
      && !target?.closest('.preview-table-cell-editor')
    ) {
      return;
    }

    const preview = previewHost.value;
    if (!preview) {
      return;
    }

    event.preventDefault();
    preview.scrollTop += event.deltaY;
  }

  return { handleGlobalWheel,
  };
}
