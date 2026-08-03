import { computed } from 'vue';
import type { usePreferences } from './preferences';
import type { useViewOptions } from './view-options';
import type { useWorkspaceState } from '../workspace/workspace-state';
import type { useDocumentSearchState } from '../editor/document-search-state';
type UseLayoutModelDependencies =
  ReturnType<typeof usePreferences> &
  ReturnType<typeof useViewOptions> &
  ReturnType<typeof useWorkspaceState> &
  ReturnType<typeof useDocumentSearchState>;

export function useLayoutModel(input: UseLayoutModelDependencies) {
  const { contentZoom,focusMode,isNativeWindow,isResizingMainPreview,isResizingSidebar,
    mainPreviewWidth,pinOutline,showDocReplacePanel,showDocSearchPanel,sidebarWidth,sourceMode,
    typewriterMode,
  } = input;

  const layoutClassName = computed(() => {
    return {
      'native-window': isNativeWindow.value,
      'pin-outline': pinOutline.value,
      'hide-sidebar': !pinOutline.value,
      'is-resizing-sidebar': isResizingSidebar.value,
      'is-resizing-main-preview': isResizingMainPreview.value,
      'typora-sourceview-on': sourceMode.value,
      'on-focus-mode': focusMode.value,
      'ty-on-typewriter-mode': typewriterMode.value,
      'on-search-panel-open': showDocSearchPanel.value,
      'on-replace-panel-open': showDocSearchPanel.value && showDocReplacePanel.value,
    };
  });

  const layoutStyle = computed(() => {
    return {
      '--sidebar-width': `${sidebarWidth.value}px`,
      '--main-preview-width': mainPreviewWidth.value > 0
        ? `${mainPreviewWidth.value}px`
        : '58%',
      '--content-zoom': `${contentZoom.value / 100}`,
    };
  });

  return { layoutClassName,layoutStyle,
  };
}
