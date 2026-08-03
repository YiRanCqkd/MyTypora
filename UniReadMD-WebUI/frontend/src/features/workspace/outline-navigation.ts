import { bridgeService } from '../../services/bridge';
import type { useCursorPosition } from '../editor/cursor-position';
import type { useViewOptions } from '../settings/view-options';
import type { useFileTreeActions } from './file-tree-actions';
import type { useTaskRunner } from './task-runner';
import type { useEditorSynchronization } from '../editor/editor-synchronization';
import type { usePreviewNavigation } from '../editor/preview-navigation';
import type { useWorkbenchStore } from '../editor/editor-store';
type UseOutlineNavigationDependencies =
  ReturnType<typeof useCursorPosition> &
  ReturnType<typeof useViewOptions> &
  ReturnType<typeof useFileTreeActions> &
  ReturnType<typeof useTaskRunner> &
  ReturnType<typeof useEditorSynchronization> &
  ReturnType<typeof usePreviewNavigation> &
  ReturnType<typeof useWorkbenchStore>;

export function useOutlineNavigation(input: UseOutlineNavigationDependencies) {
  const { currentCursorLine,isNativeWindow,openRecentFile,runTask,scrollEditorToLine,
    scrollPreviewToLine,store,
  } = input;

  function jumpToHeading(line: number) {
    currentCursorLine.value = line;
    store.setCursorHint(line);
    scrollEditorToLine(line);
    scrollPreviewToLine(line);
  }

  async function tryOpenLaunchFileFromNative() {
    if (!isNativeWindow.value) {
      return;
    }

    const result = await runTask('consume-launch-file', (traceId) => {
      return bridgeService.consumeLaunchFilePath(traceId);
    });
    const filePath = result?.filePath || '';
    if (!filePath) {
      return;
    }

    await openRecentFile(filePath);
  }

  async function handleNativeLaunchFile(filePath: string) {
    if (!filePath) {
      return;
    }

    await openRecentFile(filePath);
  }

  return { jumpToHeading,tryOpenLaunchFileFromNative,handleNativeLaunchFile,
  };
}
