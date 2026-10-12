import { bridgeService } from '../../services/bridge';
import { nextTick, onBeforeUnmount, onMounted } from 'vue';
import { measuredContextMenuStyle } from '../workspace/context-menu-position';
import type { useTableActions } from './table-actions';
import type { usePreviewEditingSession } from './preview-editing-session';
import type { useTableCellEditing } from './table-cell-editing';
import type { createTableMutations } from './table-mutations';
import type { usePreviewInteractions } from './preview-interactions';
import type { useEditorHosts } from './editor-hosts';
import type { useTableNavigation } from './table-navigation';
import type { useTaskRunner } from '../workspace/task-runner';
import type { createWorkbenchRuntime } from './workbench-runtime';
import type { usePreviewEditingState } from './preview-editing-state';
import type { useViewOptions } from '../settings/view-options';
type UseTableEditingEventsDependencies =
  ReturnType<typeof useTableActions> &
  ReturnType<typeof usePreviewEditingSession> &
  ReturnType<typeof useTableCellEditing> &
  ReturnType<typeof createTableMutations> &
  ReturnType<typeof usePreviewInteractions> &
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof useTableNavigation> &
  ReturnType<typeof useTaskRunner> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof useViewOptions>;

export function useTableEditingEvents(input: UseTableEditingEventsDependencies) {
  const { applyPreviewTableActionFromCellEditor,closePreviewInlineEditor,
    closePreviewTableCellEditor,closePreviewTableCellEditorAndCommit,closePreviewTableMenu,
    closeSourceContextMenu,editorPreviewHost,movePreviewTableCellEditor,runTask,runtime,
    sourceContextMenuStyle,sourceContextMenuVisible,sourceMode,
  } = input;

  function handlePreviewTableCellEditKeydown(event: KeyboardEvent) {
    if (event.ctrlKey && event.shiftKey && event.altKey && event.key === 'Enter') {
      event.preventDefault();
      void applyPreviewTableActionFromCellEditor('insert-col-right');
      return;
    }

    if (
      event.ctrlKey
      && event.shiftKey
      && event.altKey
      && (event.key === 'Delete' || event.key === 'Backspace')
    ) {
      event.preventDefault();
      void applyPreviewTableActionFromCellEditor('delete-col');
      return;
    }

    if (event.ctrlKey && event.shiftKey && event.key === 'Enter') {
      event.preventDefault();
      void applyPreviewTableActionFromCellEditor('insert-row-below');
      return;
    }

    if (event.ctrlKey && event.shiftKey && (event.key === 'Delete' || event.key === 'Backspace')) {
      event.preventDefault();
      void applyPreviewTableActionFromCellEditor('delete-row');
      return;
    }

    if (event.ctrlKey && event.altKey && event.key === 'ArrowUp') {
      event.preventDefault();
      void applyPreviewTableActionFromCellEditor('move-row-up');
      return;
    }

    if (event.ctrlKey && event.altKey && event.key === 'ArrowDown') {
      event.preventDefault();
      void applyPreviewTableActionFromCellEditor('move-row-down');
      return;
    }

    if (event.ctrlKey && event.altKey && event.key === 'ArrowLeft') {
      event.preventDefault();
      void applyPreviewTableActionFromCellEditor('move-col-left');
      return;
    }

    if (event.ctrlKey && event.altKey && event.key === 'ArrowRight') {
      event.preventDefault();
      void applyPreviewTableActionFromCellEditor('move-col-right');
      return;
    }

    if (event.key === 'Escape') {
      event.preventDefault();
      closePreviewTableCellEditor();
      return;
    }

    if (event.key === 'Tab') {
      event.preventDefault();
      void movePreviewTableCellEditor(event.shiftKey ? 'tab-prev' : 'tab-next');
      return;
    }

    if (event.key === 'Enter' && !event.shiftKey && !event.ctrlKey && !event.metaKey) {
      event.preventDefault();
      void movePreviewTableCellEditor('enter-down');
    }
  }

  async function handleSourceContextMenu(event: MouseEvent) {
    if (!sourceMode.value) {
      return;
    }

    event.preventDefault();
    closePreviewTableMenu();
    closePreviewTableCellEditorAndCommit();
    closePreviewInlineEditor();
    runtime.editorView?.focus();

    const x = Number.isFinite(event.clientX) ? Math.round(event.clientX) : undefined;
    const y = Number.isFinite(event.clientY) ? Math.round(event.clientY) : undefined;
    const nativeResult = await runTask('source-editor-context-menu', (traceId) => {
      return bridgeService.showSourceEditorContextMenu({
        x,
        y,
        traceId,
      });
    });

    if (nativeResult?.handled) {
      closeSourceContextMenu();
      return;
    }

    const container = editorPreviewHost.value;
    if (!container) {
      closeSourceContextMenu();
      return;
    }

    sourceContextMenuStyle.value = {
      left: `${event.clientX}px`, top: `${event.clientY}px`, visibility: 'hidden',
    };
    sourceContextMenuVisible.value = true;
    await nextTick();
    const menu = document.querySelector<HTMLElement>('.source-context-menu');
    if (menu && sourceContextMenuVisible.value) {
      sourceContextMenuStyle.value = measuredContextMenuStyle(menu, event);
      await nextTick();
      if (sourceContextMenuVisible.value) {
        menu.querySelector('button')?.focus({ preventScroll: true });
      }
    }
  }

  function dismissMenus(event: Event) {
    if (event.target instanceof Element
      && event.target.closest('.preview-table-menu, .source-context-menu')) return;
    closePreviewTableMenu();
    closeSourceContextMenu();
  }
  onMounted(() => {
    document.addEventListener('scroll', dismissMenus, true);
    document.addEventListener('wheel', dismissMenus, true);
    window.addEventListener('blur', dismissMenus);
    window.addEventListener('resize', dismissMenus);
  });
  onBeforeUnmount(() => {
    document.removeEventListener('scroll', dismissMenus, true);
    document.removeEventListener('wheel', dismissMenus, true);
    window.removeEventListener('blur', dismissMenus);
    window.removeEventListener('resize', dismissMenus);
  });

  return { handlePreviewTableCellEditKeydown,handleSourceContextMenu,
  };
}
