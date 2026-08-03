import type { Theme } from './preferences-contracts';
import { bridgeService } from '../../services/bridge';
import { nextTick } from 'vue';
import type { useAppearance } from './appearance';
import type { useExportCapabilities } from '../export/export-capabilities';
import type { useDocumentExport } from '../export/export-document';
import type { useSourceEditing } from '../editor/source-editing';
import type { useDocumentSession } from '../workspace/document-session';
import type { useDiagramBlocks } from '../markdown/diagram-blocks';
import type { useTaskRunner } from '../workspace/task-runner';
import type { usePreferences } from './preferences';
import type { useWorkbenchPreferences } from './workbench-preferences';
type UseNativeCommandsDependencies =
  ReturnType<typeof useAppearance> &
  ReturnType<typeof useExportCapabilities> &
  ReturnType<typeof useDocumentExport> &
  ReturnType<typeof useSourceEditing> &
  ReturnType<typeof useDocumentSession> &
  ReturnType<typeof useDiagramBlocks> &
  ReturnType<typeof useTaskRunner> &
  ReturnType<typeof usePreferences> &
  ReturnType<typeof useWorkbenchPreferences>;

export function useNativeCommands(input: UseNativeCommandsDependencies) {
  const { applyTheme,exportDocx,exportHtml,exportJpeg,exportPdf,exportPng,openDocSearchPanel,
    openFile,renderMermaidBlocks,runTask,saveAsFile,saveFile,theme,toggleFocusMode,
    toggleSpellPanel,toggleTypewriterMode,
  } = input;

  async function changeTheme() {
    applyTheme(theme.value);

    await runTask('settings-theme', (traceId) => {
      return bridgeService.setSetting('theme', theme.value, traceId);
    });

    await nextTick();
    await renderMermaidBlocks();
  }

  async function minimizeWindow() {
    await runTask('window-minimize', (traceId) => bridgeService.minimizeWindow(traceId));
  }

  async function toggleMaximizeWindow() {
    await runTask('window-toggle-maximize', (traceId) => {
      return bridgeService.toggleMaximizeWindow(traceId);
    });
  }

  async function closeWindow() {
    await runTask('window-close', (traceId) => bridgeService.closeWindow(traceId));
  }

  async function handleMenuCommand(command: string) {
    if (command === 'file-open') {
      await openFile();
      return;
    }

    if (command === 'file-save') {
      await saveFile();
      return;
    }

    if (command === 'file-save-as') {
      await saveAsFile();
      return;
    }

    if (command === 'file-export-html') {
      await exportHtml();
      return;
    }

    if (command === 'file-export-pdf') {
      await exportPdf();
      return;
    }

    if (command === 'file-export-png') {
      await exportPng();
      return;
    }

    if (command === 'file-export-jpeg') {
      await exportJpeg();
      return;
    }

    if (command === 'file-export-docx') {
      await exportDocx();
      return;
    }

    if (command === 'edit-search') {
      openDocSearchPanel(false);
      return;
    }

    if (command === 'edit-focus') {
      toggleFocusMode();
      return;
    }

    if (command === 'edit-typewriter') {
      toggleTypewriterMode();
      return;
    }

    if (command === 'help-spell') {
      toggleSpellPanel();
      return;
    }

    if (command === 'theme-light' || command === 'theme-dark' || command === 'theme-system') {
      const nextTheme = command.replace('theme-', '') as Theme;
      if (theme.value === nextTheme) {
        return;
      }

      theme.value = nextTheme;
      await changeTheme();
    }
  }

  return { changeTheme,minimizeWindow,toggleMaximizeWindow,closeWindow,
    handleMenuCommand,
  };
}
