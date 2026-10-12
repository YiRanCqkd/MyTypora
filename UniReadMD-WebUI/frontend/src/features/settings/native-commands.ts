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
import type { useViewOptions } from './view-options';
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
  ReturnType<typeof useViewOptions> &
  ReturnType<typeof useWorkbenchPreferences>;

export function useNativeCommands(input: UseNativeCommandsDependencies) {
  let closing = false;
  const { applyTheme,exportDocx,exportHtml,exportJpeg,exportPdf,exportPng,openDocSearchPanel,
    openFile,renderMermaidBlocks,runTask,saveAsFile,saveFile,theme,toggleFocusMode,
    toggleSpellPanel,toggleTypewriterMode,ensureCanReplaceDocument,newFile,openFolder,
  } = input;

  async function changeTheme() {
    applyTheme(theme.value);

    await runTask('settings-theme', (traceId) => {
      return bridgeService.setSetting('theme', theme.value, traceId);
    });

    await nextTick();
    await renderMermaidBlocks();
  }

  async function changeAutoSave() {
    await runTask('设置自动保存', () => bridgeService.setSetting('autoSave', input.autoSave.value));
  }

  async function selectUserCss() {
    await runTask('设置文档样式', async () => {
      const result = await bridgeService.selectUserCssFile();
      if (!result) return;
      await bridgeService.setSetting('userCssPath', result.filePath);
      await bridgeService.setSetting('userCssEnabled', true);
      const css = await window.nativeBridge.settings.getUserCss?.();
      input.applyUserCss(css?.css || '');
      input.userCssEnabled.value = true;
      input.userCssPath.value = result.filePath;
    });
  }

  async function disableUserCss() {
    await runTask('恢复默认文档样式', async () => {
      await bridgeService.setSetting('userCssEnabled', false);
      input.userCssEnabled.value = false;
      input.applyUserCss('');
    });
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
    if (closing) return;
    closing = true;
    try {
      if (!await ensureCanReplaceDocument()) return;
      await runTask('window-close', async (traceId) => {
        if (input.store.isDirty) {
          input.store.setFile(input.store.filePath, input.store.savedContent);
        }
        if (window.nativeBridge.session) {
          await input.runtime.flushSnapshot?.();
          await window.nativeBridge.session.save({
            filePath: input.store.filePath, content: input.store.content,
            isDirty: false, cursorHint: input.store.cursorHint, version: input.store.version,
          });
        }
        return bridgeService.closeWindow(traceId);
      });
    } finally {
      closing = false;
    }
  }

  async function handleMenuCommand(command: string) {
    for (const [name, key] of [['document-tools', 'documentToolsVisible'],
      ['preferences', 'preferencesVisible']] as const) {
      if (command === `view-toggle-${name}`) {
        const visible = !input[key].value;
        const saved = await runTask('保存视图显示设置', () => bridgeService.setSetting(key, visible));
        if (saved) input[key].value = visible;
        return;
      }
      if (command === `view-${name}-show` || command === `view-${name}-hide`) {
        input[key].value = command.endsWith('-show');
        return;
      }
    }
    if (command === 'file-export-odt' || command === 'file-export-rtf') {
      return input.exportDocumentByFormat(command.endsWith('odt') ? 'odt' : 'rtf');
    }
    if (command === 'window-close') return closeWindow();
    if (command === 'file-new') return newFile();
    if (command === 'file-open-folder') return openFolder();
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

  return { changeTheme,changeAutoSave,selectUserCss,disableUserCss,
    minimizeWindow,toggleMaximizeWindow,closeWindow,
    handleMenuCommand,
  };
}
