import type { SpellDictionaryStatus } from './preferences-contracts';
import type { Tab } from '../workspace/workspace-contracts';
import { bridgeService } from '../../services/bridge';
import { nextTick } from 'vue';
import type { useSidebarTabs } from '../workspace/sidebar-tabs';
import type { useCursorPosition } from '../editor/cursor-position';
import type { useViewOptions } from './view-options';
import type { usePreviewNavigation } from '../editor/preview-navigation';
import type { useTaskRunner } from '../workspace/task-runner';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';
import type { useCodeMirrorEditor } from '../editor/codemirror-editor';
import type { createPreviewEditingRanges } from '../editor/preview-editing-ranges';
type UseWorkbenchPreferencesDependencies =
  ReturnType<typeof useSidebarTabs> &
  ReturnType<typeof useCursorPosition> &
  ReturnType<typeof useViewOptions> &
  ReturnType<typeof usePreviewNavigation> &
  ReturnType<typeof useTaskRunner> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof useCodeMirrorEditor> &
  ReturnType<typeof createPreviewEditingRanges>;

export function useWorkbenchPreferences(input: UseWorkbenchPreferencesDependencies) {
  const { activeTab,currentCursorLine,focusMode,pinOutline,resolvePreviewViewportAnchorLine,
    runTask,runtime,scheduleTypewriterScroll,showSpellPanel,sourceMode,spellCheckEnabled,
    spellCheckLanguage,spellCheckLanguages,spellDictionaryStatus,syncCursorLineFromState,
    typewriterMode,
  } = input;

  function setTab(tab: Tab) {
    activeTab.value = tab;
  }

  function detectNativeWindowMode() {
    if (typeof navigator === 'undefined') {
      return false;
    }

    return /Electron/i.test(navigator.userAgent);
  }

  function togglePinOutline() {
    pinOutline.value = !pinOutline.value;
  }

  function toggleSourceMode() {
    const nextSourceMode = !sourceMode.value;
    if (nextSourceMode) {
      const previewAnchorLine = resolvePreviewViewportAnchorLine();
      if (previewAnchorLine > 0) {
        currentCursorLine.value = previewAnchorLine;
      }
    } else if (runtime.editorView) {
      syncCursorLineFromState(runtime.editorView.state);
    }

    sourceMode.value = nextSourceMode;
  }

  function toggleFocusMode() {
    focusMode.value = !focusMode.value;
  }

  function toggleTypewriterMode() {
    typewriterMode.value = !typewriterMode.value;

    if (!typewriterMode.value) {
      return;
    }

    nextTick(() => {
      if (!runtime.editorView || !sourceMode.value) {
        return;
      }

      runtime.editorView.focus();
      scheduleTypewriterScroll(runtime.editorView);
    });
  }

  function applySpellState(state: {
    enabled: boolean;
    language: string;
    availableLanguages: string[];
    dictionaryStatus: SpellDictionaryStatus;
  }) {
    spellCheckEnabled.value = state.enabled;
    spellCheckLanguage.value = state.language;
    spellCheckLanguages.value = state.availableLanguages;
    spellDictionaryStatus.value = state.dictionaryStatus;
  }

  async function refreshSpellCheckState() {
    const result = await runTask('spell-get-state', (traceId) => {
      return bridgeService.getSpellCheckState(traceId);
    });

    if (!result) {
      return;
    }

    applySpellState(result);
  }

  function toggleSpellPanel() {
    showSpellPanel.value = !showSpellPanel.value;
  }

  async function changeSpellCheckEnabled() {
    const result = await runTask('spell-set-enabled', (traceId) => {
      return bridgeService.setSpellCheckEnabled(spellCheckEnabled.value, traceId);
    });

    if (!result) {
      return;
    }

    applySpellState(result);
  }

  async function changeSpellCheckLanguage() {
    const result = await runTask('spell-set-language', (traceId) => {
      return bridgeService.setSpellCheckLanguage(spellCheckLanguage.value, traceId);
    });

    if (!result) {
      return;
    }

    applySpellState(result);
  }

  return { setTab,detectNativeWindowMode,togglePinOutline,toggleSourceMode,
    toggleFocusMode,toggleTypewriterMode,applySpellState,refreshSpellCheckState,
    toggleSpellPanel,changeSpellCheckEnabled,changeSpellCheckLanguage,
  };
}
