import { computed, nextTick, onBeforeUnmount, onMounted, reactive, ref, watch } from 'vue';
import { bridgeService } from '../../services/bridge';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import { countDocumentWords } from '../editor/document-statistics';
import { countPlainWords } from '../markdown/document-text';
import { normalizeWritingPreferences, readWritingPreferences, WRITING_KEY }
  from './writing-preferences';

export function useWritingEnhancements(domains: WorkbenchDomains) {
  const { store, runtime, previewHost } = domains.editor.foundation;
  const preferences = reactive(readWritingPreferences());
  const visible = ref(false);
  const selected = ref('');
  const selectedMarkdown = ref(false);
  const selectionWords = computed(() => selectedMarkdown.value
    ? countDocumentWords(selected.value) : countPlainWords(selected.value));
  const selectionCharacters = computed(() =>
    Array.from(selected.value.replace(/\r?\n/g, '')).length);
  const hasSelection = computed(() => selected.value.length > 0);
  const wordLabel = computed(() => {
    const count = preferences.statUnit === 'characters'
      ? hasSelection.value ? selectionCharacters.value
        : domains.editor.presentation().contentCharCount.value
      : hasSelection.value ? selectionWords.value
        : domains.editor.presentation().contentWordCount.value;
    const unit = preferences.statUnit === 'characters' ? 'chars' : 'Words';
    return `${hasSelection.value ? '选区 ' : ''}${count} ${unit}`;
  });
  const readMinutes = computed(() => Math.max(1, Math.ceil(
    domains.editor.presentation().contentWordCount.value
      / normalizeWritingPreferences(preferences).readingSpeed,
  )));
  function updateSelection() {
    const active = document.activeElement;
    selectedMarkdown.value = true;
    if (active instanceof HTMLTextAreaElement
      && active.closest('.preview-inline-editor')) {
      selected.value = active.value.slice(active.selectionStart, active.selectionEnd);
    } else if (active?.closest('.cm-editor') && runtime.editorView) {
      const { from, to } = runtime.editorView.state.selection.main;
      selected.value = runtime.editorView.state.doc.sliceString(from, to);
    } else {
      const range = window.getSelection();
      selectedMarkdown.value = false;
      selected.value = range && previewHost.value?.contains(range.anchorNode)
        && previewHost.value?.contains(range.focusNode) ? range.toString() : '';
    }
  }
  let selectionFrame = 0;
  function scheduleSelection() {
    if (selectionFrame) return;
    // 等待浏览器完成选区和输入事件，避免统计刷新覆盖尚未提交到 v-model 的文字。
    selectionFrame = requestAnimationFrame(() => { selectionFrame = 0; updateSelection(); });
  }
  function applyTheme() {
    const dark = document.documentElement.dataset.theme === 'dark';
    document.documentElement.dataset.writingTheme = dark
      ? preferences.darkTheme : preferences.lightTheme;
  }
  let observer: MutationObserver | undefined;
  let revision = 0;
  let nativeReady = false;
  let disposed = false;
  let writeQueue = Promise.resolve();
  const serialize = () => JSON.stringify(normalizeWritingPreferences(preferences));
  function persistNative() {
    const json = serialize();
    // 只通过 IPC 传递普通字符串；串行保存，避免快速切换设置时旧值覆盖新值。
    writeQueue = writeQueue.then(() => bridgeService.setSetting('writingPreferencesJson', json))
      .then(() => {}).catch((error: unknown) => {
        if (!disposed) domains.editor.foundation.errorMessage.value = `写作设置保存失败：${error}`;
      });
  }
  async function restoreNative() {
    if (!window.nativeBridge || Reflect.get(window.nativeBridge, '__fallbackBridge')) return;
    const initialRevision = revision;
    try {
      const { settings } = await bridgeService.getAllSettings();
      if (disposed) return;
      // 启动读取期间已经发生的用户修改优先，不让异步旧设置覆盖当前输入。
      if (revision === initialRevision && settings.writingPreferencesJson) {
        Object.assign(preferences, normalizeWritingPreferences(
          JSON.parse(settings.writingPreferencesJson),
        ));
        await nextTick();
      }
      if (disposed) return;
      nativeReady = true;
      persistNative();
    } catch (error) {
      if (!disposed) domains.editor.foundation.errorMessage.value = `写作设置读取失败：${error}`;
    }
  }
  const events = ['selectionchange', 'keyup', 'mouseup', 'input', 'select'];
  onMounted(() => {
    events.forEach((event) => document.addEventListener(event, scheduleSelection, true));
    observer = new MutationObserver(applyTheme);
    observer.observe(document.documentElement, {
      attributes: true, attributeFilter: ['data-theme'],
    });
    applyTheme();
    void restoreNative();
  });
  onBeforeUnmount(() => {
    disposed = true;
    cancelAnimationFrame(selectionFrame);
    events.forEach((event) => document.removeEventListener(event, scheduleSelection, true));
    observer?.disconnect();
  });
  watch(() => store.documentRevision, () => { selected.value = ''; });
  watch(preferences, () => {
    revision += 1;
    try { localStorage.setItem(WRITING_KEY, serialize()); } catch { /* 原生配置仍可保存。 */ }
    if (nativeReady) persistNative();
    applyTheme();
  }, { deep: true, flush: 'sync' });
  watch(() => preferences.equationNumbering, () => {
    domains.editor.presentation().mathPluginRevision.value += 1;
  });
  return reactive({ preferences, visible, wordLabel, readMinutes, hasSelection,
    selectionWords, selectionCharacters, open: () => { visible.value = true; } });
}
