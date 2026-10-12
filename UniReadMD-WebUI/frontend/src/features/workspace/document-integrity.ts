import { ref } from 'vue';
import type { Ref } from 'vue';
import { bridgeService } from '../../services/bridge';
import type { useWorkbenchStore } from '../editor/editor-store';
import type { WorkbenchRuntime } from '../editor/editor-contracts';

export function useDocumentIntegrity(input: ReturnType<typeof useWorkbenchStore> & {
  runtime: WorkbenchRuntime;
  errorMessage: Ref<string>;
}) {
  const { store, runtime, errorMessage } = input;
  const saving = ref(false);
  let confirmation: Promise<boolean> | null = null;

  async function saveCurrentDocument(saveAs = false) {
    runtime.flushEdits?.();
    if (saving.value) return false;
    saving.value = true;
    const revision = store.documentRevision;
    const content = store.content;
    try {
      const result = await bridgeService.saveMarkdown({
        filePath: saveAs ? undefined : store.filePath || undefined,
        content,
        version: saveAs ? undefined : store.version || undefined,
        ...(saveAs && /!\[/.test(content) ? { sourceFilePath: store.filePath || undefined } : {}),
      });
      if (!result) return false;
      if (revision !== store.documentRevision) return false;
      const savedContent = result.content ?? content;
      if (store.content === content && savedContent !== content) store.setContent(savedContent);
      store.markSaved(result.filePath, savedContent, result);
      return !store.isDirty;
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : '保存失败，请重试或另存为';
      return false;
    } finally {
      saving.value = false;
    }
  }

  async function confirmReplacement() {
    runtime.flushEdits?.();
    if (!store.isDirty) return true;
    await runtime.flushSnapshot?.();
    const native = window.nativeBridge.session;
    try {
      const action = native
        ? (await native.confirmDiscard()).action
        : window.confirm('当前文档有未保存修改。是否放弃修改并继续？') ? 'discard' : 'cancel';
      if (action === 'save') return saveCurrentDocument();
      return action === 'discard';
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : '无法确认未保存修改';
      return false;
    }
  }

  function ensureCanReplaceDocument() {
    if (!confirmation) {
      confirmation = confirmReplacement().finally(() => { confirmation = null; });
    }
    return confirmation;
  }

  return { saveCurrentDocument, ensureCanReplaceDocument, saving };
}
