import { nextTick, onMounted, onBeforeUnmount } from 'vue';
import { EditorView } from '@codemirror/view';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import { createParityState } from './parity-state';
import { bridgeService } from '../../services/bridge';
import { headingSlug } from '../markdown/code-highlight';
import { createParityImages } from './parity-images';
import { createParityPreview } from './parity-preview';
import { usePreviewImageSizing } from './preview-image-sizing';
import { usePreviewInsertMenu } from './preview-insert-menu';
import { useVisualPreviewEditing } from './visual-preview-editing';
import { createParityFileOperations } from './parity-files';
import { createClipboardWorkflows } from '../editor/clipboard-workflows';
import { setExportResources } from '../export/export-options';

export function useParityWorkflows(domains: WorkbenchDomains) {
  const editor = domains.editor.foundation;
  const { store, runtime, errorMessage } = editor;
  const workspace = domains.workspace.foundation();
  const settings = domains.settings.foundation();
  const actions = domains.workspace.actions();
  const { dialog, quickQuery, drafts, entryName, entryAction, selectedRow,
    fileMenu, fileMenuPosition, notice, externalChanged, busy, selectedExport,
    pdfOptions, exportOptions, quickFiles, pandocAvailable, canExportSelected,
  } = createParityState(domains);

  async function execute(name: string, action: () => Promise<unknown>) {
    if (busy.value) return;
    busy.value = true;
    try { return await domains.workspace.presentation().runTask(name, action); }
    finally { busy.value = false; }
  }

  function goToLine(lineNumber: number) {
    const view = runtime.editorView;
    if (!view) return;
    const line = Math.max(1, Math.min(lineNumber, view.state.doc.lines));
    view.dispatch({ selection: { anchor: view.state.doc.line(line).from },
      effects: EditorView.scrollIntoView(view.state.doc.line(line).from, { y: 'center' }) });
    editor.currentCursorLine.value = line;
    store.setCursorHint(line);
    domains.editor.navigation().scrollPreviewToLine(line);
    if (settings.sourceMode.value) view.focus();
  }

  async function openFileAt(filePath: string, line = 1) {
    const result = await actions.runDocumentLoadingTask('打开并定位文档', () => {
      return bridgeService.openFromPath(filePath);
    });
    if (!result) return;
    store.setFile(result.filePath, result.content, result);
    domains.settings.session().pushRecentFile(result.filePath);
    dialog.value = '';
    externalChanged.value = false;
    await nextTick();
    goToLine(line);
  }

  function jumpToAnchor(fragment: string) {
    let anchor = fragment.replace(/^#/, '');
    try { anchor = decodeURIComponent(anchor); } catch { /* 保留无法解码的字面锚点。 */ }
    const host = editor.previewHost.value;
    const node = [...(host?.querySelectorAll<HTMLElement>('[id]') || [])]
      .find((item) => item.id === anchor || item.id === headingSlug(anchor));
    if (!node) { errorMessage.value = `未找到标题：${anchor}`; return; }
    node.scrollIntoView({ block: 'center' });
    const line = Number(node.dataset.sourceStart || 1);
    if (line) goToLine(line);
  }

  async function followLink(href: string) {
    if (href.startsWith('#')) { jumpToAnchor(href); return; }
    const openLinked = window.nativeBridge.file.openLinkedFile;
    if (!openLinked || !store.filePath) {
      errorMessage.value = '请先保存文档，然后打开本地 Markdown 链接。';
      return;
    }
    const result = await actions.runDocumentLoadingTask('打开文档链接', () => {
      return openLinked({ sourceFilePath: store.filePath, href });
    });
    if (!result) return;
    store.setFile(result.filePath, result.content, result);
    domains.settings.session().pushRecentFile(result.filePath);
    await nextTick();
    const hash = href.indexOf('#');
    if (hash >= 0) jumpToAnchor(href.slice(hash));
  }

  async function showDrafts() {
    await runtime.flushSnapshot?.();
    await execute('读取历史草稿', async () => {
      const result = await window.nativeBridge.session?.listDrafts?.();
      if (!result) throw new Error('历史草稿仅在桌面版提供');
      drafts.value = result.drafts;
      dialog.value = 'drafts';
    });
  }

  async function collectImages() {
    runtime.flushEdits?.();
    const sourceFilePath = store.filePath;
    const content = store.content;
    if (!sourceFilePath) { errorMessage.value = '请先保存文档，再整理图片。'; return; }
    await execute('整理本地图片', async () => {
      const collect = window.nativeBridge.images?.collect;
      if (!collect) throw new Error('图片整理仅在桌面版提供');
      const result = await collect({ sourceFilePath, content });
      if (!result) return;
      if (store.filePath !== sourceFilePath || store.content !== content) {
        throw new Error('整理期间文档已变化，图片副本已保留，请重试以更新引用。');
      }
      const view = runtime.editorView;
      if (result.content !== content && view) view.dispatch({
        changes: { from: 0, to: view.state.doc.length, insert: result.content },
      });
      notice.value = `已整理 ${result.count} 张本地图片，原图保留，请保存文档。`;
    });
  }

  async function restoreDraft(id: string) {
    const result = await actions.runDocumentLoadingTask('恢复历史草稿', async () => {
      return await window.nativeBridge.session?.restoreDraft?.(id) || null;
    });
    if (!result) return;
    store.setFile('', result.content);
    store.isDirty = true;
    dialog.value = '';
    notice.value = '已恢复为未命名副本，请另存为。';
  }

  function insertText(text: string) {
    const presentation = domains.editor.presentation();
    const textarea = presentation.previewInlineEditTextarea.value;
    if (presentation.previewInlineEditVisible.value && textarea) {
      const from = textarea.selectionStart;
      const to = textarea.selectionEnd;
      presentation.previewInlineEditValue.value = presentation.previewInlineEditValue.value
        .slice(0, from) + text + presentation.previewInlineEditValue.value.slice(to);
      domains.editor.preview().flushPreviewInlineEditSync();
      domains.editor.events().updatePreviewInlineEditorSelection(from + text.length,
        from + text.length);
    } else {
      const view = runtime.editorView;
      if (!view) return;
      const { from, to } = view.state.selection.main;
      view.dispatch({ changes: { from, to, insert: text },
        selection: { anchor: from + text.length } });
    }
  }

  async function importDocument() {
    const result = await actions.runDocumentLoadingTask('导入文档', async () => {
      return await window.nativeBridge.exporter.importDocument?.() || null;
    });
    if (result) {
      store.setFile(result.filePath, result.content, result);
      domains.settings.session().pushRecentFile(result.filePath);
    }
  }

  async function changeInterval() {
    const interval = Number(settings.autoSaveInterval.value);
    if (!Number.isFinite(interval) || interval < 2 || interval > 3600) {
      settings.autoSaveInterval.value = 300;
      errorMessage.value = '保存间隔必须在 2 至 3600 秒之间，已恢复默认值。';
      return;
    }
    settings.autoSaveInterval.value = Math.round(interval);
    const result = await domains.workspace.presentation().runTask('设置自动保存间隔', () => {
      return bridgeService.setSetting('autoSaveInterval', settings.autoSaveInterval.value);
    });
    if (result?.settings.autoSaveInterval) {
      settings.autoSaveInterval.value = result.settings.autoSaveInterval;
    }
  }

  function savePdfOptions() {
    localStorage.setItem('unireadmd.pdf-options', JSON.stringify(pdfOptions.value));
    dialog.value = '';
    notice.value = 'PDF 导出设置已保存。';
  }

  function saveExportOptions() {
    setExportResources(exportOptions.value);
    localStorage.setItem('unireadmd.export-options', JSON.stringify(exportOptions.value));
    dialog.value = '';
    notice.value = '导出设置已保存。';
  }

  function applyExportPreset(event: Event) {
    const preset = (event.target as HTMLSelectElement).value;
    Object.assign(exportOptions.value, {
      tableOfContents: preset !== 'default', tocDepth: 3, epubChapterLevel: 1,
      headingPageBreak: preset === 'article', includeStyles: true,
    });
  }

  async function chooseExportResource(kind: 'referenceDoc' | 'epubCover') {
    await execute('选择导出资源', async () => {
      const choose = window.nativeBridge.exporter.chooseResource;
      if (!choose) throw new Error('选择参考文档和封面仅在桌面版提供');
      const result = await choose(kind);
      if (result) exportOptions.value[kind] = result.filePath;
    });
  }

  function setIndependentMargins(event: Event) {
    const checked = (event.target as HTMLInputElement).checked;
    const margin = pdfOptions.value.marginMm;
    pdfOptions.value.marginsMm = checked
      ? { top: margin, bottom: margin, left: margin, right: margin } : undefined;
  }

  function clickLink(event: MouseEvent) {
    const target = event.target instanceof Element
      ? event.target.closest<HTMLAnchorElement>('a') : null;
    if (!target || !editor.previewHost.value?.contains(target)
      || !(event.ctrlKey || event.metaKey)) return;
    const href = target.getAttribute('href') || '';
    if (/^(https?:|mailto:|tel:|\/\/)/i.test(href)) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    void followLink(href);
  }

  function globalKeys(event: KeyboardEvent) {
    if (event.key === 'Escape' && dialog.value) {
      event.preventDefault();
      event.stopImmediatePropagation();
      dialog.value = '';
    }
    if (!(event.ctrlKey || event.metaKey)) return;
    if (event.key.toLowerCase() === 'p' && !event.shiftKey) {
      event.preventDefault();
      event.stopImmediatePropagation();
      quickQuery.value = '';
      dialog.value = 'quick';
    }
    if (event.shiftKey && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      event.stopImmediatePropagation();
      workspace.activeTab.value = 'files';
      void nextTick(() => {
        document.querySelector<HTMLInputElement>('.workspace-search input')?.focus();
      });
    }
  }

  const { insertLocalImages, setImageRoot, pasteImage, dropImage } = createParityImages({
    domains, execute, insertText, dialog, busy, notice,
  });
  const { openFileMenu, requestEntry, applyFileOperation, trashEntry, moveEntry,
    startFileDrag, dropOnDirectory } = createParityFileOperations({
    domains, execute, selectedRow, fileMenu, fileMenuPosition, entryName, entryAction, dialog,
  });
  const { previewKeyboard, syncPreviewCursor } = createParityPreview(domains);
  const imageSizing = usePreviewImageSizing(domains);
  const visualEditing = useVisualPreviewEditing(domains);
  const previewInsertMenu = usePreviewInsertMenu(domains, insertLocalImages);
  const { pasteRichText, copyMarkdown, clipboardKeys } = createClipboardWorkflows(
    domains, insertText,
  );
  let pollTimer = 0;
  let inspecting = false;
  function closeFileMenu(event: MouseEvent) {
    if (!(event.target instanceof Element) || !event.target.closest('.workspace-context-menu')) {
      fileMenu.value = false;
    }
  }
  function recovered() { notice.value = '已恢复上次未保存的会话，可在历史草稿中选择其他版本。'; }
  onMounted(() => {
    document.addEventListener('click', clickLink, true);
    window.addEventListener('keydown', globalKeys, true);
    document.addEventListener('keydown', previewKeyboard, true);
    document.addEventListener('paste', pasteRichText, true);
    document.addEventListener('paste', pasteImage, true);
    document.addEventListener('copy', copyMarkdown, true);
    window.addEventListener('keydown', clipboardKeys, true);
    document.addEventListener('drop', dropImage, true);
    document.addEventListener('mousedown', closeFileMenu);
    document.addEventListener('input', syncPreviewCursor);
    document.addEventListener('keyup', syncPreviewCursor);
    document.addEventListener('click', syncPreviewCursor);
    window.addEventListener('unireadmd-recovery', recovered);
    pollTimer = window.setInterval(async () => {
      if (!store.filePath || !window.nativeBridge.file.inspect || inspecting
        || document.visibilityState === 'hidden') return;
      inspecting = true;
      const filePath = store.filePath;
      try {
        const result = await window.nativeBridge.file.inspect({ filePath });
        if (store.filePath === filePath) externalChanged.value = result.version !== store.version;
      } catch { if (store.filePath === filePath) externalChanged.value = true; }
      finally { inspecting = false; }
    }, 5000);
  });
  onBeforeUnmount(() => {
    window.clearInterval(pollTimer);
    document.removeEventListener('click', clickLink, true);
    window.removeEventListener('keydown', globalKeys, true);
    document.removeEventListener('keydown', previewKeyboard, true);
    document.removeEventListener('paste', pasteImage, true);
    document.removeEventListener('paste', pasteRichText, true);
    document.removeEventListener('copy', copyMarkdown, true);
    window.removeEventListener('keydown', clipboardKeys, true);
    document.removeEventListener('drop', dropImage, true);
    document.removeEventListener('mousedown', closeFileMenu);
    document.removeEventListener('input', syncPreviewCursor);
    document.removeEventListener('keyup', syncPreviewCursor);
    document.removeEventListener('click', syncPreviewCursor);
    window.removeEventListener('unireadmd-recovery', recovered);
  });

  return {
    documentToolsVisible: settings.documentToolsVisible,
    imageSizing,
    visualEditing,
    previewInsertMenu,
    collectImages,
    applyExportPreset,
    dialog, quickQuery, quickFiles, drafts, entryName, entryAction, selectedRow, errorMessage,
    fileMenu, fileMenuPosition, notice, externalChanged, busy, selectedExport,
    pdfOptions, exportOptions, pandocAvailable, canExportSelected,
    autoSaveInterval: settings.autoSaveInterval,
    openFileAt, showDrafts, restoreDraft, insertLocalImages, setImageRoot, openFileMenu,
    requestEntry, applyFileOperation, trashEntry, moveEntry, startFileDrag, dropOnDirectory,
    showNativeFileMenu: async () => {
      const row = selectedRow.value;
      fileMenu.value = false;
      if (!row) return;
      const showMenu = window.nativeBridge.context?.showFileTreeMenu;
      if (showMenu) await execute('文件路径菜单', () => showMenu({ filePath: row.path }));
    },
    importDocument, changeInterval, savePdfOptions, saveExportOptions, chooseExportResource,
    setIndependentMargins, goToLine, insertText,
    exportSelected: () => domains.export.feature().exportDocumentByFormat(selectedExport.value),
    reloadCurrent: () => openFileAt(store.filePath),
  };
}

export type ParityWorkflows = ReturnType<typeof useParityWorkflows>;
