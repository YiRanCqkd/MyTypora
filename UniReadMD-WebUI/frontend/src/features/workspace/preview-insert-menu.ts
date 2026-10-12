import { nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import { Transaction } from '@codemirror/state';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import { buildPreviewInsertion, type PreviewInsertKind } from '../editor/preview-insertion';
import { measuredContextMenuStyle } from './context-menu-position';

export function usePreviewInsertMenu(domains: WorkbenchDomains, insertImages: () => Promise<void>) {
  const { store, runtime, previewHost } = domains.editor.foundation;
  const sourceMode = domains.settings.foundation().sourceMode;
  const visible = ref(false);
  const position = ref<Record<string, string>>({ left: '0px', top: '0px' });
  let point = { x: 0, y: 0 };
  const section = ref<'main' | 'headings' | 'table'>('main');
  const rows = ref(2);
  const columns = ref(2);
  let selection: { source: string; path: string; offset: number } | null = null;
  function close() { visible.value = false; selection = null; section.value = 'main'; }
  function contextMenu(event: MouseEvent) {
    const host = previewHost.value;
    const target = event.target;
    if (!host || sourceMode.value || !(target instanceof Element) || !host.contains(target)
      || target.closest('img, table, pre, a, .mermaid-block')) { close(); return; }
    event.preventDefault();
    event.stopImmediatePropagation();
    runtime.flushVisualEdits?.();
    domains.editor.preview().closePreviewInlineEditor();
    domains.editor.table().closePreviewTableCellEditorAndCommit();
    domains.editor.preview().closePreviewTableMenu();
    domains.editor.preview().closeSourceContextMenu();
    const view = runtime.editorView;
    if (!view) return;
    const source = view.state.doc.toString();
    const blocks = [...host.querySelectorAll<HTMLElement>('[data-source-start][data-source-end]')]
      .filter((node) => node.getBoundingClientRect().top <= event.clientY);
    const previous = blocks.at(-1);
    const line = Math.min(view.state.doc.lines, Number(previous?.dataset.sourceEnd) || 1);
    const lastBottom = host.lastElementChild?.getBoundingClientRect().bottom || 0;
    const firstLine = Number(host.querySelector<HTMLElement>('[data-source-start]')
      ?.dataset.sourceStart) || 1;
    const offset = event.clientY >= lastBottom ? view.state.doc.length
      : previous ? view.state.doc.line(line).to : view.state.doc.line(firstLine).from;
    selection = { source, path: store.filePath, offset };
    section.value = 'main';
    rows.value = 2;
    columns.value = 2;
    point = { x: event.clientX, y: event.clientY };
    position.value = { left: `${point.x}px`, top: `${point.y}px`, visibility: 'hidden' };
    visible.value = true;
    void reposition().then(() => document.querySelector<HTMLElement>('.preview-insert-menu button')
      ?.focus({ preventScroll: true }));
  }
  async function reposition() {
    await nextTick();
    const menu = document.querySelector<HTMLElement>('.preview-insert-menu');
    if (menu && visible.value) position.value = measuredContextMenuStyle(menu, point);
  }
  function active() {
    const view = runtime.editorView;
    return view && selection && selection.path === store.filePath
      && selection.source === view.state.doc.toString() ? { view, ...selection } : null;
  }
  async function insert(kind: PreviewInsertKind) {
    const current = active();
    if (!current) { close(); return; }
    const patch = buildPreviewInsertion(current.source, current.offset, kind, rows.value,
      columns.value);
    close();
    current.view.dispatch({ changes: { from: patch.from, insert: patch.insert },
      selection: { anchor: patch.contentFrom, head: patch.contentTo },
      annotations: Transaction.userEvent.of('input.insert-block') });
    const insertedSource = current.view.state.doc.toString();
    const revision = store.documentRevision;
    await nextTick();
    const line = current.view.state.doc.lineAt(patch.contentFrom).number;
    let node = previewHost.value?.querySelector<HTMLElement>(`[data-source-start='${line}']`);
    if (node) {
      node.scrollIntoView({ block: 'nearest' });
      // 自动定位产生的 scroll 事件会结束旧单元格编辑，待事件处理完再打开新编辑器。
      await new Promise<void>((resolve) => requestAnimationFrame(() => {
        requestAnimationFrame(() => resolve());
      }));
      if (store.documentRevision !== revision || store.filePath !== current.path
        || current.view.state.doc.toString() !== insertedSource) return;
      node = previewHost.value?.querySelector<HTMLElement>(`[data-source-start='${line}']`);
      if (!node) return;
      const editable = node.querySelector<HTMLElement>('th, td') || node;
      // 插入是明确的编辑操作，不能让原来的跨块选区阻止占位内容获得焦点。
      window.getSelection()?.removeAllRanges();
      if (editable.matches('th, td')) {
        domains.editor.events().openPreviewTableCellEditor({
          startLine: Number(node.dataset.sourceStart),
          endLine: Number(node.dataset.sourceEnd),
          rowIndex: 0,
          colIndex: 0,
        });
      } else editable.click();
      await nextTick();
      const textarea = domains.editor.presentation().previewInlineEditTextarea.value;
      if (textarea) {
        const placeholder = /标题|正文|代码|公式|引用|列表项|任务/.exec(textarea.value);
        if (placeholder) textarea.setSelectionRange(placeholder.index,
          placeholder.index + placeholder[0].length);
      }
    }
  }
  async function image() {
    const current = active();
    if (!current) { close(); return; }
    close();
    current.view.dispatch({ selection: { anchor: current.offset } });
    await insertImages();
  }
  function dismiss(event: Event) {
    if (event.target instanceof Element && event.target.closest('.preview-insert-menu')) return;
    close();
  }
  function keydown(event: KeyboardEvent) {
    if (!visible.value || event.key !== 'Escape') return;
    event.preventDefault();
    event.stopImmediatePropagation();
    close();
  }
  watch(() => [store.content, store.filePath, sourceMode.value], close, { flush: 'sync' });
  onMounted(() => {
    document.addEventListener('contextmenu', contextMenu, true);
    document.addEventListener('pointerdown', dismiss, true);
    document.addEventListener('scroll', dismiss, true);
    window.addEventListener('resize', close);
    window.addEventListener('blur', close);
    window.addEventListener('keydown', keydown, true);
  });
  onBeforeUnmount(() => {
    document.removeEventListener('contextmenu', contextMenu, true);
    document.removeEventListener('pointerdown', dismiss, true);
    document.removeEventListener('scroll', dismiss, true);
    window.removeEventListener('resize', close);
    window.removeEventListener('blur', close);
    window.removeEventListener('keydown', keydown, true);
  });
  return { visible, position, section, rows, columns, close, insert, image, reposition };
}
