import type { WorkbenchDomains } from './workbench-domains';
import { htmlToMarkdown, markdownClipboardHtml } from './clipboard-markdown';

export function createClipboardWorkflows(
  domains: WorkbenchDomains, insertText: (text: string) => void,
) {
  const { runtime, errorMessage, previewHost } = domains.editor.foundation;
  let plainPasteUntil = 0;
  function clipboardKeys(event: KeyboardEvent) {
    if ((event.ctrlKey || event.metaKey) && event.shiftKey
      && event.key.toLowerCase() === 'v') plainPasteUntil = Date.now() + 1000;
  }
  function editableTarget(event: ClipboardEvent) {
    return event.target instanceof Element
      && event.target.closest('.editor-host, #typora-source, .preview-inline-editor');
  }
  function pasteRichText(event: ClipboardEvent) {
    if (!editableTarget(event) || event.defaultPrevented || Date.now() < plainPasteUntil) return;
    const html = event.clipboardData?.getData('text/html');
    if (!html) return;
    try {
      const markdown = htmlToMarkdown(html);
      if (!markdown) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      insertText(markdown);
    } catch (error) { errorMessage.value = String(error); }
  }
  function copyMarkdown(event: ClipboardEvent) {
    if (!event.clipboardData) return;
    let markdown = '';
    if (editableTarget(event)) {
      const target = event.target;
      if (target instanceof HTMLTextAreaElement) {
        markdown = target.value.slice(target.selectionStart, target.selectionEnd);
      } else if (runtime.editorView) {
        const { from, to } = runtime.editorView.state.selection.main;
        markdown = runtime.editorView.state.doc.sliceString(from, to);
      }
    } else {
      const selection = window.getSelection();
      if (!selection?.rangeCount || selection.isCollapsed
        || !previewHost.value?.contains(selection.anchorNode)
        || !previewHost.value?.contains(selection.focusNode)) return;
      const container = document.createElement('div');
      container.append(selection.getRangeAt(0).cloneContents());
      markdown = htmlToMarkdown(container.innerHTML);
    }
    if (!markdown) return;
    event.clipboardData.setData('text/plain', markdown);
    event.clipboardData.setData('text/html', markdownClipboardHtml(markdown));
    event.preventDefault();
    // 剪切仍由原生编辑控件处理；复制不改变文档和选区。
  }
  return { pasteRichText, copyMarkdown, clipboardKeys };
}
