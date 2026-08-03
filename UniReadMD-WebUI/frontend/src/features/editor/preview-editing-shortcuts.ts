import type { InlineShortcutKind } from '../../utils/markdown-shortcuts';
import { buildInlineShortcutPatch, indentMarkdownBlockLines } from '../../utils/markdown-shortcuts';
import type { usePreviewEditingSession } from './preview-editing-session';
import type { usePreviewInteractions } from './preview-interactions';
import type { usePreviewEditingState } from './preview-editing-state';
import type { usePreviewEditingEvents } from './preview-editing-events';
type UsePreviewEditingShortcutsDependencies =
  ReturnType<typeof usePreviewEditingSession> &
  ReturnType<typeof usePreviewInteractions> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof usePreviewEditingEvents>;

export function usePreviewEditingShortcuts(input: UsePreviewEditingShortcutsDependencies) {
  const { cancelPreviewInlineEdit,commitPreviewInlineEdit,previewInlineEditTextarea,
    previewInlineEditValue,previewInlineEditVisible,queuePreviewInlineEditSync,
    updatePreviewInlineEditorSelection,
  } = input;

  function applyPreviewInlineShortcut(kind: InlineShortcutKind) {
    const textarea = previewInlineEditTextarea.value;
    if (!previewInlineEditVisible.value || !textarea) {
      return false;
    }

    const value = previewInlineEditValue.value;
    const from = Math.min(textarea.selectionStart || 0, textarea.selectionEnd || 0);
    const to = Math.max(textarea.selectionStart || 0, textarea.selectionEnd || 0);
    const selectionText = value.slice(from, to);
    const patch = buildInlineShortcutPatch(selectionText, kind);
    if (!patch.changed) {
      return false;
    }

    const nextValue = `${value.slice(0, from)}${patch.insert}${value.slice(to)}`;
    previewInlineEditValue.value = nextValue;
    queuePreviewInlineEditSync();
    updatePreviewInlineEditorSelection(
      from + patch.selectionStart,
      from + patch.selectionEnd,
    );
    return true;
  }

  function resolvePreviewInlineOutdentPrefixLength(line: string) {
    if (line.startsWith('\t')) {
      return 1;
    }

    if (line.startsWith('  ')) {
      return 2;
    }

    if (line.startsWith(' ')) {
      return 1;
    }

    return 0;
  }

  function applyPreviewInlineIndentShortcut(outdent: boolean) {
    const textarea = previewInlineEditTextarea.value;
    if (!previewInlineEditVisible.value || !textarea) {
      return false;
    }

    const value = previewInlineEditValue.value;
    const from = Math.min(textarea.selectionStart || 0, textarea.selectionEnd || 0);
    const to = Math.max(textarea.selectionStart || 0, textarea.selectionEnd || 0);
    const lineStart = value.lastIndexOf('\n', Math.max(0, from - 1)) + 1;
    let lineEnd = value.indexOf('\n', to);
    if (lineEnd < 0) {
      lineEnd = value.length;
    }

    const block = value.slice(lineStart, lineEnd);
    const patch = indentMarkdownBlockLines(block, {
      outdent,
    });
    if (!patch.changed) {
      return false;
    }

    const nextValue = `${value.slice(0, lineStart)}${patch.text}${value.slice(lineEnd)}`;
    previewInlineEditValue.value = nextValue;
    queuePreviewInlineEditSync();

    if (from === to) {
      const currentLine = value.slice(lineStart, lineEnd).split('\n')[0] || '';
      const removedPrefix = outdent ? resolvePreviewInlineOutdentPrefixLength(currentLine) : 0;
      const nextCursor = outdent
        ? Math.max(lineStart, from - removedPrefix)
        : from + 2;
      updatePreviewInlineEditorSelection(nextCursor, nextCursor);
      return true;
    }

    updatePreviewInlineEditorSelection(lineStart, lineStart + patch.text.length);
    return true;
  }

  function handlePreviewInlineEditKeydown(event: KeyboardEvent) {
    if (!event.altKey && (event.ctrlKey || event.metaKey)) {
      const key = event.key.toLowerCase();
      if (key === 'b') {
        event.preventDefault();
        applyPreviewInlineShortcut('bold');
        return;
      }

      if (key === 'i') {
        event.preventDefault();
        applyPreviewInlineShortcut('italic');
        return;
      }

      if (key === 'k') {
        event.preventDefault();
        applyPreviewInlineShortcut('link');
        return;
      }

      if (key === ']') {
        event.preventDefault();
        applyPreviewInlineIndentShortcut(false);
        return;
      }

      if (key === '[') {
        event.preventDefault();
        applyPreviewInlineIndentShortcut(true);
        return;
      }
    }

    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
      event.preventDefault();
      commitPreviewInlineEdit();
      return;
    }

    if (event.key === 'Tab') {
      event.preventDefault();
      applyPreviewInlineIndentShortcut(event.shiftKey);
      return;
    }

    if (event.key === 'Escape') {
      event.preventDefault();
      cancelPreviewInlineEdit();
    }
  }

  return { applyPreviewInlineShortcut,resolvePreviewInlineOutdentPrefixLength,
    applyPreviewInlineIndentShortcut,handlePreviewInlineEditKeydown,
  };
}
