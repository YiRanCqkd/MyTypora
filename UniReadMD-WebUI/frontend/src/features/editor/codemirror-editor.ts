import { EditorState } from '@codemirror/state';
import {
  EditorView,
  highlightActiveLine,
  highlightActiveLineGutter,
  keymap,
} from '@codemirror/view';
import { history, indentWithTab, defaultKeymap, historyKeymap } from '@codemirror/commands';
import { markdown as markdownLanguage } from '@codemirror/lang-markdown';
import { syntaxHighlighting } from '@codemirror/language';
import { closeBrackets, closeBracketsKeymap } from '@codemirror/autocomplete';
import type { useSourceEditing } from './source-editing';
import type { useEditorSynchronization } from './editor-synchronization';
import type { useEditorHosts } from './editor-hosts';
import type { usePreviewEditingState } from './preview-editing-state';
import type { createLayoutMetrics } from '../workspace/layout-metrics';
import type { createWorkbenchRuntime } from './workbench-runtime';
import type { createMarkdownRenderer } from '../markdown/renderer';
import type { useWorkbenchStore } from './editor-store';
import type { createPreviewEditingRanges } from './preview-editing-ranges';
import type { useViewOptions } from '../settings/view-options';
type UseCodeMirrorEditorDependencies =
  ReturnType<typeof useSourceEditing> &
  ReturnType<typeof useEditorSynchronization> &
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createLayoutMetrics> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof createMarkdownRenderer> &
  ReturnType<typeof useWorkbenchStore> &
  ReturnType<typeof createPreviewEditingRanges> &
  ReturnType<typeof useViewOptions>;

export function useCodeMirrorEditor(input: UseCodeMirrorEditorDependencies) {
  const { applySourceIndentShortcut,applySourceInlineShortcut,createFontSizeExtension,
    editorHost,fontSizeCompartment,getEffectiveEditorFontSize,runtime,sourceHighlightStyle,
    store,syncCursorLineFromState,typewriterMode,
  } = input;

  function scheduleTypewriterScroll(view: EditorView) {
    if (!typewriterMode.value) {
      return;
    }

    if (runtime.typewriterScrollFrame) {
      cancelAnimationFrame(runtime.typewriterScrollFrame);
    }

    runtime.typewriterScrollFrame = requestAnimationFrame(() => {
      runtime.typewriterScrollFrame = 0;
      if (!typewriterMode.value) {
        return;
      }

      const scroller = view.scrollDOM;
      const cursorRect = view.coordsAtPos(view.state.selection.main.head);
      if (!cursorRect) {
        return;
      }

      const viewportRect = scroller.getBoundingClientRect();
      const cursorHeight = Math.max(0, cursorRect.bottom - cursorRect.top);
      const cursorCenterY = cursorRect.top + cursorHeight / 2;
      const targetCenterY = viewportRect.top + viewportRect.height / 2;
      const delta = cursorCenterY - targetCenterY;

      if (Math.abs(delta) < 2) {
        return;
      }

      scroller.scrollTop += delta;
    });
  }

  function createEditor() {
    if (!editorHost.value || runtime.editorView) {
      return;
    }

    const extensions = [
      history(),
      closeBrackets(),
      markdownLanguage(),
      syntaxHighlighting(sourceHighlightStyle),
      highlightActiveLine(),
      highlightActiveLineGutter(),
      EditorView.lineWrapping,
      keymap.of([
        {
          key: 'Mod-b',
          run: (view) => applySourceInlineShortcut('bold', view),
        },
        {
          key: 'Mod-i',
          run: (view) => applySourceInlineShortcut('italic', view),
        },
        {
          key: 'Mod-k',
          run: (view) => applySourceInlineShortcut('link', view),
        },
        {
          key: 'Mod-]',
          run: (view) => applySourceIndentShortcut(false, view),
        },
        {
          key: 'Mod-[',
          run: (view) => applySourceIndentShortcut(true, view),
        },
        {
          key: 'Shift-Tab',
          run: (view) => applySourceIndentShortcut(true, view),
        },
        indentWithTab,
        ...closeBracketsKeymap,
        ...defaultKeymap,
        ...historyKeymap,
      ]),
      fontSizeCompartment.of(createFontSizeExtension(getEffectiveEditorFontSize())),
      EditorView.updateListener.of((update) => {
        if (update.selectionSet || update.docChanged) {
          syncCursorLineFromState(update.state);
          scheduleTypewriterScroll(update.view);
        }

        if (!update.docChanged || runtime.isSyncingFromStore) {
          return;
        }

        const nextContent = update.state.doc.toString();
        store.setContent(nextContent);
      }),
    ];
    const state = EditorState.create({ doc: store.content, extensions });
    runtime.documentRevision = store.documentRevision;
    runtime.resetDocument = (content) => {
      runtime.editorView?.setState(EditorState.create({ doc: content, extensions }));
    };

    runtime.editorView = new EditorView({
      state,
      parent: editorHost.value,
    });

    syncCursorLineFromState(runtime.editorView.state);
  }

  function destroyEditor() {
    if (runtime.previewScrollSyncFrame) {
      cancelAnimationFrame(runtime.previewScrollSyncFrame);
      runtime.previewScrollSyncFrame = 0;
    }

    if (runtime.typewriterScrollFrame) {
      cancelAnimationFrame(runtime.typewriterScrollFrame);
      runtime.typewriterScrollFrame = 0;
    }

    if (!runtime.editorView) {
      return;
    }

    runtime.editorView.destroy();
    runtime.editorView = null;
  }

  return { scheduleTypewriterScroll,createEditor,destroyEditor,
  };
}
