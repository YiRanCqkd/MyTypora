import { hasMathSyntax } from '../../utils/markdown';
import { watch, nextTick } from 'vue';
import { observePreviewLayout, preservePreviewScroll } from './preview-editor-layout';
import type {
  EditorDomain,
  MarkdownDomain,
  SettingsDomain,
  WorkspaceDomain,
} from './workbench-domains';

export type ViewportLifecycleFeatures = Readonly<{
  editor: Readonly<{
    codeMirror: ReturnType<EditorDomain['codeMirror']>;
    foundation: EditorDomain['foundation'];
    navigation: ReturnType<EditorDomain['navigation']>;
    presentation: ReturnType<EditorDomain['presentation']>;
    preview: ReturnType<EditorDomain['preview']>;
    source: ReturnType<EditorDomain['source']>;
    table: ReturnType<EditorDomain['table']>;
  }>;
  markdown: Readonly<{ diagrams: ReturnType<MarkdownDomain['diagrams']> }>;
  settings: Readonly<{ foundation: ReturnType<SettingsDomain['foundation']> }>;
  workspace: Readonly<{
    actions: ReturnType<WorkspaceDomain['actions']>;
    foundation: ReturnType<WorkspaceDomain['foundation']>;
    presentation: ReturnType<WorkspaceDomain['presentation']>;
  }>;
}>;

export function useViewportLifecycle(input: ViewportLifecycleFeatures) {
  const { applyPreviewSearchHighlights,fontSizeCompartment,previewInlineEditComposing,
    previewInlineEditVisible,
    renderedHtml } = input.editor.presentation;
  const {
    applyStoreContentToEditor,createFontSizeExtension,refreshPreviewHeadingNodes,
    refreshPreviewSourceNodes,schedulePreviewScrollSync,scrollEditorToLine,scrollPreviewToLine,
    revealPreviewInlineEditorCaret,
  } = input.editor.navigation;
  const { closePreviewInlineEditor,closePreviewTableMenu,closeSourceContextMenu,
    refreshPreviewInlineEditorPosition } = input.editor.preview;
  const { closePreviewTableCellEditorAndCommit } = input.editor.table;
  const {
    contentZoom,fontSize,pinOutline,sourceMode,typewriterMode,
  } = input.settings.foundation;
  const { currentCursorLine,docSearchCaseSensitive,docSearchMatchIndex,docSearchQuery,
    docSearchWholeWord,editorPreviewHost,previewHost,runtime,showDocSearchPanel,store,
  } = input.editor.foundation;
  const {
    ensureMainPreviewWidth,getEffectiveEditorFontSize,
  } = input.workspace.presentation;
  const { ensureMathPluginLoaded,renderMermaidBlocks } = input.markdown.diagrams;
  const { refreshCurrentDirectoryFiles } = input.workspace.actions;
  const { refreshPreviewTableCellEditorPosition } = input.editor.source;
  const { scheduleTypewriterScroll } = input.editor.codeMirror;
  const { sidebarWidth } = input.workspace.foundation;

  // 容器可能因侧栏、工具栏或宿主布局改变，不能只依赖窗口 resize 事件。
  watch(() => editorPreviewHost.value, (container, _previous, onCleanup) => {
    if (!container) return;
    onCleanup(observePreviewLayout(container, () => {
      ensureMainPreviewWidth();
      refreshPreviewTableCellEditorPosition();
      refreshPreviewInlineEditorPosition();
      if (document.activeElement?.matches('.preview-inline-editor-input')) {
        revealPreviewInlineEditorCaret();
      }
    }));
  }, { immediate: true, flush: 'post' });

  function handleWindowResize() {
    ensureMainPreviewWidth();
    refreshPreviewHeadingNodes();
    refreshPreviewSourceNodes();
    refreshPreviewTableCellEditorPosition();
    refreshPreviewInlineEditorPosition();
    schedulePreviewScrollSync();
  }

  watch(
    () => [store.content, store.documentRevision] as const,
    ([nextContent]) => {
      if (hasMathSyntax(nextContent)) {
        void ensureMathPluginLoaded();
      }

      applyStoreContentToEditor(nextContent);
    },
  );

  watch(
    () => [fontSize.value, contentZoom.value],
    () => {
      if (!runtime.editorView) {
        return;
      }

      runtime.editorView.dispatch({
        effects: fontSizeCompartment.reconfigure(
          createFontSizeExtension(getEffectiveEditorFontSize()),
        ),
      });
    },
  );

  watch(
    () => contentZoom.value,
    () => {
      nextTick(() => {
        refreshPreviewHeadingNodes();
        refreshPreviewSourceNodes();
        refreshPreviewTableCellEditorPosition();
        if (!previewInlineEditComposing.value) {
          refreshPreviewInlineEditorPosition();
        }
        schedulePreviewScrollSync();
      });
    },
  );

  watch(
    () => renderedHtml.value,
    async (_html, _previous, onCleanup) => {
      closePreviewTableMenu();
      closeSourceContextMenu();

      // v-html 替换节点会先移除占位，必须在 DOM 更新之前保留滚动范围。
      const restoreScroll = previewInlineEditVisible.value && previewHost.value
        ? preservePreviewScroll(previewHost.value) : () => {};
      let cancelled = false;
      onCleanup(() => { cancelled = true; restoreScroll(); });
      try {
        await nextTick();
        if (cancelled) return;
        refreshPreviewInlineEditorPosition();
        restoreScroll();
        await renderMermaidBlocks();
        if (cancelled) return;
        refreshPreviewHeadingNodes();
        refreshPreviewSourceNodes();
        refreshPreviewTableCellEditorPosition();
        if (!previewInlineEditComposing.value) {
          refreshPreviewInlineEditorPosition();
          if (document.activeElement?.matches('.preview-inline-editor-input')) {
            revealPreviewInlineEditorCaret();
          }
        }
        applyPreviewSearchHighlights();
        schedulePreviewScrollSync();
      } finally {
        restoreScroll();
      }
    },
  );

  watch(
    () => [
      showDocSearchPanel.value,
      sourceMode.value,
      docSearchQuery.value,
      docSearchCaseSensitive.value,
      docSearchWholeWord.value,
      docSearchMatchIndex.value,
    ],
    async () => {
      await nextTick();
      applyPreviewSearchHighlights();
    },
  );

  watch(
    () => sourceMode.value,
    (enabled) => {
      closeSourceContextMenu();

      if (enabled) {
        closePreviewTableMenu();
        closePreviewTableCellEditorAndCommit();
        closePreviewInlineEditor();
        nextTick(() => {
          ensureMainPreviewWidth();
          scrollEditorToLine(currentCursorLine.value);
          if (typewriterMode.value && runtime.editorView) {
            scheduleTypewriterScroll(runtime.editorView);
          }
        });
        return;
      }

      nextTick(() => {
        ensureMainPreviewWidth();
        scrollPreviewToLine(currentCursorLine.value);
        schedulePreviewScrollSync();
      });
    },
  );

  watch(
    () => store.filePath,
    (nextFilePath) => {
      void refreshCurrentDirectoryFiles(nextFilePath);
    },
    {
      immediate: true,
    },
  );

  watch(
    () => [sidebarWidth.value, pinOutline.value],
    () => {
      nextTick(() => {
        if (sourceMode.value) ensureMainPreviewWidth();
        else refreshPreviewInlineEditorPosition();
      });
    },
  );

  return { handleWindowResize,
  };
}
