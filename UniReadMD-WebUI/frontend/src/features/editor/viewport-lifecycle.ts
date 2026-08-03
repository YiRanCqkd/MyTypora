import { hasMathSyntax } from '../../utils/markdown';
import { watch, nextTick } from 'vue';
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
    renderedHtml } = input.editor.presentation;
  const {
    applyStoreContentToEditor,createFontSizeExtension,refreshPreviewHeadingNodes,
    refreshPreviewSourceNodes,schedulePreviewScrollSync,scrollEditorToLine,scrollPreviewToLine,
  } = input.editor.navigation;
  const { closePreviewInlineEditor,closePreviewTableMenu,closeSourceContextMenu,
    refreshPreviewInlineEditorPosition } = input.editor.preview;
  const { closePreviewTableCellEditorAndCommit } = input.editor.table;
  const {
    contentZoom,fontSize,pinOutline,sourceMode,typewriterMode,
  } = input.settings.foundation;
  const { currentCursorLine,docSearchCaseSensitive,docSearchMatchIndex,docSearchQuery,
    docSearchWholeWord,runtime,showDocSearchPanel,store } = input.editor.foundation;
  const {
    ensureMainPreviewWidth,getEffectiveEditorFontSize,
  } = input.workspace.presentation;
  const { ensureMathPluginLoaded,renderMermaidBlocks } = input.markdown.diagrams;
  const { refreshCurrentDirectoryFiles } = input.workspace.actions;
  const { refreshPreviewTableCellEditorPosition } = input.editor.source;
  const { scheduleTypewriterScroll } = input.editor.codeMirror;
  const { sidebarWidth } = input.workspace.foundation;

  function handleWindowResize() {
    ensureMainPreviewWidth();
    refreshPreviewHeadingNodes();
    refreshPreviewSourceNodes();
    refreshPreviewTableCellEditorPosition();
    schedulePreviewScrollSync();
  }

  watch(
    () => store.content,
    (nextContent) => {
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
    async () => {
      closePreviewTableMenu();
      closeSourceContextMenu();

      await nextTick();
      await renderMermaidBlocks();
      refreshPreviewHeadingNodes();
      refreshPreviewSourceNodes();
      refreshPreviewTableCellEditorPosition();
      if (!previewInlineEditComposing.value) {
        refreshPreviewInlineEditorPosition();
      }
      applyPreviewSearchHighlights();
      schedulePreviewScrollSync();
    },
  );

  watch(
    () => [
      showDocSearchPanel.value,
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
      if (!sourceMode.value) {
        return;
      }

      nextTick(() => {
        ensureMainPreviewWidth();
      });
    },
  );

  return { handleWindowResize,
  };
}
