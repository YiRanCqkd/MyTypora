import { bridgeService } from '../../services/bridge';
import { hasMathSyntax } from '../../utils/markdown';
import { onMounted, nextTick, onBeforeUnmount } from 'vue';
import { resolveFeatureInitializationError } from './feature-initialization';
import type {
  EditorDomain,
  ExportDomain,
  MarkdownDomain,
  SettingsDomain,
  WorkspaceDomain,
} from './workbench-domains';
import type { useViewportLifecycle } from './viewport-lifecycle';

export type WorkbenchLifecycleFeatures = Readonly<{
  editor: Readonly<{
    codeMirror: ReturnType<EditorDomain['codeMirror']>;
    commands: ReturnType<EditorDomain['commands']>;
    foundation: EditorDomain['foundation'];
    navigation: ReturnType<EditorDomain['navigation']>;
    preview: ReturnType<EditorDomain['preview']>;
  }>;
  export: ReturnType<ExportDomain['feature']>;
  markdown: Readonly<{
    diagrams: ReturnType<MarkdownDomain['diagrams']>;
    renderer: ReturnType<MarkdownDomain['feature']>;
  }>;
  settings: Readonly<{
    actions: ReturnType<SettingsDomain['actions']>;
    foundation: ReturnType<SettingsDomain['foundation']>;
    nativeCommands: ReturnType<SettingsDomain['nativeCommands']>;
    session: ReturnType<SettingsDomain['session']>;
  }>;
  workspace: Readonly<{
    actions: ReturnType<WorkspaceDomain['actions']>;
    foundation: ReturnType<WorkspaceDomain['foundation']>;
    presentation: ReturnType<WorkspaceDomain['presentation']>;
  }>;
}>;

export function useWorkbenchLifecycle(
  input: WorkbenchLifecycleFeatures,
  viewport: ReturnType<typeof useViewportLifecycle>,
) {
  const { CONTENT_ZOOM_DEFAULT,SIDEBAR_WIDTH_DEFAULT } = input.markdown.renderer;
  const {
    appVersion,docReplaceQuery,docSearchCaseSensitive,docSearchMatchIndex,docSearchQuery,
    docSearchWholeWord,errorMessage,runtime,showDocReplacePanel,showDocSearchPanel,store,
  } = input.editor.foundation;
  const { applyTheme,applyUserCss,contentZoom,focusMode,fontSize,isNativeWindow,pinOutline,
    sourceMode,spellCheckEnabled,spellCheckLanguage,theme,typewriterMode,userCssEnabled,
    userCssPath,
  } = input.settings.foundation;
  const { clearContentZoomPersistTimer,loadRecentFilesFromNative,loadSessionSnapshot,
    normalizeLegacyUserCssSettings,persistSessionSnapshot,
  } = input.settings.session;
  const { clearPreviewInlineEditSyncTimer } = input.editor.preview;
  const { createEditor,destroyEditor } = input.editor.codeMirror;
  const { detectNativeWindowMode,refreshSpellCheckState } = input.settings.actions;
  const { ensureMainPreviewWidth,normalizeContentZoom,normalizeRecentFiles,
    normalizeSidebarWidth,runTask,
  } = input.workspace.presentation;
  const { ensureMathPluginLoaded,renderMermaidBlocks } = input.markdown.diagrams;
  const { fileSortMode,fileTreeExpanded,fileViewMode,filesFilter,mainPreviewWidth,recentFiles,
    sidebarWidth,
  } = input.workspace.foundation;
  const {
    handleGlobalKeydown,handleGlobalPointerDown,handleGlobalWheel,
  } = input.editor.commands;
  const { handleMenuCommand } = input.settings.nativeCommands;
  const {
    handleNativeLaunchFile,tryOpenLaunchFileFromNative,
  } = input.workspace.actions;
  const { handleWindowResize } = viewport;
  const { refreshExportCapabilities } = input.export;
  const {
    refreshPreviewHeadingNodes,refreshPreviewSourceNodes,schedulePreviewScrollSync,
  } = input.editor.navigation;

  onMounted(async () => {
    Object.defineProperty(window, '__unireadmdMounted', {
      configurable: true,
      writable: true,
      value: true,
    });

    const initializationError = resolveFeatureInitializationError(window.nativeBridge);
    if (initializationError) {
      errorMessage.value = initializationError.toDisplayMessage();
      return;
    }

    isNativeWindow.value = detectNativeWindowMode();
    let snapshotRecentFiles: string[] = [];
    let snapshotContentZoom = CONTENT_ZOOM_DEFAULT;

    const snapshot = loadSessionSnapshot();
    if (snapshot) {
      store.restoreSession({
        filePath: snapshot.filePath || '',
        content: snapshot.content,
        isDirty: Boolean(snapshot.isDirty),
        cursorHint: Number(snapshot.cursorHint || 0),
      });
      snapshotRecentFiles = Array.isArray(snapshot.recentFiles)
        ? snapshot.recentFiles.filter((item) => typeof item === 'string')
        : [];
      snapshotContentZoom = normalizeContentZoom(snapshot.contentZoom ?? CONTENT_ZOOM_DEFAULT);
      recentFiles.value = normalizeRecentFiles(snapshotRecentFiles);
      contentZoom.value = snapshotContentZoom;
      pinOutline.value = typeof snapshot.pinOutline === 'boolean' ? snapshot.pinOutline : true;
      sourceMode.value = false;
      focusMode.value = Boolean(snapshot.focusMode);
      typewriterMode.value = Boolean(snapshot.typewriterMode);
      sidebarWidth.value = normalizeSidebarWidth(snapshot.sidebarWidth ?? SIDEBAR_WIDTH_DEFAULT);
      mainPreviewWidth.value = Number(snapshot.mainPreviewWidth || 0);
      fileViewMode.value = 'tree';
      fileSortMode.value = 'natural';
      filesFilter.value = '';
      showDocSearchPanel.value = false;
      showDocReplacePanel.value = false;
      docSearchQuery.value = snapshot.searchQuery || '';
      docReplaceQuery.value = snapshot.replaceQuery || '';
      docSearchCaseSensitive.value = Boolean(snapshot.searchCaseSensitive);
      docSearchWholeWord.value = Boolean(snapshot.searchWholeWord);
      docSearchMatchIndex.value = Number(snapshot.searchMatchIndex || 0);
      fileTreeExpanded.value = snapshot.fileTreeExpanded || {};
    }

    createEditor();
    window.addEventListener('beforeunload', persistSessionSnapshot);
    window.addEventListener('keydown', handleGlobalKeydown);
    window.addEventListener('mousedown', handleGlobalPointerDown);
    window.addEventListener('wheel', handleGlobalWheel, {
      passive: false,
    });
    window.addEventListener('resize', handleWindowResize);
    runtime.menuCommandCleanup = window.nativeBridge.menu?.onCommand((command) => {
      void handleMenuCommand(command);
    }) || null;
    runtime.launchFileCleanup = bridgeService.onLaunchFile((payload) => {
      const filePath = payload?.filePath || '';
      if (!filePath) {
        return;
      }

      void handleNativeLaunchFile(filePath);
    });

    if (hasMathSyntax(store.content)) {
      await ensureMathPluginLoaded();
    }

    const versionResult = await runTask('app-version', (traceId) => {
      return bridgeService.getVersion(traceId);
    });

    if (versionResult) {
      appVersion.value = versionResult.version;
    }

    const settingsResult = await runTask('settings-load', (traceId) => {
      return bridgeService.getAllSettings(traceId);
    });

    if (settingsResult) {
      theme.value = settingsResult.settings.theme;
      fontSize.value = settingsResult.settings.fontSize;
      contentZoom.value = normalizeContentZoom(settingsResult.settings.contentZoom);
      userCssEnabled.value = Boolean(settingsResult.settings.userCssEnabled);
      userCssPath.value = String(settingsResult.settings.userCssPath || '');
      spellCheckEnabled.value = settingsResult.settings.spellCheckEnabled;
      spellCheckLanguage.value = settingsResult.settings.spellCheckLanguage;
    } else {
      contentZoom.value = snapshotContentZoom;
    }

    if (isNativeWindow.value) {
      await loadRecentFilesFromNative(snapshotRecentFiles);
    }

    await normalizeLegacyUserCssSettings();
    await refreshExportCapabilities();
    await refreshSpellCheckState();
    await tryOpenLaunchFileFromNative();
    applyTheme(theme.value);
    applyUserCss();
    await nextTick();
    ensureMainPreviewWidth();
    await renderMermaidBlocks();
    refreshPreviewHeadingNodes();
    refreshPreviewSourceNodes();
    schedulePreviewScrollSync();
  });

  onBeforeUnmount(() => {
    if (runtime.sidebarResizeCleanup) {
      runtime.sidebarResizeCleanup();
    }
    if (runtime.mainPreviewResizeCleanup) {
      runtime.mainPreviewResizeCleanup();
    }
    if (runtime.menuCommandCleanup) {
      runtime.menuCommandCleanup();
    }
    if (runtime.launchFileCleanup) {
      runtime.launchFileCleanup();
    }

    clearContentZoomPersistTimer();
    clearPreviewInlineEditSyncTimer();
    window.removeEventListener('beforeunload', persistSessionSnapshot);
    window.removeEventListener('keydown', handleGlobalKeydown);
    window.removeEventListener('mousedown', handleGlobalPointerDown);
    window.removeEventListener('wheel', handleGlobalWheel);
    window.removeEventListener('resize', handleWindowResize);
    persistSessionSnapshot();
    destroyEditor();
  });

}
