import { bridgeService } from '../../services/bridge';
import { hasMathSyntax } from '../../utils/markdown';
import { onMounted, nextTick, onBeforeUnmount, watch } from 'vue';
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
    table: ReturnType<EditorDomain['table']>;
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
    docSearchWholeWord,docSearchRegex,errorMessage,runtime,showDocReplacePanel,
    showDocSearchPanel,store,
  } = input.editor.foundation;
  const { applyTheme,applyUserCss,contentZoom,focusMode,fontSize,isNativeWindow,pinOutline,
    sourceMode,spellCheckEnabled,spellCheckLanguage,theme,typewriterMode,userCssEnabled,
    userCssPath,autoSave,autoSaveInterval,
  } = input.settings.foundation;
  const { clearContentZoomPersistTimer,loadRecentFilesFromNative,loadSessionSnapshot,
    persistSessionSnapshot,
    buildSessionSnapshot,
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
    workspaceDirectory,
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

    const snapshot = await loadSessionSnapshot();
    if (snapshot) {
      store.restoreSession({
        filePath: snapshot.filePath || '',
        content: snapshot.content,
        isDirty: Boolean(snapshot.isDirty),
        cursorHint: Number(snapshot.cursorHint || 0),
      });
      if (snapshot.isDirty) window.dispatchEvent(new CustomEvent('unireadmd-recovery'));
      snapshotRecentFiles = Array.isArray(snapshot.recentFiles)
        ? snapshot.recentFiles.filter((item) => typeof item === 'string')
        : [];
      snapshotContentZoom = normalizeContentZoom(snapshot.contentZoom ?? CONTENT_ZOOM_DEFAULT);
      recentFiles.value = normalizeRecentFiles(snapshotRecentFiles);
      contentZoom.value = snapshotContentZoom;
      pinOutline.value = typeof snapshot.pinOutline === 'boolean' ? snapshot.pinOutline : true;
      sourceMode.value = Boolean(snapshot.sourceMode);
      focusMode.value = Boolean(snapshot.focusMode);
      typewriterMode.value = Boolean(snapshot.typewriterMode);
      sidebarWidth.value = normalizeSidebarWidth(snapshot.sidebarWidth ?? SIDEBAR_WIDTH_DEFAULT);
      mainPreviewWidth.value = Number(snapshot.mainPreviewWidth || 0);
      fileViewMode.value = snapshot.fileViewMode === 'list' ? 'list' : 'tree';
      fileSortMode.value = ['natural', 'name', 'date'].includes(snapshot.fileSortMode || '')
        ? snapshot.fileSortMode || 'natural' : 'natural';
      filesFilter.value = snapshot.filesFilter || '';
      showDocSearchPanel.value = Boolean(snapshot.searchPanelVisible);
      showDocReplacePanel.value = Boolean(snapshot.replacePanelVisible);
      docSearchQuery.value = snapshot.searchQuery || '';
      docReplaceQuery.value = snapshot.replaceQuery || '';
      docSearchCaseSensitive.value = Boolean(snapshot.searchCaseSensitive);
      docSearchWholeWord.value = Boolean(snapshot.searchWholeWord);
      docSearchRegex.value = Boolean(snapshot.searchRegex);
      docSearchMatchIndex.value = Number(snapshot.searchMatchIndex || 0);
      fileTreeExpanded.value = snapshot.fileTreeExpanded || {};
      workspaceDirectory.value = snapshot.workspaceDirectory || '';
    }

    createEditor();
    runtime.flushEdits = () => {
      runtime.flushVisualEdits?.();
      input.editor.preview.closePreviewInlineEditor();
      input.editor.table.closePreviewTableCellEditorAndCommit();
    };
    let snapshotTimer = 0;
    runtime.flushSnapshot = async () => {
      window.clearTimeout(snapshotTimer);
      await persistSessionSnapshot();
    };
    const stopSnapshotWatch = watch(buildSessionSnapshot, () => {
      window.clearTimeout(snapshotTimer);
      snapshotTimer = window.setTimeout(() => { void persistSessionSnapshot(); }, 1000);
    }, { deep: true });
    runtime.snapshotCleanup = () => {
      stopSnapshotWatch();
      window.clearTimeout(snapshotTimer);
    };
    let polling = false;
    const directoryTimer = window.setInterval(async () => {
      if (polling || !isNativeWindow.value || !workspaceDirectory.value) return;
      polling = true;
      try { await input.workspace.actions.refreshCurrentDirectoryFiles(store.filePath); }
      finally { polling = false; }
    }, 5000);
    runtime.directoryCleanup = () => window.clearInterval(directoryTimer);
    const beforeUnload = (event: BeforeUnloadEvent) => {
      runtime.flushEdits?.();
      persistSessionSnapshot();
      if (store.isDirty && !isNativeWindow.value) {
        event.preventDefault();
        event.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', beforeUnload);
    runtime.beforeUnloadCleanup = () => window.removeEventListener('beforeunload', beforeUnload);
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
      autoSave.value = settingsResult.settings.autoSave;
      autoSaveInterval.value = settingsResult.settings.autoSaveInterval || 300;
      input.settings.foundation.documentToolsVisible.value =
        settingsResult.settings.documentToolsVisible === true;
      input.settings.foundation.preferencesVisible.value =
        settingsResult.settings.preferencesVisible === true;
      let lastAutoSave = Date.now();
      const autoSaveTimer = window.setInterval(() => {
        if (autoSave.value && store.isDirty && store.filePath
          && Date.now() - lastAutoSave >= autoSaveInterval.value * 1000) {
          lastAutoSave = Date.now();
          void input.editor.foundation.saveCurrentDocument();
        }
      }, 1000);
      runtime.autoSaveCleanup = () => window.clearInterval(autoSaveTimer);
    } else {
      contentZoom.value = snapshotContentZoom;
    }

    if (isNativeWindow.value) {
      await loadRecentFilesFromNative(snapshotRecentFiles);
    }

    await refreshExportCapabilities();
    await refreshSpellCheckState();
    await tryOpenLaunchFileFromNative();
    if (workspaceDirectory.value) {
      await input.workspace.actions.refreshCurrentDirectoryFiles(store.filePath);
    }
    if (snapshot?.filePath && isNativeWindow.value && store.filePath === snapshot.filePath) {
      const disk = await runTask('恢复文件资源', () => {
        return bridgeService.openFromPath(snapshot.filePath);
      });
      if (disk) {
        store.resourceBaseUrl = disk.resourceBaseUrl || '';
        store.version = snapshot.isDirty && snapshot.version
          ? snapshot.version : disk.version || '';
        store.savedContent = disk.content;
        if (!snapshot.isDirty) store.setContent(disk.content);
        store.isDirty = store.content !== disk.content;
      }
    }
    applyTheme(theme.value);
    const css = await runTask('加载文档样式', async () => {
      return window.nativeBridge.settings.getUserCss?.() || { css: '' };
    });
    applyUserCss(css?.css || '');
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
    runtime.beforeUnloadCleanup?.();
    runtime.autoSaveCleanup?.();
    runtime.directoryCleanup?.();
    runtime.snapshotCleanup?.();
    window.removeEventListener('keydown', handleGlobalKeydown);
    window.removeEventListener('mousedown', handleGlobalPointerDown);
    window.removeEventListener('wheel', handleGlobalWheel);
    window.removeEventListener('resize', handleWindowResize);
    persistSessionSnapshot();
    destroyEditor();
  });

}
