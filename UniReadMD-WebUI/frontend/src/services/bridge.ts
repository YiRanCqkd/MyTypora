import type {
  AppVersionResult,
  ExportCapabilitiesResult,
  ExportDocumentPayload,
  ExportDocumentResult,
  ExportHtmlPayload,
  ExportHtmlResult,
  FileTreeContextMenuPayload,
  LaunchFilePathResult,
  LaunchFileEventPayload,
  ListDirectoryMarkdownResult,
  OpenFileResult,
  RecentFilesResult,
  SaveFilePayload,
  SaveFileResult,
  SelectUserCssFileResult,
  SourceEditorContextMenuPayload,
  SpellCheckStateResult,
  SettingsGetAllResult,
  SettingsKey,
  SettingsSetResult,
} from '../types/native';

type NativeBridge = Window['nativeBridge'];

function getNativeBridge(): NativeBridge {
  if (!window.nativeBridge) {
    throw new Error('nativeBridge 不可用');
  }

  return window.nativeBridge;
}

function ensureStringField(
  value: unknown,
  field: string,
  context: string,
): string {
  if (typeof value !== 'string') {
    throw new Error(`${context} 返回非法字段: ${field}`);
  }

  return value;
}

function ensureOpenFileResult(
  value: OpenFileResult | null,
  context: string,
): OpenFileResult | null {
  if (value === null) {
    return null;
  }

  return {
    filePath: ensureStringField(value.filePath, 'filePath', context),
    content: ensureStringField(value.content, 'content', context),
  };
}

function ensureLaunchFilePathResult(
  value: LaunchFilePathResult | null,
  context: string,
): LaunchFilePathResult | null {
  if (value === null) {
    return null;
  }

  return {
    filePath: ensureStringField(value.filePath, 'filePath', context),
  };
}

function ensureLaunchFileEventPayload(
  value: LaunchFileEventPayload | null,
  context: string,
): LaunchFileEventPayload | null {
  if (value === null) {
    return null;
  }

  return {
    filePath: ensureStringField(value.filePath, 'filePath', context),
  };
}

function ensureSaveFileResult(
  value: SaveFileResult | null,
  context: string,
): SaveFileResult | null {
  if (value === null) {
    return null;
  }

  return {
    filePath: ensureStringField(value.filePath, 'filePath', context),
  };
}

function ensureExportHtmlResult(
  value: ExportHtmlResult | null,
  context: string,
): ExportHtmlResult | null {
  if (value === null) {
    return null;
  }

  return {
    filePath: ensureStringField(value.filePath, 'filePath', context),
  };
}

function ensureExportFormatField(
  value: unknown,
  field: string,
  context: string,
) {
  const text = ensureStringField(value, field, context);
  return text;
}

function ensureExportDocumentResult(
  value: ExportDocumentResult | null,
  context: string,
): ExportDocumentResult | null {
  if (value === null) {
    return null;
  }

  const engine = ensureStringField(value.engine, 'engine', context);
  if (engine !== 'builtin' && engine !== 'pandoc') {
    throw new Error(`${context} 返回非法字段: engine`);
  }

  return {
    filePath: ensureStringField(value.filePath, 'filePath', context),
    format: ensureExportFormatField(value.format, 'format', context) as ExportDocumentResult['format'],
    engine,
  };
}

function ensureExportCapabilitiesResult(
  value: ExportCapabilitiesResult,
  context: string,
): ExportCapabilitiesResult {
  if (!value || typeof value !== 'object') {
    throw new Error(`${context} 返回非法数据`);
  }

  const normalizeFormatList = (raw: unknown) => {
    if (!Array.isArray(raw)) {
      return [];
    }

    return raw.filter((item) => typeof item === 'string');
  };

  return {
    builtinFormats: normalizeFormatList(value.builtinFormats) as ExportCapabilitiesResult['builtinFormats'],
    pandocAvailable: Boolean(value.pandocAvailable),
    pandocVersion: typeof value.pandocVersion === 'string' ? value.pandocVersion : '',
    pandocFormats: normalizeFormatList(value.pandocFormats) as ExportCapabilitiesResult['pandocFormats'],
  };
}

function ensureSettingsResult(value: SettingsGetAllResult): SettingsGetAllResult {
  if (!value || typeof value !== 'object' || !value.settings) {
    throw new Error('settings:getAll 返回非法数据');
  }

  return value;
}

function ensureRecentFilesResult(
  value: RecentFilesResult,
  context: string,
): RecentFilesResult {
  if (!value || typeof value !== 'object') {
    throw new Error(`${context} 返回非法数据`);
  }

  return {
    filePaths: Array.isArray(value.filePaths)
      ? value.filePaths.filter((item) => typeof item === 'string')
      : [],
  };
}

function ensureDirectoryMarkdownResult(
  value: ListDirectoryMarkdownResult,
): ListDirectoryMarkdownResult {
  if (!value || typeof value !== 'object') {
    throw new Error('file:listMarkdownInDirectory 返回非法数据');
  }

  const directoryPath = ensureStringField(
    value.directoryPath,
    'directoryPath',
    'file:listMarkdownInDirectory',
  );

  const filePaths = Array.isArray(value.filePaths)
    ? value.filePaths.filter((item) => typeof item === 'string')
    : [];

  return {
    directoryPath,
    filePaths,
  };
}

function ensureUserCssSelection(
  value: SelectUserCssFileResult | null,
): SelectUserCssFileResult | null {
  if (value === null) {
    return null;
  }

  return {
    filePath: ensureStringField(value.filePath, 'filePath', 'settings:selectUserCssFile'),
  };
}

function ensureSpellCheckStateResult(
  value: SpellCheckStateResult,
  context: string,
): SpellCheckStateResult {
  if (!value || typeof value !== 'object') {
    throw new Error(`${context} 返回非法数据`);
  }

  const enabled = Boolean(value.enabled);
  const language = ensureStringField(value.language, 'language', context);
  const availableLanguages = Array.isArray(value.availableLanguages)
    ? value.availableLanguages.filter((item) => typeof item === 'string')
    : [];
  const dictionaryStatus = value.dictionaryStatus === 'missing' ? 'missing' : 'ready';

  return {
    enabled,
    language,
    availableLanguages,
    dictionaryStatus,
  };
}

export const bridgeService = {
  async getVersion(traceId?: string): Promise<AppVersionResult> {
    const result = await getNativeBridge().app.getVersion(traceId);
    return {
      version: ensureStringField(result?.version, 'version', 'app:getVersion'),
    };
  },
  async consumeLaunchFilePath(traceId?: string): Promise<LaunchFilePathResult | null> {
    const result = await getNativeBridge().app.consumeLaunchFilePath(traceId);
    return ensureLaunchFilePathResult(result, 'app:consumeLaunchFilePath');
  },
  onLaunchFile(listener: (payload: LaunchFileEventPayload | null) => void): () => void {
    return getNativeBridge().app.onLaunchFile((payload) => {
      listener(ensureLaunchFileEventPayload(payload, 'app:onLaunchFile'));
    });
  },
  async openMarkdown(traceId?: string): Promise<OpenFileResult | null> {
    const result = await getNativeBridge().file.openMarkdown({ traceId });
    return ensureOpenFileResult(result, 'file:openMarkdown');
  },
  async openFromPath(
    filePath: string,
    traceId?: string,
  ): Promise<OpenFileResult | null> {
    const payload = {
      filePath,
      traceId,
    };
    const result = await getNativeBridge().file.openFromPath(payload);
    return ensureOpenFileResult(result, 'file:openFromPath');
  },
  async saveMarkdown(payload: SaveFilePayload): Promise<SaveFileResult | null> {
    const result = await getNativeBridge().file.saveMarkdown(payload);
    return ensureSaveFileResult(result, 'file:saveMarkdown');
  },
  getPathForFile(file: File): string {
    const fileApi = getNativeBridge().file;
    if (typeof fileApi.getPathForFile !== 'function') {
      return '';
    }

    try {
      return fileApi.getPathForFile(file);
    } catch {
      return '';
    }
  },
  async listMarkdownInDirectory(
    filePath: string,
    traceId?: string,
  ): Promise<ListDirectoryMarkdownResult> {
    const result = await getNativeBridge().file.listMarkdownInDirectory({
      filePath,
      traceId,
    });
    return ensureDirectoryMarkdownResult(result);
  },
  async exportHtml(payload: ExportHtmlPayload): Promise<ExportHtmlResult | null> {
    const result = await getNativeBridge().exporter.exportHtml(payload);
    return ensureExportHtmlResult(result, 'export:html');
  },
  async exportDocument(
    payload: ExportDocumentPayload,
  ): Promise<ExportDocumentResult | null> {
    const result = await getNativeBridge().exporter.exportDocument(payload);
    return ensureExportDocumentResult(result, 'export:document');
  },
  async getExportCapabilities(traceId?: string): Promise<ExportCapabilitiesResult> {
    const result = await getNativeBridge().exporter.getCapabilities(traceId);
    return ensureExportCapabilitiesResult(result, 'export:getCapabilities');
  },
  async getAllSettings(traceId?: string): Promise<SettingsGetAllResult> {
    const result = await getNativeBridge().settings.getAll(traceId);
    return ensureSettingsResult(result);
  },
  async getRecentFiles(traceId?: string): Promise<RecentFilesResult> {
    const result = await getNativeBridge().settings.getRecentFiles(traceId);
    return ensureRecentFilesResult(result, 'settings:getRecentFiles');
  },
  async setSetting(
    key: SettingsKey,
    value: unknown,
    traceId?: string,
  ): Promise<SettingsSetResult> {
    return getNativeBridge().settings.set(key, value as never, traceId);
  },
  async setRecentFiles(
    filePaths: string[],
    traceId?: string,
  ): Promise<RecentFilesResult> {
    const result = await getNativeBridge().settings.setRecentFiles(filePaths, traceId);
    return ensureRecentFilesResult(result, 'settings:setRecentFiles');
  },
  async selectUserCssFile(
    traceId?: string,
  ): Promise<SelectUserCssFileResult | null> {
    const result = await getNativeBridge().settings.selectUserCssFile(traceId);
    return ensureUserCssSelection(result);
  },
  async minimizeWindow(traceId?: string): Promise<{ ok: true }> {
    return getNativeBridge().window.minimize(traceId);
  },
  async toggleMaximizeWindow(traceId?: string): Promise<{ ok: true }> {
    return getNativeBridge().window.toggleMaximize(traceId);
  },
  async closeWindow(traceId?: string): Promise<{ ok: true }> {
    return getNativeBridge().window.close(traceId);
  },
  async showFileTreeContextMenu(payload: FileTreeContextMenuPayload): Promise<{ ok: true }> {
    const contextApi = getNativeBridge().context;
    if (!contextApi?.showFileTreeMenu) {
      return {
        ok: true,
      };
    }

    return contextApi.showFileTreeMenu(payload);
  },
  async showSourceEditorContextMenu(payload: SourceEditorContextMenuPayload): Promise<{
    ok: true;
    handled: boolean;
  }> {
    const contextApi = getNativeBridge().context;
    if (!contextApi?.showSourceEditorMenu) {
      return {
        ok: true,
        handled: false,
      };
    }

    await contextApi.showSourceEditorMenu(payload);
    return {
      ok: true,
      handled: true,
    };
  },
  async getSpellCheckState(traceId?: string): Promise<SpellCheckStateResult> {
    const result = await getNativeBridge().spell.getState(traceId);
    return ensureSpellCheckStateResult(result, 'spell:getState');
  },
  async setSpellCheckEnabled(
    enabled: boolean,
    traceId?: string,
  ): Promise<SpellCheckStateResult> {
    const result = await getNativeBridge().spell.setEnabled(enabled, traceId);
    return ensureSpellCheckStateResult(result, 'spell:setEnabled');
  },
  async setSpellCheckLanguage(
    language: string,
    traceId?: string,
  ): Promise<SpellCheckStateResult> {
    const result = await getNativeBridge().spell.setLanguage(language, traceId);
    return ensureSpellCheckStateResult(result, 'spell:setLanguage');
  },
};
