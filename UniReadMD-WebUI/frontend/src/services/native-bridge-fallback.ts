import type {
  AppLogPayload,
  ExportCapabilitiesResult,
  ExportDocumentPayload,
  ExportDocumentResult,
  ExportHtmlPayload,
  ExportHtmlResult,
  ListDirectoryMarkdownResult,
  OpenFileResult,
  SaveFilePayload,
  SaveFileResult,
  SpellCheckStateResult,
  SettingsGetAllResult,
  SettingsKey,
  SettingsSetResult,
  SettingsState,
  SelectUserCssFileResult,
  WindowStateResult,
} from '../types/native';

const DEFAULT_SETTINGS: SettingsState = {
  theme: 'light',
  fontSize: 14,
  contentZoom: 100,
  viewMode: 'split',
  autoSave: false,
  userCssEnabled: false,
  userCssPath: '',
  spellCheckEnabled: true,
  spellCheckLanguage: 'en-US',
};
const FALLBACK_SPELL_LANGUAGES = ['en-US', 'en-GB', 'zh-CN'];
const FALLBACK_EXPORT_CAPABILITIES: ExportCapabilitiesResult = {
  builtinFormats: [
    'html',
    'pdf',
    'png',
    'jpeg',
  ],
  pandocAvailable: false,
  pandocVersion: '',
  pandocFormats: [],
};

function deriveFallbackExportPath(
  sourceFilePath: string | undefined,
  format: string,
) {
  const normalizedPath = typeof sourceFilePath === 'string'
    ? sourceFilePath.replace(/\\/g, '/')
    : '';
  const sourceName = normalizedPath.split('/').pop() || 'document.md';
  const baseName = sourceName.replace(/\.[^.]+$/g, '') || 'document';
  return `/virtual/${baseName}.${format}`;
}

const VIRTUAL_STORE = new Map<string, string>();
let virtualFileId = 0;
let settingsState: SettingsState = {
  ...DEFAULT_SETTINGS,
};
let recentFilesState: string[] = [];

function nextVirtualPath() {
  virtualFileId += 1;
  return `/virtual/document-${virtualFileId}.md`;
}

function saveVirtualMarkdown(payload: SaveFilePayload): SaveFileResult {
  const target = payload.filePath || nextVirtualPath();
  VIRTUAL_STORE.set(target, payload.content);

  return {
    filePath: target,
  };
}

function openVirtualMarkdown(filePath: string): OpenFileResult | null {
  if (!VIRTUAL_STORE.has(filePath)) {
    return null;
  }

  return {
    filePath,
    content: VIRTUAL_STORE.get(filePath) || '',
  };
}

function normalizeVirtualPath(inputPath: string): string {
  return inputPath.replace(/\\/g, '/');
}

function getVirtualDirectoryPath(filePath: string): string {
  const normalized = normalizeVirtualPath(filePath);
  const index = normalized.lastIndexOf('/');
  if (index <= 0) {
    return '/';
  }
  return normalized.slice(0, index);
}

function isMarkdownVirtualPath(filePath: string): boolean {
  const normalized = normalizeVirtualPath(filePath).toLowerCase();
  return normalized.endsWith('.md') || normalized.endsWith('.markdown');
}

function normalizeRecentFiles(filePaths: string[]): string[] {
  const seen = new Set<string>();
  const output: string[] = [];

  filePaths.forEach((item) => {
    if (typeof item !== 'string') {
      return;
    }

    const normalizedPath = item.trim();
    if (!normalizedPath || !isMarkdownVirtualPath(normalizedPath)) {
      return;
    }

    const dedupeKey = normalizeVirtualPath(normalizedPath).toLowerCase();
    if (seen.has(dedupeKey)) {
      return;
    }

    seen.add(dedupeKey);
    output.push(normalizedPath);
  });

  return output.slice(0, 12);
}

function setSetting(
  key: SettingsKey,
  value: SettingsState[SettingsKey],
): SettingsSetResult {
  settingsState = {
    ...settingsState,
    [key]: value,
  };

  return {
    key,
    value: settingsState[key],
    settings: settingsState,
  };
}

function buildSpellCheckState(): SpellCheckStateResult {
  const language = settingsState.spellCheckLanguage || FALLBACK_SPELL_LANGUAGES[0];
  return {
    enabled: settingsState.spellCheckEnabled,
    language,
    availableLanguages: [...FALLBACK_SPELL_LANGUAGES],
    dictionaryStatus: FALLBACK_SPELL_LANGUAGES.includes(language) ? 'ready' : 'missing',
  };
}

function createFallbackBridge() {
  return {
    __fallbackBridge: true,
    app: {
      async getVersion(): Promise<{ version: string }> {
        return {
          version: 'web-fallback',
        };
      },
      async logEvent(payload: AppLogPayload): Promise<{ ok: true }> {
        const prefix = payload.level === 'error' ? '[bridge:error]' : '[bridge:info]';
        console.log(prefix, payload.message, payload.traceId || '');
        return {
          ok: true,
        };
      },
      async consumeLaunchFilePath(): Promise<null> {
        return null;
      },
      onLaunchFile() {
        return () => {};
      },
    },
    file: {
      async openMarkdown(): Promise<OpenFileResult | null> {
        return null;
      },
      async openFromPath(payload: {
        filePath: string;
      }): Promise<OpenFileResult | null> {
        return openVirtualMarkdown(payload.filePath);
      },
      getPathForFile(): string {
        return '';
      },
      async saveMarkdown(payload: SaveFilePayload): Promise<SaveFileResult | null> {
        return saveVirtualMarkdown(payload);
      },
      async listMarkdownInDirectory(payload: {
        filePath: string;
      }): Promise<ListDirectoryMarkdownResult> {
        const directoryPath = getVirtualDirectoryPath(payload.filePath);
        const filePaths = Array.from(VIRTUAL_STORE.keys())
          .filter((item) => {
            if (!isMarkdownVirtualPath(item)) {
              return false;
            }

            return getVirtualDirectoryPath(item) === directoryPath;
          })
          .sort((left, right) => left.localeCompare(right, 'zh-CN'));

        return {
          directoryPath,
          filePaths,
        };
      },
    },
    exporter: {
      async exportHtml(payload: ExportHtmlPayload): Promise<ExportHtmlResult | null> {
        const target = payload.targetPath || '/virtual/export.html';
        VIRTUAL_STORE.set(target, payload.html);

        return {
          filePath: target,
        };
      },
      async exportDocument(
        payload: ExportDocumentPayload,
      ): Promise<ExportDocumentResult | null> {
        const format = String(payload.format || '').toLowerCase();
        const isBuiltin = FALLBACK_EXPORT_CAPABILITIES.builtinFormats.includes(
          format as ExportCapabilitiesResult['builtinFormats'][number],
        );
        if (!isBuiltin) {
          throw new Error(
            '当前环境未检测到 Pandoc，无法导出该格式。请安装 Pandoc 或改用 HTML/PDF/Image。',
          );
        }

        const target = payload.targetPath
          || deriveFallbackExportPath(payload.sourceFilePath, format);
        VIRTUAL_STORE.set(target, payload.html || payload.markdown || '');

        return {
          filePath: target,
          format: format as ExportDocumentResult['format'],
          engine: 'builtin',
        };
      },
      async getCapabilities(): Promise<ExportCapabilitiesResult> {
        return {
          ...FALLBACK_EXPORT_CAPABILITIES,
        };
      },
    },
    window: {
      async minimize(): Promise<{ ok: true }> {
        return {
          ok: true,
        };
      },
      async toggleMaximize(): Promise<{ ok: true }> {
        return {
          ok: true,
        };
      },
      async isMaximized(): Promise<WindowStateResult> {
        return {
          isMaximized: false,
        };
      },
      async close(): Promise<{ ok: true }> {
        return {
          ok: true,
        };
      },
      async toggleDevTools(): Promise<{ ok: true }> {
        return {
          ok: true,
        };
      },
    },
    settings: {
      async getAll(): Promise<SettingsGetAllResult> {
        return {
          settings: settingsState,
        };
      },
      async getRecentFiles() {
        return {
          filePaths: [...recentFilesState],
        };
      },
      async get(key: SettingsKey) {
        return {
          key,
          value: settingsState[key],
        };
      },
      async set(
        key: SettingsKey,
        value: SettingsState[SettingsKey],
      ): Promise<SettingsSetResult> {
        return setSetting(key, value);
      },
      async setRecentFiles(filePaths: string[]) {
        recentFilesState = normalizeRecentFiles(filePaths);
        return {
          filePaths: [...recentFilesState],
        };
      },
      async selectUserCssFile(): Promise<SelectUserCssFileResult | null> {
        return null;
      },
    },
    context: {
      async showFileTreeMenu(): Promise<{ ok: true }> {
        return {
          ok: true,
        };
      },
    },
    spell: {
      async getState(): Promise<SpellCheckStateResult> {
        return buildSpellCheckState();
      },
      async setEnabled(enabled: boolean): Promise<SpellCheckStateResult> {
        settingsState = {
          ...settingsState,
          spellCheckEnabled: Boolean(enabled),
        };
        return buildSpellCheckState();
      },
      async setLanguage(language: string): Promise<SpellCheckStateResult> {
        settingsState = {
          ...settingsState,
          spellCheckLanguage: language,
        };
        return buildSpellCheckState();
      },
    },
    menu: {
      onCommand() {
        return () => {};
      },
    },
    async getAppVersion() {
      return 'web-fallback';
    },
    async openMarkdownFile(): Promise<OpenFileResult | null> {
      return null;
    },
    async saveMarkdownFile(payload: SaveFilePayload): Promise<SaveFileResult | null> {
      return saveVirtualMarkdown(payload);
    },
  };
}

export function installNativeBridgeFallback() {
  if (window.nativeBridge) {
    return;
  }

  const isElectronRuntime = typeof navigator !== 'undefined'
    && /Electron/i.test(navigator.userAgent || '');
  if (isElectronRuntime) {
    // Electron 环境必须优先使用 preload 注入的真实桥接，避免降级桥接屏蔽原生能力。
    console.error('[bridge:fallback] nativeBridge missing in Electron runtime');
    return;
  }

  Object.defineProperty(window, 'nativeBridge', {
    configurable: true,
    writable: true,
    value: createFallbackBridge(),
  });
}
