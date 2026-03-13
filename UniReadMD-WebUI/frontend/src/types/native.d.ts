export interface OpenFileResult {
  filePath: string;
  content: string;
}

export interface SaveFilePayload {
  filePath?: string;
  content: string;
  traceId?: string;
}

export interface SaveFileResult {
  filePath: string;
}

export interface ListDirectoryMarkdownResult {
  directoryPath: string;
  filePaths: string[];
}

export interface ExportHtmlPayload {
  html: string;
  sourceFilePath?: string;
  targetPath?: string;
  traceId?: string;
}

export interface ExportHtmlResult {
  filePath: string;
}

export type ExportFormat =
  | 'html'
  | 'pdf'
  | 'png'
  | 'jpeg'
  | 'docx'
  | 'odt'
  | 'rtf';

export interface ExportDocumentPayload {
  format: ExportFormat;
  html: string;
  markdown: string;
  sourceFilePath?: string;
  targetPath?: string;
  traceId?: string;
}

export interface ExportDocumentResult {
  filePath: string;
  format: ExportFormat;
  engine: 'builtin' | 'pandoc';
}

export interface ExportCapabilitiesResult {
  builtinFormats: ExportFormat[];
  pandocAvailable: boolean;
  pandocVersion: string;
  pandocFormats: ExportFormat[];
}

export interface AppVersionResult {
  version: string;
}

export interface LaunchFilePathResult {
  filePath: string;
}

export interface LaunchFileEventPayload {
  filePath: string;
}

export interface AppLogPayload {
  traceId?: string;
  level?: 'info' | 'error';
  message: string;
  detail?: unknown;
}

export interface SettingsState {
  theme: 'light' | 'dark' | 'system';
  fontSize: number;
  contentZoom: number;
  viewMode: 'edit' | 'preview' | 'split';
  autoSave: boolean;
  userCssEnabled: boolean;
  userCssPath: string;
  spellCheckEnabled: boolean;
  spellCheckLanguage: string;
}

export type SettingsKey = keyof SettingsState;

export interface SettingsGetAllResult {
  settings: SettingsState;
}

export interface RecentFilesResult {
  filePaths: string[];
}

export interface SettingsGetResult {
  key: SettingsKey;
  value: SettingsState[SettingsKey];
}

export interface SettingsSetResult {
  key: SettingsKey;
  value: SettingsState[SettingsKey];
  settings: SettingsState;
}

export interface WindowStateResult {
  isMaximized: boolean;
}

export interface SelectUserCssFileResult {
  filePath: string;
}

export interface SpellCheckStateResult {
  enabled: boolean;
  language: string;
  availableLanguages: string[];
  dictionaryStatus: 'ready' | 'missing';
}

export interface FileTreeContextMenuPayload {
  filePath: string;
  x?: number;
  y?: number;
  traceId?: string;
}

export interface SourceEditorContextMenuPayload {
  x?: number;
  y?: number;
  traceId?: string;
}

declare global {
  interface Window {
    nativeBridge: {
      app: {
        getVersion: (traceId?: string) => Promise<AppVersionResult>;
        logEvent: (payload: AppLogPayload) => Promise<{ ok: true }>;
        consumeLaunchFilePath: (
          traceId?: string
        ) => Promise<LaunchFilePathResult | null>;
        onLaunchFile: (
          listener: (payload: LaunchFileEventPayload | null) => void
        ) => () => void;
      };
      file: {
        openMarkdown: (payload?: { traceId?: string }) => Promise<OpenFileResult | null>;
        openFromPath: (payload: {
          filePath: string;
          traceId?: string;
        }) => Promise<OpenFileResult | null>;
        getPathForFile: (file: File) => string;
        listMarkdownInDirectory: (payload: {
          filePath: string;
          traceId?: string;
        }) => Promise<ListDirectoryMarkdownResult>;
        saveMarkdown: (payload: SaveFilePayload) => Promise<SaveFileResult | null>;
      };
      exporter: {
        exportHtml: (payload: ExportHtmlPayload) => Promise<ExportHtmlResult | null>;
        exportDocument: (
          payload: ExportDocumentPayload
        ) => Promise<ExportDocumentResult | null>;
        getCapabilities: (traceId?: string) => Promise<ExportCapabilitiesResult>;
      };
      window: {
        minimize: (traceId?: string) => Promise<{ ok: true }>;
        toggleMaximize: (traceId?: string) => Promise<{ ok: true }>;
        isMaximized: (traceId?: string) => Promise<WindowStateResult>;
        close: (traceId?: string) => Promise<{ ok: true }>;
        toggleDevTools: (traceId?: string) => Promise<{ ok: true }>;
      };
      settings: {
        getAll: (traceId?: string) => Promise<SettingsGetAllResult>;
        getRecentFiles: (traceId?: string) => Promise<RecentFilesResult>;
        get: (key: SettingsKey, traceId?: string) => Promise<SettingsGetResult>;
        set: (
          key: SettingsKey,
          value: SettingsState[SettingsKey],
          traceId?: string
        ) => Promise<SettingsSetResult>;
        setRecentFiles: (
          filePaths: string[],
          traceId?: string
        ) => Promise<RecentFilesResult>;
        selectUserCssFile: (traceId?: string) => Promise<SelectUserCssFileResult | null>;
      };
      spell: {
        getState: (traceId?: string) => Promise<SpellCheckStateResult>;
        setEnabled: (
          enabled: boolean,
          traceId?: string
        ) => Promise<SpellCheckStateResult>;
        setLanguage: (
          language: string,
          traceId?: string
        ) => Promise<SpellCheckStateResult>;
      };
      context?: {
        showFileTreeMenu: (
          payload: FileTreeContextMenuPayload
        ) => Promise<{ ok: true }>;
        showSourceEditorMenu?: (
          payload: SourceEditorContextMenuPayload
        ) => Promise<{ ok: true }>;
      };
      menu?: {
        onCommand: (
          listener: (command: string) => void
        ) => () => void;
      };

      // backward compatibility
      getAppVersion: () => Promise<string>;
      openMarkdownFile: (traceId?: string) => Promise<OpenFileResult | null>;
      saveMarkdownFile: (payload: SaveFilePayload) => Promise<SaveFileResult | null>;
    };
  }
}

export {};
