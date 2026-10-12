export interface OpenFileResult {
  filePath: string;
  content: string;
  resourceBaseUrl?: string;
  version?: string;
}

export interface SaveFilePayload {
  sourceFilePath?: string;
  filePath?: string;
  content: string;
  traceId?: string;
  version?: string;
}

export interface SaveFileResult {
  content?: string;
  filePath: string;
  resourceBaseUrl?: string;
  version?: string;
}

export interface ListDirectoryMarkdownResult {
  directoryPath: string;
  filePaths: string[];
  modifiedTimes?: Record<string, number>;
  directories?: string[];
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
  | 'rtf'
  | 'epub'
  | 'latex'
  | 'revealjs'
  | 'mediawiki'
  | 'rst'
  | 'textile'
  | 'opml';

export interface PdfOptions {
  pageSize: 'A4' | 'A5' | 'Letter' | 'Legal' | 'Custom';
  widthMm?: number;
  heightMm?: number;
  marginsMm?: { top: number; bottom: number; left: number; right: number };
  marginMm: number;
  landscape: boolean;
  bookmarks: boolean;
  header: string;
  footer: string;
}

export interface DraftItem {
  id: string;
  filePath: string;
  savedAt: number;
  title: string;
}

export interface ImageInsertPayload {
  sourceFilePath: string;
  copyDirectory?: string;
}

export interface ImageInsertResult {
  images: Array<{ reference: string; name: string }>;
}

export interface ExportDocumentPayload {
  format: ExportFormat;
  html: string;
  markdown: string;
  sourceFilePath?: string;
  targetPath?: string;
  traceId?: string;
  pdfOptions?: PdfOptions;
  options?: ExportOptions;
}

export interface ExportOptions {
  tableOfContents?: boolean;
  tocDepth?: number;
  epubChapterLevel?: number;
  imageWidth: number;
  jpegQuality: number;
  includeStyles: boolean;
  headingPageBreak: boolean;
  theme: 'current' | 'light' | 'dark';
  title: string;
  author: string;
  language: string;
  referenceDoc: string;
  epubCover: string;
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
  autoSaveInterval?: number;
  userCssEnabled: boolean;
  userCssPath: string;
  spellCheckEnabled: boolean;
  spellCheckLanguage: string;
  writingPreferencesJson?: string;
  documentToolsVisible?: boolean;
  preferencesVisible?: boolean;
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
        openDroppedFile?: (file: File) => Promise<OpenFileResult | null>;
        openLinkedFile?: (payload: { sourceFilePath: string; href: string }) =>
          Promise<OpenFileResult | null>;
        operate?: (payload: {
          filePath: string; action: string; name?: string; directoryPath?: string;
        }) => Promise<{ filePath: string; previousPath?: string; action: string }>;
        chooseMoveDirectory?: () => Promise<{ directoryPath: string } | null>;
        inspect?: (payload: { filePath: string }) => Promise<{ version: string }>;
        openDirectory?: (payload?: { traceId?: string; directoryPath?: string }) =>
          Promise<ListDirectoryMarkdownResult | null>;
        searchDirectory?: (payload: { query: string; caseSensitive?: boolean;
          wholeWord?: boolean; regularExpression?: boolean }) => Promise<{
          truncated?: boolean;
          matches: Array<{ filePath: string; line: number; text: string }>;
        }>;
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
        chooseResource?: (kind: 'referenceDoc' | 'epubCover') =>
          Promise<{ filePath: string } | null>;
        importDocument?: () => Promise<OpenFileResult | null>;
        exportHtml: (payload: ExportHtmlPayload) => Promise<ExportHtmlResult | null>;
        exportDocument: (
          payload: ExportDocumentPayload
        ) => Promise<ExportDocumentResult | null>;
        getCapabilities: (traceId?: string) => Promise<ExportCapabilitiesResult>;
      };
      session?: {
        listDrafts?: () => Promise<{ drafts: DraftItem[] }>;
        restoreDraft?: (id: string) => Promise<{
          content: string; originalFilePath: string; filePath: string; isDirty: boolean;
        }>;
        load: () => Promise<unknown>;
        save: (snapshot: unknown) => Promise<{ ok: true }>;
        confirmDiscard: () => Promise<{ action: 'save' | 'discard' | 'cancel' }>;
      };
      window: {
        minimize: (traceId?: string) => Promise<{ ok: true }>;
        toggleMaximize: (traceId?: string) => Promise<{ ok: true }>;
        isMaximized: (traceId?: string) => Promise<WindowStateResult>;
        close: (traceId?: string) => Promise<{ ok: true }>;
        toggleDevTools: (traceId?: string) => Promise<{ ok: true }>;
      };
      images?: {
        collect?: (payload: { sourceFilePath: string; content: string }) =>
          Promise<{ content: string; count: number } | null>;
        applyRoot?: (payload: { sourceFilePath: string; frontMatter: string }) =>
          Promise<{ ok: boolean }>;
        select: (payload: ImageInsertPayload) => Promise<ImageInsertResult>;
        paste: (payload: ImageInsertPayload) => Promise<ImageInsertResult>;
        setRoot: (payload: ImageInsertPayload) => Promise<{ directoryPath: string } | null>;
        drop?: (file: File, payload: ImageInsertPayload) => Promise<ImageInsertResult>;
      };
      settings: {
        getUserCss?: () => Promise<{ css: string }>;
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
