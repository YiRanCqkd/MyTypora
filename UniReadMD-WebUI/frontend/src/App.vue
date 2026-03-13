<script setup lang='ts'>
import { HighlightStyle, syntaxHighlighting } from '@codemirror/language';
import { markdown as markdownLanguage } from '@codemirror/lang-markdown';
import { Compartment, EditorState, Transaction } from '@codemirror/state';
import {
  EditorView,
  highlightActiveLine,
  highlightActiveLineGutter,
  keymap,
} from '@codemirror/view';
import {
  defaultKeymap,
  history,
  historyKeymap,
  indentLess,
  indentMore,
  indentWithTab,
} from '@codemirror/commands';
import { tags } from '@lezer/highlight';
import MarkdownIt from 'markdown-it';
import type { PluginWithParams } from 'markdown-it';
import * as markdownItEmoji from 'markdown-it-emoji';
import markdownItFootnote from 'markdown-it-footnote';
import markdownItFrontMatter from 'markdown-it-front-matter';
import markdownItMark from 'markdown-it-mark';
import markdownItSub from 'markdown-it-sub';
import markdownItSup from 'markdown-it-sup';
import markdownItTaskLists from 'markdown-it-task-lists';
import markdownItTocDoneRight from 'markdown-it-toc-done-right';
import 'katex/dist/katex.min.css';
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import { useEditorStore } from './stores/editor';
import { bridgeService } from './services/bridge';
import type { ExportCapabilitiesResult, ExportFormat } from './types/native';
import {
  extractHeadings,
  findActiveHeadingLine,
  hasMathSyntax,
} from './utils/markdown';
import {
  applyPreviewTableActionToModel,
  parseMarkdownTable,
  serializeMarkdownTable,
  type ParsedMarkdownTable,
  type PreviewTableAction,
} from './utils/preview-table';
import {
  buildInlineShortcutPatch,
  indentMarkdownBlockLines,
  type InlineShortcutKind,
} from './utils/markdown-shortcuts';
import { createTraceId, logBridge } from './utils/trace';

type Tab = 'files' | 'outline';
type Theme = 'light' | 'dark' | 'system';
type FileViewMode = 'list' | 'tree';
type FileSortMode = 'natural' | 'name' | 'date';
type SpellDictionaryStatus = 'ready' | 'missing';
const BUILTIN_EXPORT_FORMATS: ExportFormat[] = ['html', 'pdf', 'png', 'jpeg'];
const EXPORT_LABELS: Record<ExportFormat, string> = {
  html: 'HTML',
  pdf: 'PDF',
  png: 'PNG',
  jpeg: 'JPEG',
  docx: 'Word(docx)',
  odt: 'ODT',
  rtf: 'RTF',
};
const SESSION_STORAGE_KEY = 'unireadmd.phase2.session.v1';
const RECENT_FILES_LIMIT = 12;
const CONTENT_ZOOM_DEFAULT = 100;
const CONTENT_ZOOM_MIN = 50;
const CONTENT_ZOOM_MAX = 200;
const CONTENT_ZOOM_STEP = 10;
const SIDEBAR_WIDTH_DEFAULT = 300;
const SIDEBAR_WIDTH_MIN = 220;
const SIDEBAR_WIDTH_MAX = 520;
const MAIN_PREVIEW_WIDTH_DEFAULT_RATIO = 0.58;
const MAIN_PREVIEW_WIDTH_MIN = 320;
const MAIN_SOURCE_WIDTH_MIN = 320;
const PREVIEW_INLINE_EDIT_SYNC_DELAY = 120;

interface SessionSnapshot {
  filePath: string;
  content: string;
  isDirty: boolean;
  cursorHint: number;
  recentFiles: string[];
  contentZoom?: number;
  pinOutline?: boolean;
  sourceMode?: boolean;
  focusMode?: boolean;
  typewriterMode?: boolean;
  sidebarWidth?: number;
  mainPreviewWidth?: number;
  fileViewMode?: FileViewMode;
  fileSortMode?: FileSortMode;
  filesFilter?: string;
  searchPanelVisible?: boolean;
  replacePanelVisible?: boolean;
  searchQuery?: string;
  replaceQuery?: string;
  searchCaseSensitive?: boolean;
  searchWholeWord?: boolean;
  searchMatchIndex?: number;
  fileTreeExpanded?: Record<string, boolean>;
}

interface FileListItem {
  path: string;
  name: string;
  parent: string;
  ext: string;
  order: number;
}

interface FileTreeRow {
  id: string;
  type: 'dir' | 'file';
  depth: number;
  label: string;
  path: string;
  expandable: boolean;
  expanded: boolean;
  filePath?: string;
}

interface DocMatchRange {
  from: number;
  to: number;
  text: string;
}

interface PreviewInlineEditSelection {
  startLine: number;
  endLine: number;
  token: string;
}

interface PreviewInlineEditSession {
  from: number;
  to: number;
  startLine: number;
  originalText: string;
  appliedText: string;
}

interface PreviewTableSelection {
  startLine: number;
  endLine: number;
  rowIndex: number;
  colIndex: number;
}

interface PreviewTableCellContext extends PreviewTableSelection {
  table: HTMLElement;
  cell: HTMLElement;
}

interface PreviewTaskItemContext {
  startLine: number;
  endLine: number;
  taskIndex: number;
  globalTaskIndex: number;
}

interface MarkdownSourceToken {
  type: string;
  map?: [number, number];
  info?: string;
  content?: string;
  attrSet?: (name: string, value: string) => void;
  attrGet?: (name: string) => string | null;
}

const SOURCE_MAPPED_BLOCK_TOKENS = [
  'paragraph_open',
  'heading_open',
  'blockquote_open',
  'list_item_open',
  'table_open',
  'code_block',
];

const PREVIEW_INLINE_EDITABLE_TOKENS = new Set([
  'paragraph_open',
  'heading_open',
  'blockquote_open',
  'list_item_open',
  'fence',
  'code_block',
]);

const PREVIEW_TOKEN_LABELS: Record<string, string> = {
  paragraph_open: '段落',
  heading_open: '标题',
  blockquote_open: '引用块',
  list_item_open: '列表项',
  table_open: '表格',
  fence: '代码块',
  code_block: '代码块',
};

const sourceHighlightStyle = HighlightStyle.define([
  {
    tag: [
      tags.heading,
      tags.propertyName,
      tags.labelName,
    ],
    color: '#cebcca',
    fontWeight: '400',
    textDecoration: 'none',
  },
  {
    tag: tags.comment,
    color: '#9fb1ff',
  },
  {
    tag: [
      tags.string,
      tags.special(tags.string),
    ],
    color: '#a7a7d9',
  },
  {
    tag: [
      tags.number,
      tags.atom,
      tags.bool,
      tags.null,
    ],
    color: '#848695',
    fontStyle: 'italic',
  },
  {
    tag: [
      tags.link,
      tags.url,
    ],
    color: '#95b94b',
    textDecoration: 'none',
  },
]);

const markdown = new MarkdownIt({
  html: true,
  linkify: true,
  breaks: true,
});

function resolveMarkdownPlugin(pluginLike: unknown, exportKey?: string) {
  if (typeof pluginLike === 'function') {
    return pluginLike as PluginWithParams;
  }

  if (!pluginLike || typeof pluginLike !== 'object') {
    return null;
  }

  const record = pluginLike as Record<string, unknown>;
  if (exportKey && typeof record[exportKey] === 'function') {
    return record[exportKey] as PluginWithParams;
  }

  if (typeof record.default === 'function') {
    return record.default as PluginWithParams;
  }

  for (const candidateKey of ['full', 'light', 'bare']) {
    if (typeof record[candidateKey] === 'function') {
      return record[candidateKey] as PluginWithParams;
    }
  }

  return null;
}

function useMarkdownPlugin(pluginLike: unknown, ...params: unknown[]) {
  const plugin = resolveMarkdownPlugin(pluginLike);
  if (!plugin) {
    console.warn('[markdown] 插件加载失败：未找到可用导出。');
    return;
  }

  markdown.use(plugin as PluginWithParams, ...params);
}

let latestFrontMatter = '';
useMarkdownPlugin(markdownItFrontMatter, (frontMatter: string) => {
  latestFrontMatter = frontMatter;
});
useMarkdownPlugin(markdownItFootnote);
useMarkdownPlugin(markdownItSub);
useMarkdownPlugin(markdownItSup);
useMarkdownPlugin(markdownItMark);
useMarkdownPlugin(markdownItTaskLists, {
  enabled: true,
  label: true,
  labelAfter: true,
});
useMarkdownPlugin(resolveMarkdownPlugin(markdownItEmoji, 'full'));
useMarkdownPlugin(markdownItTocDoneRight, {
  level: [1, 2, 3, 4, 5, 6],
  containerClass: 'md-toc',
});

function annotateTokenSourceRange(token: MarkdownSourceToken) {
  if (!token.map || typeof token.attrSet !== 'function') {
    return;
  }

  const startLine = token.map[0] + 1;
  const endLine = Math.max(startLine, token.map[1]);
  token.attrSet('data-source-start', String(startLine));
  token.attrSet('data-source-end', String(endLine));
  token.attrSet('data-source-token', token.type);
}

function installSourceMapRendererRule(ruleName: string) {
  const fallbackRule = markdown.renderer.rules[ruleName];
  markdown.renderer.rules[ruleName] = (tokens, idx, options, env, self) => {
    const token = tokens[idx] as MarkdownSourceToken;
    annotateTokenSourceRange(token);

    if (fallbackRule) {
      return fallbackRule(tokens, idx, options, env, self);
    }

    return self.renderToken(tokens, idx, options);
  };
}

SOURCE_MAPPED_BLOCK_TOKENS.forEach((ruleName) => {
  installSourceMapRendererRule(ruleName);
});

const originalFenceRule = markdown.renderer.rules.fence;
const diagramFenceLanguages = new Set(['mermaid', 'sequence', 'flow']);
markdown.renderer.rules.fence = (tokens, idx, options, env, self) => {
  const token = tokens[idx] as MarkdownSourceToken;
  annotateTokenSourceRange(token);

  const sourceStart = token.attrGet?.('data-source-start') || '';
  const sourceEnd = token.attrGet?.('data-source-end') || '';
  const sourceToken = token.attrGet?.('data-source-token') || 'fence';
  const sourceAttrs = sourceStart && sourceEnd
    ? ` data-source-start='${sourceStart}' data-source-end='${sourceEnd}' data-source-token='${sourceToken}'`
    : '';

  const info = (token.info || '').trim().split(/\s+/)[0] || '';

  if (!diagramFenceLanguages.has(info)) {
    if (originalFenceRule) {
      return originalFenceRule(tokens, idx, options, env, self);
    }

    return self.renderToken(tokens, idx, options);
  }

  const escaped = markdown.utils.escapeHtml(token.content || '');
  const encoded = encodeURIComponent(token.content || '');

  return `
<div class='mermaid-block'${sourceAttrs} data-diagram-kind='${info}' data-diagram-code='${encoded}'>
  <div class='mermaid-output'></div>
  <pre class='mermaid-fallback'>${escaped}</pre>
  <div class='mermaid-error'></div>
</div>`;
};

const originalLinkOpenRule = markdown.renderer.rules.link_open;
markdown.renderer.rules.link_open = (tokens, idx, options, env, self) => {
  const token = tokens[idx] as MarkdownSourceToken;
  token.attrSet?.('target', '_blank');
  token.attrSet?.('rel', 'noopener noreferrer');

  if (originalLinkOpenRule) {
    return originalLinkOpenRule(tokens, idx, options, env, self);
  }

  return self.renderToken(tokens, idx, options);
};

const store = useEditorStore();
const activeTab = ref<Tab>('outline');
const appVersion = ref('0.0.0');
const errorMessage = ref('');
const theme = ref<Theme>('light');
const fontSize = ref(14);
const contentZoom = ref(CONTENT_ZOOM_DEFAULT);
const userCssEnabled = ref(false);
const userCssPath = ref('');
const currentCursorLine = ref(1);
const recentFiles = ref<string[]>([]);
const editorPreviewHost = ref<HTMLElement | null>(null);
const editorHost = ref<HTMLDivElement | null>(null);
const previewHost = ref<HTMLElement | null>(null);
const isNativeWindow = ref(false);
const pinOutline = ref(true);
const sourceMode = ref(false);
const focusMode = ref(false);
const typewriterMode = ref(false);
const showSpellPanel = ref(false);
const spellCheckEnabled = ref(true);
const spellCheckLanguage = ref('en-US');
const spellCheckLanguages = ref<string[]>([]);
const spellDictionaryStatus = ref<SpellDictionaryStatus>('ready');
const exportCapabilities = ref<ExportCapabilitiesResult | null>(null);
const sidebarWidth = ref(SIDEBAR_WIDTH_DEFAULT);
const isResizingSidebar = ref(false);
const mainPreviewWidth = ref(0);
const isResizingMainPreview = ref(false);
const fileViewMode = ref<FileViewMode>('tree');
const fileSortMode = ref<FileSortMode>('natural');
const filesFilter = ref('');
const searchText = ref('');
const fileTreeExpanded = ref<Record<string, boolean>>({});
const currentDirectoryPath = ref('');
const currentDirectoryMarkdownFiles = ref<string[]>([]);
const showDocSearchPanel = ref(false);
const showDocReplacePanel = ref(false);
const docSearchInput = ref<HTMLInputElement | null>(null);
const docSearchQuery = ref('');
const docReplaceQuery = ref('');
const docSearchCaseSensitive = ref(false);
const docSearchWholeWord = ref(false);
const docSearchMatchIndex = ref(0);
const isDocumentLoading = ref(false);
const isDragOverFile = ref(false);
const previewInlineEditVisible = ref(false);
const previewInlineEditValue = ref('');
const previewInlineEditStyle = ref<Record<string, string>>({});
const previewInlineEditSelection = ref<PreviewInlineEditSelection | null>(null);
const previewInlineEditSession = ref<PreviewInlineEditSession | null>(null);
const previewInlineEditComposing = ref(false);
const previewInlineEditTextarea = ref<HTMLTextAreaElement | null>(null);
const previewTableMenuVisible = ref(false);
const previewTableMenuStyle = ref<Record<string, string>>({});
const previewTableSelection = ref<PreviewTableSelection | null>(null);
const previewTableCellEditVisible = ref(false);
const previewTableCellEditValue = ref('');
const previewTableCellEditStyle = ref<Record<string, string>>({});
const previewTableCellEditSelection = ref<PreviewTableSelection | null>(null);
const previewTableCellEditInput = ref<HTMLInputElement | null>(null);
const sourceContextMenuVisible = ref(false);
const sourceContextMenuStyle = ref<Record<string, string>>({});

const fontSizeCompartment = new Compartment();
let editorView: EditorView | null = null;
let isSyncingFromStore = false;
let mermaidRenderSeed = 0;
let mermaidThemeKey = '';
let mermaidLoadPromise: Promise<unknown> | null = null;
let mermaidApi: {
  initialize: (config: Record<string, unknown>) => void;
  render: (id: string, source: string) => Promise<{ svg: string }>;
} | null = null;
let flowchartLoadPromise: Promise<unknown> | null = null;
let flowchartApi: {
  parse: (code: string) => {
    drawSVG: (container: HTMLElement | string, options?: Record<string, unknown>) => void;
  };
} | null = null;
let mathPluginLoadPromise: Promise<void> | null = null;
const mathPluginRevision = ref(0);
let sidebarResizeCleanup: (() => void) | null = null;
let mainPreviewResizeCleanup: (() => void) | null = null;
let menuCommandCleanup: (() => void) | null = null;
let launchFileCleanup: (() => void) | null = null;
let typewriterScrollFrame = 0;
let previewScrollSyncFrame = 0;
let previewInlineEditSyncTimer = 0;
let contentZoomPersistTimer = 0;
let previewHeadingNodes: Array<{
  line: number;
  node: HTMLElement;
}> = [];
let previewSourceNodes: Array<{
  line: number;
  node: HTMLElement;
}> = [];
let fileDragDepth = 0;

const renderedHtml = computed(() => {
  void latestFrontMatter;
  const normalizedContent = store.content.replace(/^\uFEFF/, '');
  const normalized = normalizedContent.replace(/^\[toc\]$/gim, '[[toc]]');
  return markdown.render(normalized);
});

const headings = computed(() => {
  return extractHeadings(store.content);
});

const activeHeadingLine = computed(() => {
  return findActiveHeadingLine(headings.value, currentCursorLine.value);
});

const sidebarSearchPlaceholder = computed(() => {
  if (activeTab.value === 'files') {
    return '匹配文档名（当前目录 / 最近打开）';
  }

  return '匹配当前文档标题';
});

const sidebarSearchText = computed({
  get() {
    if (activeTab.value === 'files') {
      return filesFilter.value;
    }

    return searchText.value;
  },
  set(value: string) {
    if (activeTab.value === 'files') {
      filesFilter.value = value;
      return;
    }

    searchText.value = value;
  },
});

const outlineDisplayItems = computed(() => {
  const items = headings.value;
  const keyword = searchText.value.trim().toLowerCase();
  if (!keyword) {
    return items;
  }

  const includedLines = new Set<number>();
  items.forEach((item, index) => {
    if (!item.text.toLowerCase().includes(keyword)) {
      return;
    }

    includedLines.add(item.line);
    let parentLevel = item.level;
    for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
      const candidate = items[cursor];
      if (candidate.level < parentLevel) {
        includedLines.add(candidate.line);
        parentLevel = candidate.level;
      }

      if (parentLevel <= 1) {
        break;
      }
    }
  });

  return items.filter((item) => includedLines.has(item.line));
});

const currentFileName = computed(() => {
  if (!store.filePath) {
    return 'Untitled.md';
  }

  const normalized = store.filePath.replace(/\\/g, '/');
  const parts = normalized.split('/');
  return parts[parts.length - 1] || store.filePath;
});

const windowTitle = computed(() => {
  return `${currentFileName.value} - UniReadMD`;
});

const contentCharCount = computed(() => {
  return store.content.replace(/\r?\n/g, '').length;
});

const contentWordCount = computed(() => {
  const tokens = store.content
    .replace(/\r?\n/g, ' ')
    .split(/\s+/)
    .map((item) => item.trim())
    .filter((item) => item.length > 0);
  return tokens.length;
});

const contentReadMinutes = computed(() => {
  return Math.max(1, Math.ceil(contentWordCount.value / 220));
});

const statusWordLabel = computed(() => {
  return `${contentWordCount.value} Words`;
});

const contentZoomLabel = computed(() => {
  return `${contentZoom.value}%`;
});

const statusLineLabel = computed(() => {
  if (store.cursorHint > 0) {
    return `L${store.cursorHint}`;
  }

  return `L${currentCursorLine.value}`;
});

const sourceToggleLabel = computed(() => {
  if (sourceMode.value) {
    return '退出源码模式';
  }

  return '源码模式';
});

const spellDictionaryLabel = computed(() => {
  return spellDictionaryStatus.value === 'ready' ? 'Dictionary Ready' : 'Dictionary Missing';
});

const previewInlineEditTitle = computed(() => {
  const selection = previewInlineEditSelection.value;
  if (!selection) {
    return '渲染态编辑';
  }

  return PREVIEW_TOKEN_LABELS[selection.token] || '渲染块';
});

const fileListItems = computed<FileListItem[]>(() => {
  const unique = new Map<string, FileListItem>();
  const paths = store.filePath
    ? [store.filePath, ...recentFiles.value]
    : [...recentFiles.value];

  paths.forEach((path, index) => {
    if (!path || unique.has(path) || !isMarkdownPath(path)) {
      return;
    }

    unique.set(path, {
      path,
      name: getFileName(path),
      parent: getParentPath(path),
      ext: getFileExtension(path),
      order: index,
    });
  });

  let items = Array.from(unique.values());
  const sortMode = fileSortMode.value;
  if (sortMode === 'name') {
    items = [...items].sort((left, right) => {
      return left.name.localeCompare(right.name, 'zh-CN');
    });
  } else if (sortMode === 'date') {
    items = [...items].sort((left, right) => {
      return left.order - right.order;
    });
  } else {
    items = [...items].sort((left, right) => left.order - right.order);
  }

  return items;
});

const currentDirectoryLabel = computed(() => {
  const fallbackPath = store.filePath ? getParentPath(store.filePath) : '';
  const directoryPath = currentDirectoryPath.value || fallbackPath;
  if (!directoryPath) {
    return '当前目录';
  }

  return getDirectoryDisplayName(directoryPath);
});

const fileTreeRows = computed<FileTreeRow[]>(() => {
  const output: FileTreeRow[] = [];
  const currentId = 'dir:group-current';
  const recentId = 'dir:group-recent';
  const fileKeyword = filesFilter.value.trim().toLowerCase();
  const filterByName = (paths: string[]) => {
    if (!fileKeyword) {
      return paths;
    }

    return paths.filter((item) => getFileName(item).toLowerCase().includes(fileKeyword));
  };
  const currentFiles = filterByName(dedupeMarkdownPaths(currentDirectoryMarkdownFiles.value));
  const recentFilesInSidebar = filterByName(fileListItems.value.map((item) => item.path));

  const pushGroupRows = (groupId: string, groupLabel: string, paths: string[]) => {
    if (paths.length === 0) {
      return;
    }

    const expanded = fileTreeExpanded.value[groupId] !== false;
    output.push({
      id: groupId,
      type: 'dir',
      depth: 0,
      label: groupLabel,
      path: groupId,
      expandable: true,
      expanded,
    });

    if (!expanded) {
      return;
    }

    paths.forEach((filePath) => {
      output.push({
        id: `file:${filePath}`,
        type: 'file',
        depth: 1,
        label: getFileName(filePath),
        path: filePath,
        expandable: false,
        expanded: false,
        filePath,
      });
    });
  };

  pushGroupRows(currentId, currentDirectoryLabel.value, currentFiles);
  pushGroupRows(recentId, '最近打开', recentFilesInSidebar);
  return output;
});

const docSearchMatches = computed<DocMatchRange[]>(() => {
  const expression = buildSearchRegExp(
    docSearchQuery.value,
    docSearchCaseSensitive.value,
    docSearchWholeWord.value,
    true,
  );

  if (!expression) {
    return [];
  }

  const matches: DocMatchRange[] = [];
  const content = store.content;
  let current: RegExpExecArray | null = null;
  while ((current = expression.exec(content)) !== null) {
    matches.push({
      from: current.index,
      to: current.index + current[0].length,
      text: current[0],
    });

    if (matches.length >= 400) {
      break;
    }

    if (current[0].length === 0) {
      expression.lastIndex += 1;
    }
  }

  return matches;
});

const docSearchSummary = computed(() => {
  const total = docSearchMatches.value.length;
  if (total === 0) {
    return 'No Results';
  }

  return `${docSearchMatchIndex.value + 1}/${total}`;
});

const layoutClassName = computed(() => {
  return {
    'native-window': isNativeWindow.value,
    'pin-outline': pinOutline.value,
    'hide-sidebar': !pinOutline.value,
    'is-resizing-sidebar': isResizingSidebar.value,
    'is-resizing-main-preview': isResizingMainPreview.value,
    'typora-sourceview-on': sourceMode.value,
    'on-focus-mode': focusMode.value,
    'ty-on-typewriter-mode': typewriterMode.value,
    'on-search-panel-open': showDocSearchPanel.value,
    'on-replace-panel-open': showDocSearchPanel.value && showDocReplacePanel.value,
  };
});

const layoutStyle = computed(() => {
  return {
    '--sidebar-width': `${sidebarWidth.value}px`,
    '--main-preview-width': mainPreviewWidth.value > 0
      ? `${mainPreviewWidth.value}px`
      : '58%',
    '--content-zoom': `${contentZoom.value / 100}`,
  };
});

const availableExportFormats = computed(() => {
  if (!exportCapabilities.value) {
    return new Set<ExportFormat>(BUILTIN_EXPORT_FORMATS);
  }

  const formats = [
    ...exportCapabilities.value.builtinFormats,
    ...exportCapabilities.value.pandocFormats,
  ];
  const uniqueFormats = formats.filter((item, index) => formats.indexOf(item) === index);
  return new Set<ExportFormat>(uniqueFormats);
});

function normalizeSidebarWidth(width: number) {
  if (!Number.isFinite(width)) {
    return SIDEBAR_WIDTH_DEFAULT;
  }

  const rounded = Math.round(width);
  return Math.min(SIDEBAR_WIDTH_MAX, Math.max(SIDEBAR_WIDTH_MIN, rounded));
}

function normalizeContentZoom(value: number) {
  if (!Number.isFinite(value)) {
    return CONTENT_ZOOM_DEFAULT;
  }

  const rounded = Math.round(value / CONTENT_ZOOM_STEP) * CONTENT_ZOOM_STEP;
  return Math.min(CONTENT_ZOOM_MAX, Math.max(CONTENT_ZOOM_MIN, rounded));
}

function normalizeMainPreviewWidth(width: number, containerWidth: number) {
  if (!Number.isFinite(width) || !Number.isFinite(containerWidth) || containerWidth <= 0) {
    return 0;
  }

  const minWidth = Math.min(MAIN_PREVIEW_WIDTH_MIN, Math.floor(containerWidth / 2));
  const maxWidth = Math.max(
    minWidth,
    Math.floor(containerWidth - MAIN_SOURCE_WIDTH_MIN),
  );
  const rounded = Math.round(width);
  return Math.min(maxWidth, Math.max(minWidth, rounded));
}

function resolveMainPreviewContainerWidth() {
  return editorPreviewHost.value?.clientWidth ?? 0;
}

function ensureMainPreviewWidth() {
  if (!sourceMode.value) {
    return;
  }

  const containerWidth = resolveMainPreviewContainerWidth();
  if (containerWidth <= 0) {
    return;
  }

  if (mainPreviewWidth.value <= 0) {
    const initialWidth = containerWidth * MAIN_PREVIEW_WIDTH_DEFAULT_RATIO;
    mainPreviewWidth.value = normalizeMainPreviewWidth(initialWidth, containerWidth);
    return;
  }

  mainPreviewWidth.value = normalizeMainPreviewWidth(
    mainPreviewWidth.value,
    containerWidth,
  );
}

function normalizeSlashes(inputPath: string) {
  return inputPath.replace(/\\/g, '/');
}

function splitPath(inputPath: string) {
  return normalizeSlashes(inputPath).split('/').filter((item) => item.length > 0);
}

function getFileName(inputPath: string) {
  const parts = splitPath(inputPath);
  return parts[parts.length - 1] || inputPath;
}

function getParentPath(inputPath: string) {
  const parts = splitPath(inputPath);
  if (parts.length <= 1) {
    return '/';
  }

  return parts.slice(0, -1).join('/');
}

function getDirectoryDisplayName(inputPath: string) {
  const normalized = normalizeSlashes(inputPath).trim();
  if (!normalized) {
    return '当前目录';
  }

  if (normalized === '/') {
    return '/';
  }

  const withoutTailSlash = normalized.replace(/\/+$/g, '');
  if (/^[A-Za-z]:$/.test(withoutTailSlash)) {
    return `${withoutTailSlash}/`;
  }

  if (!withoutTailSlash) {
    return '/';
  }

  const parts = withoutTailSlash.split('/').filter((item) => item.length > 0);
  if (parts.length === 0) {
    return '/';
  }

  return parts[parts.length - 1];
}

function getFileExtension(inputPath: string) {
  const name = getFileName(inputPath);
  const index = name.lastIndexOf('.');
  if (index < 0 || index === name.length - 1) {
    return '';
  }

  return name.slice(index + 1).toLowerCase();
}

function isMarkdownPath(inputPath: string) {
  const ext = getFileExtension(inputPath);
  return ext === 'md' || ext === 'markdown';
}

function dedupeMarkdownPaths(paths: string[]) {
  const seen = new Set<string>();
  const result: string[] = [];
  paths.forEach((item) => {
    if (!item || !isMarkdownPath(item)) {
      return;
    }

    const key = normalizeSlashes(item).toLowerCase();
    if (seen.has(key)) {
      return;
    }

    seen.add(key);
    result.push(item);
  });
  return result;
}

function normalizeRecentFiles(paths: string[]) {
  return dedupeMarkdownPaths(paths).slice(0, RECENT_FILES_LIMIT);
}

function getEffectiveEditorFontSize() {
  const scaledSize = (fontSize.value * contentZoom.value) / 100;
  return Number(scaledSize.toFixed(2));
}

function escapeRegExp(input: string) {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function buildSearchRegExp(
  query: string,
  isCaseSensitive: boolean,
  isWholeWord: boolean,
  isGlobal: boolean,
) {
  const normalizedQuery = query.trim();
  if (!normalizedQuery) {
    return null;
  }

  const escaped = escapeRegExp(normalizedQuery);
  const pattern = isWholeWord ? `\\b${escaped}\\b` : escaped;
  const flags = `${isGlobal ? 'g' : ''}${isCaseSensitive ? '' : 'i'}`;

  try {
    return new RegExp(pattern, flags);
  } catch {
    return null;
  }
}

function clearPreviewSearchHighlights() {
  const preview = previewHost.value;
  if (!preview) {
    return;
  }

  const marks = Array.from(
    preview.querySelectorAll<HTMLElement>("mark[data-doc-search-highlight='true']"),
  );
  marks.forEach((mark) => {
    const parent = mark.parentNode;
    if (!parent) {
      return;
    }

    while (mark.firstChild) {
      parent.insertBefore(mark.firstChild, mark);
    }
    parent.removeChild(mark);
  });
  preview.normalize();
}

function shouldSkipPreviewSearchHighlight(node: Text) {
  const parent = node.parentElement;
  const value = node.nodeValue || '';
  if (!parent || !value.trim()) {
    return true;
  }

  if (parent.closest("mark[data-doc-search-highlight='true']")) {
    return true;
  }

  const tagName = parent.tagName;
  if (['SCRIPT', 'STYLE', 'TEXTAREA', 'INPUT', 'OPTION'].includes(tagName)) {
    return true;
  }

  return Boolean(
    parent.closest(
      '.mermaid-block, .preview-inline-editor, .preview-table-cell-editor, .preview-table-menu',
    ),
  );
}

function applyPreviewSearchHighlights() {
  const preview = previewHost.value;
  clearPreviewSearchHighlights();
  if (!preview || !showDocSearchPanel.value) {
    return;
  }

  const expression = buildSearchRegExp(
    docSearchQuery.value,
    docSearchCaseSensitive.value,
    docSearchWholeWord.value,
    true,
  );
  if (!expression) {
    return;
  }

  const textNodes: Text[] = [];
  const walker = document.createTreeWalker(preview, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      return shouldSkipPreviewSearchHighlight(node as Text)
        ? NodeFilter.FILTER_REJECT
        : NodeFilter.FILTER_ACCEPT;
    },
  });

  let currentNode = walker.nextNode();
  while (currentNode) {
    textNodes.push(currentNode as Text);
    currentNode = walker.nextNode();
  }

  const marks: HTMLElement[] = [];
  textNodes.forEach((textNode) => {
    const value = textNode.nodeValue || '';
    const localExpression = new RegExp(expression.source, expression.flags);
    let match = localExpression.exec(value);
    if (!match) {
      return;
    }

    const fragment = document.createDocumentFragment();
    let lastIndex = 0;
    while (match) {
      const [matchedText] = match;
      const matchIndex = match.index;
      if (matchIndex > lastIndex) {
        fragment.appendChild(document.createTextNode(value.slice(lastIndex, matchIndex)));
      }

      const mark = document.createElement('mark');
      mark.className = 'doc-search-highlight';
      mark.dataset.docSearchHighlight = 'true';
      mark.textContent = matchedText;
      fragment.appendChild(mark);
      marks.push(mark);
      lastIndex = matchIndex + matchedText.length;

      if (matchedText.length === 0) {
        localExpression.lastIndex += 1;
      }
      match = localExpression.exec(value);
    }

    if (lastIndex < value.length) {
      fragment.appendChild(document.createTextNode(value.slice(lastIndex)));
    }

    textNode.parentNode?.replaceChild(fragment, textNode);
  });

  if (marks.length === 0) {
    return;
  }

  const activeIndex = normalizeSearchIndex(docSearchMatchIndex.value, marks.length);
  const activeMark = marks[activeIndex];
  if (activeMark) {
    activeMark.classList.add('is-active');
  }
}

function applyTheme(mode: Theme) {
  document.documentElement.dataset.theme = mode;
}

function toFileHref(filePath: string) {
  const normalized = filePath.replace(/\\/g, '/');
  if (/^[A-Za-z]:\//.test(normalized)) {
    return `file:///${encodeURI(normalized)}`;
  }

  return `file://${encodeURI(normalized)}`;
}

function getUserCssLinkElement() {
  const linkId = 'unireadmd-user-css';
  let link = document.getElementById(linkId) as HTMLLinkElement | null;

  if (link) {
    return link;
  }

  link = document.createElement('link');
  link.id = linkId;
  link.rel = 'stylesheet';
  document.head.appendChild(link);

  return link;
}

function applyUserCss() {
  const link = getUserCssLinkElement();
  // 第二阶段已移除用户 CSS 功能，为避免遗留配置污染渲染样式，强制禁用外部 CSS 注入。
  link.href = '';
}

async function normalizeLegacyUserCssSettings() {
  if (!userCssEnabled.value && !userCssPath.value) {
    return;
  }

  await runTask('settings-usercss-disable', async (traceId) => {
    await bridgeService.setSetting('userCssEnabled', false, traceId);
    await bridgeService.setSetting('userCssPath', '', traceId);
    return {
      ok: true,
    };
  }, {
    fallbackHint: '若样式仍异常，请删除用户目录中的 settings.json 后重启。',
  });

  userCssEnabled.value = false;
  userCssPath.value = '';
}

function pushRecentFile(filePath: string) {
  if (!filePath) {
    return;
  }

  const filtered = recentFiles.value.filter((item) => item !== filePath);
  const nextRecentFiles = normalizeRecentFiles([filePath, ...filtered]);
  recentFiles.value = nextRecentFiles;

  if (isNativeWindow.value) {
    void bridgeService.setRecentFiles(nextRecentFiles, createTraceId('recent-files'))
      .then((result) => {
        recentFiles.value = normalizeRecentFiles(result.filePaths);
      })
      .catch(async (error) => {
        const message = error instanceof Error ? error.message : String(error);
        await logBridge('error', 'recent-files-sync:fail', createTraceId('recent-files'), {
          message,
          filePath,
        });
      });
    return;
  }

  persistSessionSnapshot();
}

function buildSessionSnapshot(): SessionSnapshot {
  return {
    filePath: store.filePath,
    content: store.content,
    isDirty: store.isDirty,
    cursorHint: store.cursorHint,
    recentFiles: recentFiles.value,
    contentZoom: contentZoom.value,
    pinOutline: pinOutline.value,
    sourceMode: sourceMode.value,
    focusMode: focusMode.value,
    typewriterMode: typewriterMode.value,
    sidebarWidth: sidebarWidth.value,
    mainPreviewWidth: mainPreviewWidth.value,
    fileViewMode: fileViewMode.value,
    fileSortMode: fileSortMode.value,
    filesFilter: filesFilter.value,
    searchPanelVisible: showDocSearchPanel.value,
    replacePanelVisible: showDocReplacePanel.value,
    searchQuery: docSearchQuery.value,
    replaceQuery: docReplaceQuery.value,
    searchCaseSensitive: docSearchCaseSensitive.value,
    searchWholeWord: docSearchWholeWord.value,
    searchMatchIndex: docSearchMatchIndex.value,
    fileTreeExpanded: fileTreeExpanded.value,
  };
}

function persistSessionSnapshot() {
  const snapshot = buildSessionSnapshot();
  localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(snapshot));
}

function loadSessionSnapshot() {
  const raw = localStorage.getItem(SESSION_STORAGE_KEY);
  if (!raw) {
    return null;
  }

  try {
    const parsed = JSON.parse(raw) as SessionSnapshot;
    if (typeof parsed.content !== 'string') {
      return null;
    }

    return parsed;
  } catch {
    return null;
  }
}

async function loadRecentFilesFromNative(snapshotRecentFiles: string[]) {
  const fallbackRecentFiles = normalizeRecentFiles(snapshotRecentFiles);

  try {
    const traceId = createTraceId('recent-files-load');
    const result = await bridgeService.getRecentFiles(traceId);
    const nativeRecentFiles = normalizeRecentFiles(result.filePaths);
    if (nativeRecentFiles.length > 0) {
      recentFiles.value = nativeRecentFiles;
      return;
    }

    if (fallbackRecentFiles.length > 0) {
      recentFiles.value = fallbackRecentFiles;
      const saved = await bridgeService.setRecentFiles(fallbackRecentFiles, traceId);
      recentFiles.value = normalizeRecentFiles(saved.filePaths);
      await logBridge('info', 'recent-files-migrated', traceId, {
        count: recentFiles.value.length,
      });
      return;
    }

    recentFiles.value = [];
  } catch (error) {
    recentFiles.value = fallbackRecentFiles;
    const message = error instanceof Error ? error.message : String(error);
    await logBridge('error', 'recent-files-load:fail', createTraceId('recent-files-load'), {
      message,
    });
  }
}

function clearContentZoomPersistTimer() {
  if (!contentZoomPersistTimer) {
    return;
  }

  window.clearTimeout(contentZoomPersistTimer);
  contentZoomPersistTimer = 0;
}

function schedulePersistContentZoom() {
  clearContentZoomPersistTimer();
  contentZoomPersistTimer = window.setTimeout(() => {
    contentZoomPersistTimer = 0;
    void bridgeService.setSetting('contentZoom', contentZoom.value)
      .catch(async (error) => {
        const message = error instanceof Error ? error.message : String(error);
        await logBridge('error', 'content-zoom-sync:fail', createTraceId('content-zoom'), {
          message,
          contentZoom: contentZoom.value,
        });
      });
  }, 180);
}

function setContentZoom(nextZoom: number) {
  const normalizedZoom = normalizeContentZoom(nextZoom);
  if (normalizedZoom === contentZoom.value) {
    return;
  }

  contentZoom.value = normalizedZoom;
  schedulePersistContentZoom();
  persistSessionSnapshot();
}

function stepContentZoom(direction: 'in' | 'out') {
  const delta = direction === 'in' ? CONTENT_ZOOM_STEP : -CONTENT_ZOOM_STEP;
  setContentZoom(contentZoom.value + delta);
}

function resetContentZoom() {
  setContentZoom(CONTENT_ZOOM_DEFAULT);
}

async function ensureMathPluginLoaded() {
  if (mathPluginRevision.value > 0) {
    return;
  }

  if (mathPluginLoadPromise) {
    await mathPluginLoadPromise;
    return;
  }

  mathPluginLoadPromise = import('markdown-it-mathjax3')
    .then((module) => {
      const plugin = (module.default ?? module) as (instance: MarkdownIt) => void;
      markdown.use(plugin);
      mathPluginRevision.value += 1;
    })
    .finally(() => {
      mathPluginLoadPromise = null;
    });

  await mathPluginLoadPromise;
}

async function ensureMermaidLoaded() {
  if (mermaidApi) {
    return mermaidApi;
  }

  if (mermaidLoadPromise) {
    await mermaidLoadPromise;
    return mermaidApi;
  }

  mermaidLoadPromise = import('mermaid')
    .then((module) => {
      const api = module.default;
      mermaidApi = {
        initialize: api.initialize.bind(api),
        render: api.render.bind(api),
      };
    })
    .finally(() => {
      mermaidLoadPromise = null;
    });

  await mermaidLoadPromise;
  return mermaidApi;
}

async function ensureFlowchartLoaded() {
  if (flowchartApi) {
    return flowchartApi;
  }

  if (flowchartLoadPromise) {
    await flowchartLoadPromise;
    return flowchartApi;
  }

  flowchartLoadPromise = import('flowchart.js')
    .then((module) => {
      const api = (module.default ?? module) as {
        parse?: (code: string) => {
          drawSVG: (container: HTMLElement | string, options?: Record<string, unknown>) => void;
        };
      };

      if (typeof api.parse !== 'function') {
        throw new Error('flowchart.js 未提供 parse 方法');
      }

      flowchartApi = {
        parse: api.parse,
      };
    })
    .finally(() => {
      flowchartLoadPromise = null;
    });

  await flowchartLoadPromise;
  return flowchartApi;
}

function decodeDiagramSource(encoded: string) {
  try {
    return decodeURIComponent(encoded);
  } catch {
    return encoded;
  }
}

function toMermaidSource(kind: string, source: string) {
  const trimmed = source.trimStart();
  if (kind === 'sequence') {
    if (/^sequenceDiagram\b/.test(trimmed)) {
      return source;
    }

    return `sequenceDiagram\n${source}`;
  }

  if (kind === 'flow') {
    if (/^(flowchart|graph)\b/.test(trimmed)) {
      return source;
    }

    return `flowchart TD\n${source}`;
  }

  return source;
}

async function renderMermaidSource(
  mermaid: NonNullable<typeof mermaidApi>,
  outputNode: HTMLElement,
  source: string,
) {
  mermaidRenderSeed += 1;
  const renderId = `unireadmd-mermaid-${mermaidRenderSeed}`;
  const result = await mermaid.render(renderId, source);
  outputNode.innerHTML = result.svg;
}

function renderFlowSource(
  flowchart: NonNullable<typeof flowchartApi>,
  outputNode: HTMLElement,
  source: string,
) {
  outputNode.innerHTML = '';
  const chart = flowchart.parse(source);
  chart.drawSVG(outputNode, {
    'line-width': 2,
    'font-size': 13,
    'font-family': 'Consolas, Microsoft YaHei UI, sans-serif',
    'yes-text': 'yes',
    'no-text': 'no',
  });
}

async function renderMermaidBlocks() {
  if (!previewHost.value) {
    return;
  }

  const blocks = Array.from(previewHost.value.querySelectorAll<HTMLElement>('.mermaid-block'));
  if (blocks.length === 0) {
    return;
  }

  const mermaid = await ensureMermaidLoaded();
  let flowchart: NonNullable<typeof flowchartApi> | null = null;
  const hasFlowBlock = blocks.some((block) => {
    return (block.dataset.diagramKind || 'mermaid') === 'flow';
  });
  if (hasFlowBlock) {
    try {
      flowchart = await ensureFlowchartLoaded();
    } catch (error) {
      console.warn(
        '[diagram] flowchart.js 加载失败，将尝试走 Mermaid 兼容渲染。',
        error,
      );
      flowchart = null;
    }
  }

  if (mermaid) {
    const nextThemeKey = theme.value === 'dark' ? 'dark' : 'default';
    if (mermaidThemeKey !== nextThemeKey) {
      mermaid.initialize({
        startOnLoad: false,
        securityLevel: 'strict',
        theme: nextThemeKey,
      });
      mermaidThemeKey = nextThemeKey;
    }
  }

  for (const block of blocks) {
    const kind = block.dataset.diagramKind || 'mermaid';
    const encoded = block.dataset.diagramCode || '';
    const source = decodeDiagramSource(encoded);
    const outputNode = block.querySelector<HTMLElement>('.mermaid-output');
    const fallbackNode = block.querySelector<HTMLElement>('.mermaid-fallback');
    const errorNode = block.querySelector<HTMLElement>('.mermaid-error');

    if (!source || !outputNode || !fallbackNode || !errorNode) {
      continue;
    }

    try {
      let rendered = false;
      if (kind === 'flow' && flowchart) {
        renderFlowSource(flowchart, outputNode, source);
        rendered = true;
      }

      if (!rendered) {
        if (!mermaid) {
          throw new Error('Mermaid 未加载成功');
        }

        const mermaidSource = toMermaidSource(kind, source);
        await renderMermaidSource(mermaid, outputNode, mermaidSource);
      }

      outputNode.style.display = 'block';
      fallbackNode.style.display = 'none';
      errorNode.style.display = 'none';
      errorNode.textContent = '';
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);

      outputNode.innerHTML = '';
      outputNode.style.display = 'none';
      fallbackNode.style.display = 'block';
      errorNode.style.display = 'block';
      errorNode.textContent = `${kind} 渲染失败：${message}`;
    }
  }
}

function createFontSizeExtension(size: number) {
  return EditorView.theme({
    '&': {
      fontSize: `${size}px`,
    },
  });
}

function applyStoreContentToEditor(nextContent: string) {
  if (!editorView) {
    return;
  }

  const currentContent = editorView.state.doc.toString();
  if (currentContent === nextContent) {
    return;
  }

  isSyncingFromStore = true;
  editorView.dispatch({
    changes: {
      from: 0,
      to: editorView.state.doc.length,
      insert: nextContent,
    },
  });
  isSyncingFromStore = false;
}

function scrollEditorToLine(lineNumber: number) {
  if (!editorView) {
    return;
  }

  const totalLines = editorView.state.doc.lines;
  const safeLine = Math.min(Math.max(lineNumber, 1), totalLines);
  const line = editorView.state.doc.line(safeLine);

  editorView.dispatch({
    selection: {
      anchor: line.from,
    },
    effects: EditorView.scrollIntoView(line.from, {
      y: 'center',
    }),
  });

  editorView.focus();
}

function resolvePreviewHeadingNode(lineNumber: number) {
  if (previewHeadingNodes.length === 0) {
    refreshPreviewHeadingNodes();
  }

  if (previewHeadingNodes.length === 0) {
    return null;
  }

  const exactHeading = previewHeadingNodes.find((item) => item.line === lineNumber);
  if (exactHeading) {
    return exactHeading.node;
  }

  let nearestNext: {
    line: number;
    node: HTMLElement;
  } | null = null;
  let nearestPrev: {
    line: number;
    node: HTMLElement;
  } | null = null;

  for (const item of previewHeadingNodes) {
    const line = item.line;
    if (line >= lineNumber) {
      if (!nearestNext || line < nearestNext.line) {
        nearestNext = {
          line,
          node: item.node,
        };
      }
      continue;
    }

    if (!nearestPrev || line > nearestPrev.line) {
      nearestPrev = {
        line,
        node: item.node,
      };
    }
  }

  if (nearestNext) {
    return nearestNext.node;
  }

  if (nearestPrev) {
    return nearestPrev.node;
  }

  return null;
}

function refreshPreviewSourceNodes() {
  const preview = previewHost.value;
  if (!preview) {
    previewSourceNodes = [];
    return;
  }

  const nodes = Array.from(
    preview.querySelectorAll<HTMLElement>('[data-source-start]'),
  );
  previewSourceNodes = nodes
    .map((node) => {
      const line = Number(node.dataset.sourceStart || 0);
      if (!Number.isFinite(line) || line <= 0) {
        return null;
      }

      return {
        line,
        node,
      };
    })
    .filter((item): item is {
      line: number;
      node: HTMLElement;
    } => item !== null);
}

function resolvePreviewNodeByLine(lineNumber: number) {
  const headingNode = resolvePreviewHeadingNode(lineNumber);
  if (headingNode) {
    return headingNode;
  }

  if (previewSourceNodes.length === 0) {
    refreshPreviewSourceNodes();
  }

  if (previewSourceNodes.length === 0) {
    return null;
  }

  const exact = previewSourceNodes.find((item) => item.line === lineNumber);
  if (exact) {
    return exact.node;
  }

  let nearestNext: {
    line: number;
    node: HTMLElement;
  } | null = null;
  let nearestPrev: {
    line: number;
    node: HTMLElement;
  } | null = null;

  for (const item of previewSourceNodes) {
    if (item.line >= lineNumber) {
      if (!nearestNext || item.line < nearestNext.line) {
        nearestNext = item;
      }
      continue;
    }

    if (!nearestPrev || item.line > nearestPrev.line) {
      nearestPrev = item;
    }
  }

  if (nearestNext) {
    return nearestNext.node;
  }

  if (nearestPrev) {
    return nearestPrev.node;
  }

  return previewSourceNodes[0].node;
}

function resolvePreviewViewportAnchorLine() {
  const preview = previewHost.value;
  if (!preview) {
    return 0;
  }

  if (previewSourceNodes.length === 0) {
    refreshPreviewSourceNodes();
  }

  if (previewSourceNodes.length === 0) {
    return 0;
  }

  const anchorTop = preview.scrollTop + 18;
  let activeLine = previewSourceNodes[0].line;
  for (const item of previewSourceNodes) {
    if (item.node.offsetTop <= anchorTop) {
      activeLine = item.line;
      continue;
    }

    break;
  }

  return activeLine;
}

function scrollPreviewToLine(lineNumber: number) {
  const preview = previewHost.value;
  if (!preview) {
    return;
  }

  const target = resolvePreviewNodeByLine(lineNumber);
  if (!target) {
    return;
  }

  const paddingTop = 16;
  const targetTop = Math.max(0, target.offsetTop - paddingTop);
  preview.scrollTo({
    top: targetTop,
    behavior: 'auto',
  });
}

function refreshPreviewHeadingNodes() {
  const preview = previewHost.value;
  if (!preview) {
    previewHeadingNodes = [];
    return;
  }

  const nodes = Array.from(
    preview.querySelectorAll<HTMLElement>(
      "[data-source-token='heading_open'][data-source-start]",
    ),
  );
  previewHeadingNodes = nodes
    .map((node) => {
      const line = Number(node.dataset.sourceStart || 0);
      if (!Number.isFinite(line) || line <= 0) {
        return null;
      }

      return {
        line,
        node,
      };
    })
    .filter((item): item is {
      line: number;
      node: HTMLElement;
    } => item !== null)
    .sort((left, right) => left.line - right.line);
}

function syncCursorLineFromPreviewViewport() {
  const activeLine = resolvePreviewViewportAnchorLine();

  if (activeLine > 0 && currentCursorLine.value !== activeLine) {
    currentCursorLine.value = activeLine;
  }
}

function schedulePreviewScrollSync() {
  if (previewScrollSyncFrame) {
    return;
  }

  previewScrollSyncFrame = requestAnimationFrame(() => {
    previewScrollSyncFrame = 0;
    syncCursorLineFromPreviewViewport();
  });
}

function syncCursorLineFromState(state: EditorState) {
  currentCursorLine.value = state.doc.lineAt(state.selection.main.head).number;
}

function normalizeLineRange(startLine: number, endLine: number) {
  if (!editorView) {
    return null;
  }

  const totalLines = editorView.state.doc.lines;
  if (totalLines <= 0) {
    return null;
  }

  const safeStart = Math.min(Math.max(startLine, 1), totalLines);
  const safeEnd = Math.min(Math.max(endLine, safeStart), totalLines);
  return {
    startLine: safeStart,
    endLine: safeEnd,
  };
}

function getEditorRangeByLines(startLine: number, endLine: number) {
  if (!editorView) {
    return null;
  }

  const lineRange = normalizeLineRange(startLine, endLine);
  if (!lineRange) {
    return null;
  }

  const doc = editorView.state.doc;
  const from = doc.line(lineRange.startLine).from;
  const to = doc.line(lineRange.endLine).to;
  return {
    ...lineRange,
    from,
    to,
  };
}

function getEditorFragmentByLines(startLine: number, endLine: number) {
  const range = getEditorRangeByLines(startLine, endLine);
  if (!range || !editorView) {
    return '';
  }

  return editorView.state.doc.sliceString(range.from, range.to);
}

function normalizePreviewInlineEditContent(content: string) {
  return content.replace(/\r\n/g, '\n');
}

function getPreviewInlineEditLineCount(content: string) {
  const normalized = normalizePreviewInlineEditContent(content);
  if (normalized.length === 0) {
    return 1;
  }

  return normalized.split('\n').length;
}

function buildPreviewInlineEditStyle(node: HTMLElement) {
  const container = editorPreviewHost.value;
  const preview = previewHost.value;
  if (!container || !preview) {
    return null;
  }

  const containerRect = container.getBoundingClientRect();
  const previewRect = preview.getBoundingClientRect();
  const targetRect = node.getBoundingClientRect();
  const left = Math.max(targetRect.left, previewRect.left) - containerRect.left;
  const top = targetRect.top - containerRect.top;
  const availableWidth = Math.max(120, previewRect.right - containerRect.left - left - 2);
  const targetWidth = Math.max(120, Math.round(targetRect.width) + 4);
  const width = Math.min(targetWidth, availableWidth);
  const minHeight = Math.max(24, Math.round(targetRect.height));

  return {
    left: `${left}px`,
    top: `${top}px`,
    width: `${width}px`,
    minHeight: `${minHeight}px`,
  };
}

function findPreviewInlineEditAnchorNode() {
  const preview = previewHost.value;
  const selection = previewInlineEditSelection.value;
  if (!preview || !selection) {
    return null;
  }

  const startLineSelector = `[data-source-start='${selection.startLine}']`;
  const tokenSelector = `${startLineSelector}[data-source-token='${selection.token}']`;
  const tokenMatched = preview.querySelector<HTMLElement>(tokenSelector);
  if (tokenMatched) {
    return tokenMatched;
  }

  const startLineMatched = preview.querySelector<HTMLElement>(startLineSelector);
  if (startLineMatched) {
    return startLineMatched;
  }

  const mappedNodes = Array.from(
    preview.querySelectorAll<HTMLElement>('[data-source-start][data-source-end][data-source-token]'),
  );
  return mappedNodes.find((node) => {
    const start = Number(node.dataset.sourceStart || 0);
    const end = Number(node.dataset.sourceEnd || 0);
    return Number.isFinite(start)
      && Number.isFinite(end)
      && start <= selection.startLine
      && end >= selection.startLine;
  }) || null;
}

function refreshPreviewInlineEditorPosition() {
  if (!previewInlineEditVisible.value || !previewInlineEditSelection.value) {
    return;
  }

  const anchorNode = findPreviewInlineEditAnchorNode();
  if (!anchorNode) {
    return;
  }

  const nextStyle = buildPreviewInlineEditStyle(anchorNode);
  if (!nextStyle) {
    return;
  }

  previewInlineEditStyle.value = nextStyle;
  const nextToken = anchorNode.dataset.sourceToken || '';
  if (!nextToken) {
    return;
  }

  const selection = previewInlineEditSelection.value;
  if (!selection || selection.token === nextToken) {
    return;
  }

  previewInlineEditSelection.value = {
    ...selection,
    token: nextToken,
  };
}

function clearPreviewInlineEditSyncTimer() {
  if (!previewInlineEditSyncTimer) {
    return;
  }

  window.clearTimeout(previewInlineEditSyncTimer);
  previewInlineEditSyncTimer = 0;
}

function applyPreviewInlineEditToSource() {
  if (!previewInlineEditVisible.value || !editorView) {
    return;
  }

  const session = previewInlineEditSession.value;
  const selection = previewInlineEditSelection.value;
  if (!session || !selection) {
    return;
  }

  const insert = normalizePreviewInlineEditContent(previewInlineEditValue.value);
  if (insert === session.appliedText) {
    return;
  }

  editorView.dispatch({
    changes: {
      from: session.from,
      to: session.to,
      insert,
    },
    selection: {
      anchor: session.from + insert.length,
    },
    annotations: Transaction.userEvent.of('input.type'),
  });

  const lineCount = getPreviewInlineEditLineCount(insert);
  previewInlineEditSession.value = {
    ...session,
    to: session.from + insert.length,
    appliedText: insert,
  };
  previewInlineEditSelection.value = {
    ...selection,
    endLine: session.startLine + lineCount - 1,
  };
}

function flushPreviewInlineEditSync() {
  clearPreviewInlineEditSyncTimer();
  applyPreviewInlineEditToSource();
}

function queuePreviewInlineEditSync() {
  clearPreviewInlineEditSyncTimer();
  previewInlineEditSyncTimer = window.setTimeout(() => {
    previewInlineEditSyncTimer = 0;
    applyPreviewInlineEditToSource();
  }, PREVIEW_INLINE_EDIT_SYNC_DELAY);
}

function closePreviewInlineEditor(
  options?: {
    revert?: boolean;
  },
) {
  const currentSession = previewInlineEditSession.value;
  const shouldRevert = Boolean(options?.revert);
  if (shouldRevert) {
    clearPreviewInlineEditSyncTimer();
  } else {
    flushPreviewInlineEditSync();
  }

  previewInlineEditComposing.value = false;
  previewInlineEditVisible.value = false;
  previewInlineEditSelection.value = null;
  previewInlineEditStyle.value = {};
  previewInlineEditValue.value = '';

  if (currentSession && editorView) {
    if (shouldRevert && currentSession.appliedText !== currentSession.originalText) {
      editorView.dispatch({
        changes: {
          from: currentSession.from,
          to: currentSession.to,
          insert: currentSession.originalText,
        },
        selection: {
          anchor: currentSession.from + currentSession.originalText.length,
        },
        annotations: Transaction.addToHistory.of(false),
      });
    }
  }

  previewInlineEditSession.value = null;
}

function cancelPreviewInlineEdit() {
  closePreviewInlineEditor({
    revert: true,
  });
}

function focusPreviewInlineEditor() {
  void nextTick(() => {
    const textarea = previewInlineEditTextarea.value;
    if (!textarea) {
      return;
    }

    textarea.focus();
    const valueLength = textarea.value.length;
    textarea.setSelectionRange(valueLength, valueLength);
  });
}

function findSourceMappedNode(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return null;
  }

  return target.closest<HTMLElement>('[data-source-start][data-source-end][data-source-token]');
}

function parseInlineEditSelection(node: HTMLElement) {
  const rawStart = Number(node.dataset.sourceStart || 0);
  const rawEnd = Number(node.dataset.sourceEnd || 0);
  const token = node.dataset.sourceToken || '';
  if (!Number.isFinite(rawStart) || rawStart <= 0) {
    return null;
  }

  if (!Number.isFinite(rawEnd) || rawEnd <= 0) {
    return null;
  }

  if (!token) {
    return null;
  }

  return {
    startLine: Math.round(rawStart),
    endLine: Math.round(rawEnd),
    token,
  };
}

function resolvePreviewTaskItemContext(target: HTMLElement | null) {
  if (!target) {
    return null;
  }

  let checkbox = target.closest<HTMLInputElement>('.task-list-item-checkbox');
  if (!checkbox) {
    const label = target.closest<HTMLLabelElement>('label.task-list-item-label[for]');
    if (label) {
      const labelFor = label.getAttribute('for');
      const candidate = labelFor ? document.getElementById(labelFor) : null;
      if (candidate instanceof HTMLInputElement && candidate.classList.contains('task-list-item-checkbox')) {
        checkbox = candidate;
      }
    }
  }

  if (!checkbox) {
    return null;
  }

  let mappedNode = checkbox.closest<HTMLElement>(
    "[data-source-token='list_item_open'][data-source-start][data-source-end]",
  );
  if (!mappedNode) {
    mappedNode = checkbox.closest<HTMLElement>('[data-source-start][data-source-end]');
  }

  if (!mappedNode) {
    let parent = checkbox.parentElement;
    while (parent) {
      if (parent.dataset.sourceStart && parent.dataset.sourceEnd) {
        mappedNode = parent;
        break;
      }
      parent = parent.parentElement;
    }
  }

  if (!mappedNode) {
    return null;
  }

  const startLine = Number(mappedNode.dataset.sourceStart || 0);
  const endLine = Number(mappedNode.dataset.sourceEnd || 0);
  if (!Number.isFinite(startLine) || !Number.isFinite(endLine) || startLine <= 0 || endLine <= 0) {
    return null;
  }

  const taskCheckboxes = Array.from(mappedNode.querySelectorAll<HTMLInputElement>('.task-list-item-checkbox'));
  const resolvedTaskIndex = Math.max(0, taskCheckboxes.indexOf(checkbox));
  const allTaskCheckboxes = previewHost.value
    ? Array.from(previewHost.value.querySelectorAll<HTMLInputElement>('.task-list-item-checkbox'))
    : [];
  const resolvedGlobalTaskIndex = Math.max(0, allTaskCheckboxes.indexOf(checkbox));

  return {
    startLine: Math.round(startLine),
    endLine: Math.round(endLine),
    taskIndex: resolvedTaskIndex,
    globalTaskIndex: resolvedGlobalTaskIndex,
  } as PreviewTaskItemContext;
}

function togglePreviewTaskItem(context: PreviewTaskItemContext) {
  if (!editorView) {
    return false;
  }

  const range = getEditorRangeByLines(context.startLine, context.endLine);
  if (!range) {
    return false;
  }

  const doc = editorView.state.doc;
  const taskPrefixPattern = /^(\s*(?:[-+*]|\d+[.)])\s+\[)([ xX])(\]\s*)/;
  const allMatchedLineNumbers: number[] = [];
  for (let lineNo = 1; lineNo <= doc.lines; lineNo += 1) {
    const line = doc.line(lineNo);
    if (taskPrefixPattern.test(line.text)) {
      allMatchedLineNumbers.push(lineNo);
    }
  }
  if (allMatchedLineNumbers.length === 0) {
    errorMessage.value = '未定位到任务列表项源码行，请切换 Source 手动调整。';
    return false;
  }

  const globalLineNo = allMatchedLineNumbers[
    Math.min(context.globalTaskIndex, allMatchedLineNumbers.length - 1)
  ];
  const globalLine = doc.line(globalLineNo);
  const globalNextLineText = globalLine.text.replace(taskPrefixPattern, (_match, start, current, end) => {
    const marker = current === ' ' ? 'x' : ' ';
    return `${start}${marker}${end}`;
  });
  if (globalNextLineText !== globalLine.text) {
    editorView.dispatch({
      changes: {
        from: globalLine.from,
        to: globalLine.to,
        insert: globalNextLineText,
      },
      selection: {
        anchor: globalLine.from + globalNextLineText.length,
      },
      annotations: Transaction.userEvent.of('input.type'),
    });
    errorMessage.value = '';
    return true;
  }

  const minLine = Math.max(1, context.startLine - 1);
  const maxLine = Math.min(doc.lines, context.endLine + 1);

  const matchedLineNumbers: number[] = [];
  for (let lineNo = minLine; lineNo <= maxLine; lineNo += 1) {
    const line = doc.line(lineNo);
    if (taskPrefixPattern.test(line.text)) {
      matchedLineNumbers.push(lineNo);
    }
  }
  if (matchedLineNumbers.length === 0) {
    errorMessage.value = '未定位到任务列表项源码行，请切换 Source 手动调整。';
    return false;
  }

  const lineNo = matchedLineNumbers[Math.min(context.taskIndex, matchedLineNumbers.length - 1)];
  const line = doc.line(lineNo);
  const nextLineText = line.text.replace(taskPrefixPattern, (_match, start, current, end) => {
    const marker = current === ' ' ? 'x' : ' ';
    return `${start}${marker}${end}`;
  });
  if (nextLineText === line.text) {
    return true;
  }

  editorView.dispatch({
    changes: {
      from: line.from,
      to: line.to,
      insert: nextLineText,
    },
    selection: {
      anchor: line.from + nextLineText.length,
    },
    annotations: Transaction.userEvent.of('input.type'),
  });
  errorMessage.value = '';
  return true;

}

function beginPreviewInlineEdit(
  node: HTMLElement,
  options?: {
    silentUnsupported?: boolean;
  },
) {
  if (previewInlineEditVisible.value) {
    closePreviewInlineEditor();
  }

  const selection = parseInlineEditSelection(node);
  if (!selection) {
    return;
  }

  if (!PREVIEW_INLINE_EDITABLE_TOKENS.has(selection.token)) {
    if (!options?.silentUnsupported) {
      const label = PREVIEW_TOKEN_LABELS[selection.token] || selection.token;
      errorMessage.value = `渲染态暂不支持直接编辑「${label}」，请切换 Source 编辑。`;
    }
    return;
  }

  const range = normalizeLineRange(selection.startLine, selection.endLine);
  if (!range) {
    return;
  }

  const editorRange = getEditorRangeByLines(range.startLine, range.endLine);
  if (!editorRange) {
    return;
  }

  const nextValue = getEditorFragmentByLines(range.startLine, range.endLine);
  if (!nextValue && !editorView) {
    return;
  }

  const nextStyle = buildPreviewInlineEditStyle(node);
  if (!nextStyle) {
    return;
  }

  const normalizedValue = normalizePreviewInlineEditContent(nextValue);

  previewInlineEditSelection.value = {
    startLine: range.startLine,
    endLine: range.endLine,
    token: selection.token,
  };
  previewInlineEditSession.value = {
    from: editorRange.from,
    to: editorRange.to,
    startLine: range.startLine,
    originalText: normalizedValue,
    appliedText: normalizedValue,
  };
  previewInlineEditStyle.value = nextStyle;
  previewInlineEditValue.value = normalizedValue;
  previewInlineEditVisible.value = true;
  errorMessage.value = '';
  focusPreviewInlineEditor();
}

function commitPreviewInlineEdit() {
  closePreviewInlineEditor();
}

function handlePreviewDoubleClick(event: MouseEvent) {
  if (sourceMode.value) {
    return;
  }

  // 预览态禁用双击进入源码覆盖编辑，避免与渲染内容重叠。
  if (previewInlineEditVisible.value) {
    closePreviewInlineEditor();
  }
}

function handlePreviewClick(event: MouseEvent) {
  if (sourceMode.value) {
    return;
  }

  const target = event.target instanceof HTMLElement ? event.target : null;
  const taskItemContext = resolvePreviewTaskItemContext(target);
  if (taskItemContext) {
    event.preventDefault();
    event.stopPropagation();
    closePreviewTableMenu();
    closePreviewTableCellEditorAndCommit();
    closeSourceContextMenu();
    closePreviewInlineEditor();
    void togglePreviewTaskItem(taskItemContext);
    return;
  }

  const tableContext = resolvePreviewTableCellContext(target);
  if (tableContext) {
    event.preventDefault();
    openPreviewTableCellEditor({
      startLine: tableContext.startLine,
      endLine: tableContext.endLine,
      rowIndex: tableContext.rowIndex,
      colIndex: tableContext.colIndex,
    });
    return;
  }

  const anchor = target?.closest('a[href]');
  if (anchor && (event.ctrlKey || event.metaKey)) {
    return;
  }

  if (anchor) {
    event.preventDefault();
  }

  const node = findSourceMappedNode(event.target);
  if (!node) {
    if (previewInlineEditVisible.value) {
      closePreviewInlineEditor();
    }
    return;
  }

  const selection = parseInlineEditSelection(node);
  if (!selection) {
    return;
  }

  currentCursorLine.value = selection.startLine;
  store.setCursorHint(selection.startLine);

  const currentSelection = previewInlineEditSelection.value;
  const isSameInlineBlock = currentSelection
    && currentSelection.startLine === selection.startLine
    && currentSelection.endLine === selection.endLine
    && currentSelection.token === selection.token;
  if (previewInlineEditVisible.value && !isSameInlineBlock) {
    closePreviewInlineEditor();
  }
}

function handlePreviewInlineEditInput() {
  if (!previewInlineEditVisible.value || !previewInlineEditSession.value) {
    return;
  }

  if (previewInlineEditComposing.value) {
    return;
  }

  queuePreviewInlineEditSync();
}

function handlePreviewInlineEditCompositionStart() {
  previewInlineEditComposing.value = true;
}

function handlePreviewInlineEditCompositionEnd() {
  previewInlineEditComposing.value = false;
  flushPreviewInlineEditSync();
}

function handlePreviewInlineEditBlur() {
  if (!previewInlineEditVisible.value || previewInlineEditComposing.value) {
    return;
  }

  closePreviewInlineEditor();
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

function handlePreviewInlineEditorWheel(event: WheelEvent) {
  const preview = previewHost.value;
  if (!preview) {
    return;
  }

  event.preventDefault();
  preview.scrollTop += event.deltaY;
}

function handlePreviewScroll() {
  schedulePreviewScrollSync();
  closePreviewTableMenu();
  if (previewTableCellEditVisible.value) {
    commitPreviewTableCellEditValue();
    closePreviewTableCellEditor();
  }

  if (!previewInlineEditVisible.value) {
    return;
  }

  closePreviewInlineEditor();
}

function closePreviewTableMenu() {
  previewTableMenuVisible.value = false;
  previewTableMenuStyle.value = {};
  previewTableSelection.value = null;
}

function closeSourceContextMenu() {
  sourceContextMenuVisible.value = false;
  sourceContextMenuStyle.value = {};
}

function selectAllSourceContent() {
  if (!editorView) {
    return;
  }

  const total = editorView.state.doc.length;
  editorView.dispatch({
    selection: {
      anchor: 0,
      head: total,
    },
  });
  editorView.focus();
}

function runLegacyEditCommand(command: 'cut' | 'copy' | 'paste' | 'selectAll') {
  try {
    return document.execCommand(command);
  } catch {
    return false;
  }
}

function applySourceInlineShortcut(
  kind: InlineShortcutKind,
  viewArg?: EditorView | null,
) {
  const targetView = viewArg || editorView;
  if (!targetView) {
    return false;
  }

  const selection = targetView.state.selection.main;
  const selectionText = targetView.state.doc.sliceString(selection.from, selection.to);
  const patch = buildInlineShortcutPatch(selectionText, kind);
  if (!patch.changed) {
    return false;
  }

  targetView.dispatch({
    changes: {
      from: selection.from,
      to: selection.to,
      insert: patch.insert,
    },
    selection: {
      anchor: selection.from + patch.selectionStart,
      head: selection.from + patch.selectionEnd,
    },
    annotations: Transaction.userEvent.of('input.type'),
  });
  return true;
}

function applySourceIndentShortcut(
  outdent: boolean,
  viewArg?: EditorView | null,
) {
  const targetView = viewArg || editorView;
  if (!targetView) {
    return false;
  }

  if (outdent) {
    return indentLess(targetView);
  }

  return indentMore(targetView);
}

function updatePreviewInlineEditorSelection(
  selectionStart: number,
  selectionEnd: number,
) {
  void nextTick(() => {
    const textarea = previewInlineEditTextarea.value;
    if (!textarea) {
      return;
    }

    const valueLength = previewInlineEditValue.value.length;
    const start = Math.min(Math.max(selectionStart, 0), valueLength);
    const end = Math.min(Math.max(selectionEnd, start), valueLength);
    textarea.focus();
    textarea.setSelectionRange(start, end);
  });
}

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

async function handleSourceContextAction(action: string) {
  if (!sourceMode.value || !editorView) {
    closeSourceContextMenu();
    return;
  }

  if (action === 'search') {
    openDocSearchPanel(false);
    closeSourceContextMenu();
    return;
  }

  editorView.focus();
  if (action === 'select-all') {
    selectAllSourceContent();
    closeSourceContextMenu();
    return;
  }

  const selection = editorView.state.selection.main;
  if (action === 'copy') {
    const text = editorView.state.doc.sliceString(selection.from, selection.to);
    if (text.length > 0) {
      try {
        await navigator.clipboard.writeText(text);
      } catch {
        runLegacyEditCommand('copy');
      }
    } else {
      runLegacyEditCommand('copy');
    }
    closeSourceContextMenu();
    return;
  }

  if (action === 'cut') {
    const text = editorView.state.doc.sliceString(selection.from, selection.to);
    if (text.length > 0) {
      let manualCopied = false;
      try {
        await navigator.clipboard.writeText(text);
        manualCopied = true;
      } catch {
        runLegacyEditCommand('cut');
      }

      if (manualCopied) {
        editorView.dispatch({
          changes: {
            from: selection.from,
            to: selection.to,
            insert: '',
          },
          selection: {
            anchor: selection.from,
          },
        });
      }
    } else {
      runLegacyEditCommand('cut');
    }

    closeSourceContextMenu();
    return;
  }

  if (action === 'paste') {
    try {
      const text = await navigator.clipboard.readText();
      if (typeof text === 'string') {
        editorView.dispatch({
          changes: {
            from: selection.from,
            to: selection.to,
            insert: text,
          },
          selection: {
            anchor: selection.from + text.length,
          },
        });
      }
    } catch {
      runLegacyEditCommand('paste');
    }

    closeSourceContextMenu();
  }
}

function closePreviewTableCellEditor() {
  previewTableCellEditVisible.value = false;
  previewTableCellEditSelection.value = null;
  previewTableCellEditStyle.value = {};
  previewTableCellEditValue.value = '';
}

function focusPreviewTableCellEditor() {
  void nextTick(() => {
    const input = previewTableCellEditInput.value;
    if (!input) {
      return;
    }

    input.focus();
    input.select();
  });
}

function buildPreviewTableCellEditStyle(cell: HTMLElement) {
  const container = editorPreviewHost.value;
  if (!container) {
    return null;
  }

  const containerRect = container.getBoundingClientRect();
  const cellRect = cell.getBoundingClientRect();
  return {
    left: `${cellRect.left - containerRect.left}px`,
    top: `${cellRect.top - containerRect.top}px`,
    width: `${Math.max(92, Math.round(cellRect.width))}px`,
    minHeight: `${Math.max(30, Math.round(cellRect.height))}px`,
  };
}

function resolvePreviewTableCellContext(target: HTMLElement | null) {
  if (!target) {
    return null;
  }

  const cell = target.closest<HTMLElement>('th,td');
  const table = target.closest<HTMLElement>('table[data-source-start][data-source-end]');
  if (!cell || !table) {
    return null;
  }

  const row = cell.closest('tr');
  if (!row) {
    return null;
  }

  const rows = Array.from(table.querySelectorAll('tr'));
  const rowIndex = rows.indexOf(row);
  if (rowIndex < 0) {
    return null;
  }

  const cells = Array.from(row.querySelectorAll('th,td'));
  const colIndex = cells.indexOf(cell);
  if (colIndex < 0) {
    return null;
  }

  const startLine = Number(table.dataset.sourceStart || 0);
  const endLine = Number(table.dataset.sourceEnd || 0);
  if (!Number.isFinite(startLine) || !Number.isFinite(endLine) || startLine <= 0 || endLine <= 0) {
    return null;
  }

  return {
    table,
    cell,
    startLine: Math.round(startLine),
    endLine: Math.round(endLine),
    rowIndex,
    colIndex,
  } as PreviewTableCellContext;
}

function findPreviewTableCellNode(selection: PreviewTableSelection) {
  const preview = previewHost.value;
  if (!preview) {
    return null;
  }

  const tableSelector = `table[data-source-start='${selection.startLine}'][data-source-end='${selection.endLine}']`;
  const table = preview.querySelector<HTMLElement>(tableSelector);
  if (!table) {
    return null;
  }

  const rows = Array.from(table.querySelectorAll('tr'));
  const row = rows[selection.rowIndex];
  if (!row) {
    return null;
  }

  const cells = Array.from(row.querySelectorAll<HTMLElement>('th,td'));
  return cells[selection.colIndex] || null;
}

function readPreviewTableCellValue(selection: PreviewTableSelection) {
  if (!editorView) {
    return '';
  }

  const range = getEditorRangeByLines(selection.startLine, selection.endLine);
  if (!range) {
    return '';
  }

  const fragment = editorView.state.doc.sliceString(range.from, range.to);
  const table = parseMarkdownTable(fragment);
  if (!table) {
    return '';
  }

  if (selection.rowIndex <= 0) {
    return table.header[selection.colIndex] || '';
  }

  const bodyRow = table.rows[selection.rowIndex - 1];
  return bodyRow?.[selection.colIndex] || '';
}

function refreshPreviewTableCellEditorPosition() {
  if (!previewTableCellEditVisible.value || !previewTableCellEditSelection.value) {
    return;
  }

  const cell = findPreviewTableCellNode(previewTableCellEditSelection.value);
  if (!cell) {
    return;
  }

  const nextStyle = buildPreviewTableCellEditStyle(cell);
  if (!nextStyle) {
    return;
  }

  previewTableCellEditStyle.value = nextStyle;
}

function openPreviewTableCellEditor(selection: PreviewTableSelection) {
  closePreviewTableCellEditorAndCommit();

  const cell = findPreviewTableCellNode(selection);
  if (!cell) {
    return;
  }

  const nextStyle = buildPreviewTableCellEditStyle(cell);
  if (!nextStyle) {
    return;
  }

  closePreviewInlineEditor();
  closePreviewTableMenu();
  closeSourceContextMenu();
  previewTableSelection.value = {
    ...selection,
  };
  previewTableCellEditSelection.value = {
    ...selection,
  };
  previewTableCellEditStyle.value = nextStyle;
  previewTableCellEditValue.value = readPreviewTableCellValue(selection);
  previewTableCellEditVisible.value = true;
  errorMessage.value = '';
  focusPreviewTableCellEditor();
}

function mutatePreviewTable(
  selection: PreviewTableSelection,
  mutator: (table: ParsedMarkdownTable, selection: PreviewTableSelection) => boolean,
) {
  if (!editorView) {
    return null;
  }

  const range = getEditorRangeByLines(selection.startLine, selection.endLine);
  if (!range) {
    return null;
  }

  const fragment = editorView.state.doc.sliceString(range.from, range.to);
  const table = parseMarkdownTable(fragment);
  if (!table) {
    errorMessage.value = '无法解析当前表格源码，请切换 Source 模式手动调整。';
    return null;
  }

  if (!mutator(table, selection)) {
    return {
      changed: false,
      selection,
      table,
    };
  }

  const insert = serializeMarkdownTable(table);
  const lineCount = insert.split('\n').length;
  const nextSelection = {
    ...selection,
    endLine: selection.startLine + lineCount - 1,
  };
  editorView.dispatch({
    changes: {
      from: range.from,
      to: range.to,
      insert,
    },
    selection: {
      anchor: range.from,
    },
  });

  previewTableSelection.value = {
    ...nextSelection,
  };
  if (previewTableCellEditSelection.value) {
    previewTableCellEditSelection.value = {
      ...previewTableCellEditSelection.value,
      endLine: nextSelection.endLine,
    };
  }

  return {
    changed: true,
    selection: nextSelection,
    table,
  };
}

function applyPreviewTableMutation(
  mutator: (table: ParsedMarkdownTable, selection: PreviewTableSelection) => boolean,
) {
  if (!previewTableSelection.value) {
    closePreviewTableMenu();
    return;
  }

  const result = mutatePreviewTable(previewTableSelection.value, mutator);
  if (!result) {
    closePreviewTableMenu();
    return;
  }

  closePreviewTableMenu();
}

function applyPreviewTableAction(action: string) {
  const tableAction = action as PreviewTableAction;
  applyPreviewTableMutation((table, selection) => {
    const result = applyPreviewTableActionToModel(
      table,
      {
        rowIndex: selection.rowIndex,
        colIndex: selection.colIndex,
      },
      tableAction,
    );
    if (!result.changed && result.errorMessage) {
      errorMessage.value = result.errorMessage;
    }
    return result.changed;
  });
}

function handlePreviewContextMenu(event: MouseEvent) {
  if (sourceMode.value) {
    return;
  }

  const target = event.target instanceof HTMLElement ? event.target : null;
  const context = resolvePreviewTableCellContext(target);
  if (!context) {
    closePreviewTableMenu();
    return;
  }

  const container = editorPreviewHost.value;
  if (!container) {
    closePreviewTableMenu();
    return;
  }

  event.preventDefault();
  closePreviewTableCellEditorAndCommit();
  closeSourceContextMenu();
  closePreviewInlineEditor();
  previewTableSelection.value = {
    startLine: context.startLine,
    endLine: context.endLine,
    rowIndex: context.rowIndex,
    colIndex: context.colIndex,
  };
  const containerRect = container.getBoundingClientRect();
  previewTableMenuStyle.value = {
    left: `${event.clientX - containerRect.left}px`,
    top: `${event.clientY - containerRect.top}px`,
  };
  previewTableMenuVisible.value = true;
}

function commitPreviewTableCellEditValue() {
  const selection = previewTableCellEditSelection.value;
  if (!selection) {
    return null;
  }

  const nextValue = previewTableCellEditValue.value.replace(/\r?\n/g, ' ').trim();
  const result = mutatePreviewTable(selection, (table, currentSelection) => {
    if (currentSelection.rowIndex <= 0) {
      const current = table.header[currentSelection.colIndex] || '';
      if (current === nextValue) {
        return false;
      }
      table.header[currentSelection.colIndex] = nextValue;
      return true;
    }

    const bodyRowIndex = currentSelection.rowIndex - 1;
    if (!table.rows[bodyRowIndex]) {
      return false;
    }

    const current = table.rows[bodyRowIndex][currentSelection.colIndex] || '';
    if (current === nextValue) {
      return false;
    }
    table.rows[bodyRowIndex][currentSelection.colIndex] = nextValue;
    return true;
  });

  if (!result) {
    return null;
  }

  if (result.changed) {
    previewTableCellEditSelection.value = {
      ...result.selection,
    };
    previewTableCellEditValue.value = nextValue;
  }

  return {
    selection: result.selection,
    table: result.table,
  };
}

function getPreviewTableSelectionAfterAction(
  action: PreviewTableAction,
  selection: PreviewTableSelection,
  table: ParsedMarkdownTable,
) {
  const rowCount = table.rows.length + 1;
  const colCount = table.header.length;
  let rowIndex = selection.rowIndex;
  let colIndex = selection.colIndex;

  if (action === 'insert-row-below') {
    rowIndex = Math.min(selection.rowIndex + 1, rowCount - 1);
  } else if (action === 'insert-row-above') {
    rowIndex = selection.rowIndex;
  } else if (action === 'delete-row') {
    rowIndex = Math.min(selection.rowIndex, rowCount - 1);
  } else if (action === 'move-row-up') {
    rowIndex = Math.max(0, selection.rowIndex - 1);
  } else if (action === 'move-row-down') {
    rowIndex = Math.min(rowCount - 1, selection.rowIndex + 1);
  } else if (action === 'insert-col-right') {
    colIndex = Math.min(selection.colIndex + 1, colCount - 1);
  } else if (action === 'insert-col-left') {
    colIndex = selection.colIndex;
  } else if (action === 'delete-col') {
    colIndex = Math.min(selection.colIndex, colCount - 1);
  } else if (action === 'move-col-left') {
    colIndex = Math.max(0, selection.colIndex - 1);
  } else if (action === 'move-col-right') {
    colIndex = Math.min(colCount - 1, selection.colIndex + 1);
  }

  return {
    ...selection,
    rowIndex,
    colIndex,
  };
}

async function applyPreviewTableActionFromCellEditor(action: PreviewTableAction) {
  const activeSelection = previewTableCellEditSelection.value;
  if (!activeSelection) {
    return;
  }

  const commitResult = commitPreviewTableCellEditValue();
  const baseSelection = commitResult?.selection || activeSelection;
  const result = mutatePreviewTable(baseSelection, (table, selection) => {
    const actionResult = applyPreviewTableActionToModel(
      table,
      {
        rowIndex: selection.rowIndex,
        colIndex: selection.colIndex,
      },
      action,
    );
    if (!actionResult.changed && actionResult.errorMessage) {
      errorMessage.value = actionResult.errorMessage;
    }
    return actionResult.changed;
  });
  if (!result) {
    closePreviewTableCellEditor();
    return;
  }

  const nextSelection = getPreviewTableSelectionAfterAction(action, result.selection, result.table);
  closePreviewTableCellEditor();
  await nextTick();
  openPreviewTableCellEditor(nextSelection);
}

function getPreviewTableMetrics(selection: PreviewTableSelection) {
  if (!editorView) {
    return null;
  }

  const range = getEditorRangeByLines(selection.startLine, selection.endLine);
  if (!range) {
    return null;
  }

  const fragment = editorView.state.doc.sliceString(range.from, range.to);
  const table = parseMarkdownTable(fragment);
  if (!table) {
    return null;
  }

  return {
    rowCount: table.rows.length + 1,
    colCount: table.header.length,
  };
}

function closePreviewTableCellEditorAndCommit() {
  if (previewTableCellEditVisible.value) {
    commitPreviewTableCellEditValue();
  }
  closePreviewTableCellEditor();
}

async function movePreviewTableCellEditor(direction: 'tab-next' | 'tab-prev' | 'enter-down') {
  const activeSelection = previewTableCellEditSelection.value;
  if (!activeSelection) {
    return;
  }

  const commitResult = commitPreviewTableCellEditValue();
  const committedSelection = commitResult?.selection || activeSelection;
  const metrics = getPreviewTableMetrics(committedSelection);
  if (!metrics) {
    closePreviewTableCellEditor();
    return;
  }

  let targetRow = committedSelection.rowIndex;
  let targetCol = committedSelection.colIndex;
  let needAppendRow = false;
  if (direction === 'tab-next') {
    if (targetCol < metrics.colCount - 1) {
      targetCol += 1;
    } else if (targetRow < metrics.rowCount - 1) {
      targetRow += 1;
      targetCol = 0;
    } else {
      needAppendRow = true;
      targetRow += 1;
      targetCol = 0;
    }
  } else if (direction === 'tab-prev') {
    if (targetCol > 0) {
      targetCol -= 1;
    } else if (targetRow > 0) {
      targetRow -= 1;
      targetCol = metrics.colCount - 1;
    }
  } else if (direction === 'enter-down') {
    if (targetRow < metrics.rowCount - 1) {
      targetRow += 1;
    } else {
      needAppendRow = true;
      targetRow += 1;
    }
  }

  let finalSelection: PreviewTableSelection = {
    ...committedSelection,
  };
  if (needAppendRow) {
    const appendResult = mutatePreviewTable(finalSelection, (table, currentSelection) => {
      const result = applyPreviewTableActionToModel(
        table,
        {
          rowIndex: currentSelection.rowIndex,
          colIndex: currentSelection.colIndex,
        },
        'insert-row-below',
      );
      return result.changed;
    });
    if (!appendResult) {
      closePreviewTableCellEditor();
      return;
    }
    finalSelection = appendResult.selection;
  }

  closePreviewTableCellEditor();
  await nextTick();
  openPreviewTableCellEditor({
    ...finalSelection,
    rowIndex: targetRow,
    colIndex: targetCol,
  });
}

function handlePreviewTableCellEditBlur() {
  closePreviewTableCellEditorAndCommit();
}

function handlePreviewTableCellEditKeydown(event: KeyboardEvent) {
  if (event.ctrlKey && event.shiftKey && event.altKey && event.key === 'Enter') {
    event.preventDefault();
    void applyPreviewTableActionFromCellEditor('insert-col-right');
    return;
  }

  if (event.ctrlKey && event.shiftKey && event.altKey && (event.key === 'Delete' || event.key === 'Backspace')) {
    event.preventDefault();
    void applyPreviewTableActionFromCellEditor('delete-col');
    return;
  }

  if (event.ctrlKey && event.shiftKey && event.key === 'Enter') {
    event.preventDefault();
    void applyPreviewTableActionFromCellEditor('insert-row-below');
    return;
  }

  if (event.ctrlKey && event.shiftKey && (event.key === 'Delete' || event.key === 'Backspace')) {
    event.preventDefault();
    void applyPreviewTableActionFromCellEditor('delete-row');
    return;
  }

  if (event.ctrlKey && event.altKey && event.key === 'ArrowUp') {
    event.preventDefault();
    void applyPreviewTableActionFromCellEditor('move-row-up');
    return;
  }

  if (event.ctrlKey && event.altKey && event.key === 'ArrowDown') {
    event.preventDefault();
    void applyPreviewTableActionFromCellEditor('move-row-down');
    return;
  }

  if (event.ctrlKey && event.altKey && event.key === 'ArrowLeft') {
    event.preventDefault();
    void applyPreviewTableActionFromCellEditor('move-col-left');
    return;
  }

  if (event.ctrlKey && event.altKey && event.key === 'ArrowRight') {
    event.preventDefault();
    void applyPreviewTableActionFromCellEditor('move-col-right');
    return;
  }

  if (event.key === 'Escape') {
    event.preventDefault();
    closePreviewTableCellEditor();
    return;
  }

  if (event.key === 'Tab') {
    event.preventDefault();
    void movePreviewTableCellEditor(event.shiftKey ? 'tab-prev' : 'tab-next');
    return;
  }

  if (event.key === 'Enter' && !event.shiftKey && !event.ctrlKey && !event.metaKey) {
    event.preventDefault();
    void movePreviewTableCellEditor('enter-down');
  }
}

async function handleSourceContextMenu(event: MouseEvent) {
  if (!sourceMode.value) {
    return;
  }

  event.preventDefault();
  closePreviewTableMenu();
  closePreviewTableCellEditorAndCommit();
  closePreviewInlineEditor();
  editorView?.focus();

  const x = Number.isFinite(event.clientX) ? Math.round(event.clientX) : undefined;
  const y = Number.isFinite(event.clientY) ? Math.round(event.clientY) : undefined;
  const nativeResult = await runTask('source-editor-context-menu', (traceId) => {
    return bridgeService.showSourceEditorContextMenu({
      x,
      y,
      traceId,
    });
  });

  if (nativeResult?.handled) {
    closeSourceContextMenu();
    return;
  }

  const container = editorPreviewHost.value;
  if (!container) {
    closeSourceContextMenu();
    return;
  }

  const containerRect = container.getBoundingClientRect();
  sourceContextMenuStyle.value = {
    left: `${event.clientX - containerRect.left}px`,
    top: `${event.clientY - containerRect.top}px`,
  };
  sourceContextMenuVisible.value = true;
}

function scheduleTypewriterScroll(view: EditorView) {
  if (!typewriterMode.value) {
    return;
  }

  if (typewriterScrollFrame) {
    cancelAnimationFrame(typewriterScrollFrame);
  }

  typewriterScrollFrame = requestAnimationFrame(() => {
    typewriterScrollFrame = 0;
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
  if (!editorHost.value || editorView) {
    return;
  }

  const state = EditorState.create({
    doc: store.content,
    extensions: [
      history(),
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
        ...defaultKeymap,
        ...historyKeymap,
      ]),
      fontSizeCompartment.of(createFontSizeExtension(getEffectiveEditorFontSize())),
      EditorView.updateListener.of((update) => {
        if (update.selectionSet || update.docChanged) {
          syncCursorLineFromState(update.state);
          scheduleTypewriterScroll(update.view);
        }

        if (!update.docChanged || isSyncingFromStore) {
          return;
        }

        const nextContent = update.state.doc.toString();
        store.setContent(nextContent);
      }),
    ],
  });

  editorView = new EditorView({
    state,
    parent: editorHost.value,
  });

  syncCursorLineFromState(editorView.state);
}

function destroyEditor() {
  if (previewScrollSyncFrame) {
    cancelAnimationFrame(previewScrollSyncFrame);
    previewScrollSyncFrame = 0;
  }

  if (typewriterScrollFrame) {
    cancelAnimationFrame(typewriterScrollFrame);
    typewriterScrollFrame = 0;
  }

  if (!editorView) {
    return;
  }

  editorView.destroy();
  editorView = null;
}

function setTab(tab: Tab) {
  activeTab.value = tab;
}

function detectNativeWindowMode() {
  if (typeof navigator === 'undefined') {
    return false;
  }

  return /Electron/i.test(navigator.userAgent);
}

function togglePinOutline() {
  pinOutline.value = !pinOutline.value;
}

function toggleSourceMode() {
  const nextSourceMode = !sourceMode.value;
  if (nextSourceMode) {
    const previewAnchorLine = resolvePreviewViewportAnchorLine();
    if (previewAnchorLine > 0) {
      currentCursorLine.value = previewAnchorLine;
    }
  } else if (editorView) {
    syncCursorLineFromState(editorView.state);
  }

  sourceMode.value = nextSourceMode;
}

function toggleFocusMode() {
  focusMode.value = !focusMode.value;
}

function toggleTypewriterMode() {
  typewriterMode.value = !typewriterMode.value;

  if (!typewriterMode.value) {
    return;
  }

  nextTick(() => {
    if (!editorView || !sourceMode.value) {
      return;
    }

    editorView.focus();
    scheduleTypewriterScroll(editorView);
  });
}

function applySpellState(state: {
  enabled: boolean;
  language: string;
  availableLanguages: string[];
  dictionaryStatus: SpellDictionaryStatus;
}) {
  spellCheckEnabled.value = state.enabled;
  spellCheckLanguage.value = state.language;
  spellCheckLanguages.value = state.availableLanguages;
  spellDictionaryStatus.value = state.dictionaryStatus;
}

async function refreshSpellCheckState() {
  const result = await runTask('spell-get-state', (traceId) => {
    return bridgeService.getSpellCheckState(traceId);
  });

  if (!result) {
    return;
  }

  applySpellState(result);
}

function toggleSpellPanel() {
  showSpellPanel.value = !showSpellPanel.value;
}

async function changeSpellCheckEnabled() {
  const result = await runTask('spell-set-enabled', (traceId) => {
    return bridgeService.setSpellCheckEnabled(spellCheckEnabled.value, traceId);
  });

  if (!result) {
    return;
  }

  applySpellState(result);
}

async function changeSpellCheckLanguage() {
  const result = await runTask('spell-set-language', (traceId) => {
    return bridgeService.setSpellCheckLanguage(spellCheckLanguage.value, traceId);
  });

  if (!result) {
    return;
  }

  applySpellState(result);
}

function startSidebarResize(event: MouseEvent) {
  if (!pinOutline.value) {
    return;
  }

  event.preventDefault();

  const startX = event.clientX;
  const startWidth = sidebarWidth.value;
  isResizingSidebar.value = true;

  const onMouseMove = (moveEvent: MouseEvent) => {
    const deltaX = moveEvent.clientX - startX;
    sidebarWidth.value = normalizeSidebarWidth(startWidth + deltaX);
  };

  const stopResize = () => {
    isResizingSidebar.value = false;
    window.removeEventListener('mousemove', onMouseMove);
    window.removeEventListener('mouseup', stopResize);
    sidebarResizeCleanup = null;
  };

  window.addEventListener('mousemove', onMouseMove);
  window.addEventListener('mouseup', stopResize);

  sidebarResizeCleanup = stopResize;
}

function startMainPreviewResize(event: MouseEvent) {
  if (!sourceMode.value) {
    return;
  }

  event.preventDefault();

  const containerWidth = resolveMainPreviewContainerWidth();
  if (containerWidth <= 0) {
    return;
  }

  if (mainPreviewWidth.value <= 0) {
    ensureMainPreviewWidth();
  }

  const startX = event.clientX;
  const startWidth = mainPreviewWidth.value;
  isResizingMainPreview.value = true;

  const onMouseMove = (moveEvent: MouseEvent) => {
    const deltaX = moveEvent.clientX - startX;
    mainPreviewWidth.value = normalizeMainPreviewWidth(startWidth + deltaX, containerWidth);
  };

  const stopResize = () => {
    isResizingMainPreview.value = false;
    window.removeEventListener('mousemove', onMouseMove);
    window.removeEventListener('mouseup', stopResize);
    mainPreviewResizeCleanup = null;
  };

  window.addEventListener('mousemove', onMouseMove);
  window.addEventListener('mouseup', stopResize);
  mainPreviewResizeCleanup = stopResize;
}

function toggleTreeRow(row: FileTreeRow) {
  if (!row.expandable) {
    return;
  }

  fileTreeExpanded.value = {
    ...fileTreeExpanded.value,
    [row.id]: !row.expanded,
  };
}

function openFileFromList(path: string) {
  void openRecentFile(path);
}

function handleFileTreeContextMenu(event: MouseEvent, row: FileTreeRow) {
  if (row.type !== 'file' || !row.filePath) {
    return;
  }

  const filePath = row.filePath;
  const x = Number.isFinite(event.clientX) ? Math.round(event.clientX) : undefined;
  const y = Number.isFinite(event.clientY) ? Math.round(event.clientY) : undefined;

  void runTask('file-tree-context-menu', (traceId) => {
    return bridgeService.showFileTreeContextMenu({
      filePath,
      x,
      y,
      traceId,
    });
  });
}

function isFileDragEvent(event: DragEvent) {
  const types = event.dataTransfer?.types;
  if (!types) {
    return false;
  }

  return Array.from(types).includes('Files');
}

function parseFileUriToPath(value: string) {
  try {
    const parsed = new URL(value);
    if (parsed.protocol !== 'file:') {
      return '';
    }

    const decodedPath = decodeURIComponent(parsed.pathname || '');
    if (/^\/[A-Za-z]:\//.test(decodedPath)) {
      return decodedPath.slice(1).replace(/\//g, '\\');
    }

    return decodedPath.replace(/\//g, '\\');
  } catch {
    return '';
  }
}

function resolveDroppedMarkdownPath(event: DragEvent) {
  const transfer = event.dataTransfer;
  if (!transfer) {
    return '';
  }

  const droppedFiles = Array.from(transfer.files || []) as Array<File & { path?: string }>;
  for (const fileItem of droppedFiles) {
    let filePath = typeof fileItem.path === 'string' ? fileItem.path : '';
    if (!filePath) {
      filePath = bridgeService.getPathForFile(fileItem);
    }

    if (filePath && isMarkdownPath(filePath)) {
      return filePath;
    }
  }

  const uriList = transfer.getData('text/uri-list');
  if (!uriList) {
    return '';
  }

  const uriItems = uriList
    .split(/\r?\n/)
    .map((item) => item.trim())
    .filter((item) => item && !item.startsWith('#'));

  for (const uriItem of uriItems) {
    const filePath = parseFileUriToPath(uriItem);
    if (filePath && isMarkdownPath(filePath)) {
      return filePath;
    }
  }

  return '';
}

function handleWindowDragOver(event: DragEvent) {
  if (!isFileDragEvent(event)) {
    return;
  }

  event.preventDefault();
  isDragOverFile.value = true;
  if (event.dataTransfer) {
    event.dataTransfer.dropEffect = 'copy';
  }
}

function handleWindowDragEnter(event: DragEvent) {
  if (!isFileDragEvent(event)) {
    return;
  }

  event.preventDefault();
  fileDragDepth += 1;
  isDragOverFile.value = true;
}

function handleWindowDragLeave(event: DragEvent) {
  if (!isFileDragEvent(event)) {
    return;
  }

  event.preventDefault();
  fileDragDepth = Math.max(0, fileDragDepth - 1);
  if (fileDragDepth === 0) {
    isDragOverFile.value = false;
  }
}

function handleWindowDrop(event: DragEvent) {
  if (!isFileDragEvent(event)) {
    return;
  }

  event.preventDefault();
  fileDragDepth = 0;
  isDragOverFile.value = false;

  const filePath = resolveDroppedMarkdownPath(event);
  if (!filePath) {
    errorMessage.value = '仅支持拖拽打开 Markdown 文件（.md/.markdown）。';
    return;
  }

  void openRecentFile(filePath);
}

async function refreshCurrentDirectoryFiles(filePath: string) {
  if (!filePath) {
    currentDirectoryPath.value = '';
    currentDirectoryMarkdownFiles.value = [];
    return;
  }

  const fallbackPath = dedupeMarkdownPaths([filePath]);

  try {
    const traceId = createTraceId('list-dir');
    const result = await bridgeService.listMarkdownInDirectory(filePath, traceId);
    currentDirectoryPath.value = result.directoryPath;
    currentDirectoryMarkdownFiles.value = dedupeMarkdownPaths([
      ...result.filePaths,
      ...fallbackPath,
    ]);
  } catch (error) {
    currentDirectoryPath.value = getParentPath(filePath);
    currentDirectoryMarkdownFiles.value = fallbackPath;
    const message = error instanceof Error ? error.message : String(error);
    await logBridge('error', 'list-directory-markdown:fail', createTraceId('list-dir'), {
      message,
      filePath,
    });
  }
}

function openDocSearchPanel(replaceMode = false) {
  showDocSearchPanel.value = true;
  showDocReplacePanel.value = replaceMode;

  if (!replaceMode && docReplaceQuery.value && !docSearchQuery.value) {
    docSearchQuery.value = docReplaceQuery.value;
  }

  void nextTick(() => {
    const input = docSearchInput.value;
    if (!input) {
      return;
    }

    input.focus();
    input.select();
  });
}

function closeDocSearchPanel() {
  showDocSearchPanel.value = false;
  showDocReplacePanel.value = false;
}

function normalizeSearchIndex(index: number, length: number) {
  if (length <= 0) {
    return 0;
  }

  if (index < 0) {
    return length - 1;
  }

  if (index >= length) {
    return 0;
  }

  return index;
}

function resolveLineNumberByOffset(offset: number) {
  if (editorView) {
    return editorView.state.doc.lineAt(offset).number;
  }

  const safeOffset = Math.max(0, Math.min(offset, store.content.length));
  const prefix = store.content.slice(0, safeOffset);
  return prefix.split('\n').length;
}

function selectDocMatch(index: number, options?: { focusEditor?: boolean }) {
  const matches = docSearchMatches.value;
  if (!editorView || matches.length === 0) {
    return;
  }

  const normalized = normalizeSearchIndex(index, matches.length);
  docSearchMatchIndex.value = normalized;
  const target = matches[normalized];

  editorView.dispatch({
    selection: {
      anchor: target.from,
      head: target.to,
    },
    effects: EditorView.scrollIntoView(target.from, {
      y: 'center',
    }),
  });

  const lineNumber = resolveLineNumberByOffset(target.from);
  currentCursorLine.value = lineNumber;
  store.setCursorHint(lineNumber);
  scrollPreviewToLine(lineNumber);

  if (options?.focusEditor) {
    editorView.focus();
  }
}

function goToNextDocMatch() {
  selectDocMatch(docSearchMatchIndex.value + 1, {
    focusEditor: sourceMode.value && !showDocSearchPanel.value,
  });
}

function goToPrevDocMatch() {
  selectDocMatch(docSearchMatchIndex.value - 1, {
    focusEditor: sourceMode.value && !showDocSearchPanel.value,
  });
}

function replaceCurrentDocMatch() {
  const matches = docSearchMatches.value;
  if (matches.length === 0) {
    return;
  }

  const normalized = normalizeSearchIndex(docSearchMatchIndex.value, matches.length);
  const target = matches[normalized];
  const nextContent = `${store.content.slice(0, target.from)}${docReplaceQuery.value}${
    store.content.slice(target.to)
  }`;
  store.setContent(nextContent);

  nextTick(() => {
    if (docSearchMatches.value.length > 0) {
      selectDocMatch(normalized, {
        focusEditor: false,
      });
    }
  });
}

function replaceAllDocMatch() {
  const expression = buildSearchRegExp(
    docSearchQuery.value,
    docSearchCaseSensitive.value,
    docSearchWholeWord.value,
    true,
  );

  if (!expression || docSearchMatches.value.length === 0) {
    return;
  }

  const nextContent = store.content.replace(expression, docReplaceQuery.value);
  store.setContent(nextContent);

  nextTick(() => {
    docSearchMatchIndex.value = 0;
    if (docSearchMatches.value.length > 0) {
      selectDocMatch(0, {
        focusEditor: false,
      });
    }
  });
}

function handleDocSearchEnter(isShift: boolean) {
  if (isShift) {
    goToPrevDocMatch();
    return;
  }

  goToNextDocMatch();
}

function handleGlobalPointerDown(event: MouseEvent) {
  if (!previewTableMenuVisible.value && !sourceContextMenuVisible.value) {
    return;
  }

  const target = event.target as HTMLElement | null;
  if (!target) {
    closePreviewTableMenu();
    closeSourceContextMenu();
    return;
  }

  if (previewTableMenuVisible.value && target.closest('.preview-table-menu')) {
    return;
  }

  if (sourceContextMenuVisible.value && target.closest('.source-context-menu')) {
    return;
  }

  closePreviewTableMenu();
  closeSourceContextMenu();
}

function handleGlobalWheel(event: WheelEvent) {
  const target = event.target instanceof HTMLElement ? event.target : null;
  const isZoomGesture = event.ctrlKey || event.metaKey;
  if (isZoomGesture && target && editorPreviewHost.value?.contains(target)) {
    event.preventDefault();
    stepContentZoom(event.deltaY < 0 ? 'in' : 'out');
    return;
  }

  if (!previewInlineEditVisible.value && !previewTableCellEditVisible.value) {
    return;
  }

  if (!target?.closest('.preview-inline-editor') && !target?.closest('.preview-table-cell-editor')) {
    return;
  }

  const preview = previewHost.value;
  if (!preview) {
    return;
  }

  event.preventDefault();
  preview.scrollTop += event.deltaY;
}

function handleGlobalKeydown(event: KeyboardEvent) {
  if (event.key === 'Escape') {
    if (previewTableMenuVisible.value) {
      event.preventDefault();
      closePreviewTableMenu();
      return;
    }

    if (sourceContextMenuVisible.value) {
      event.preventDefault();
      closeSourceContextMenu();
      return;
    }

    if (showDocSearchPanel.value) {
      event.preventDefault();
      closeDocSearchPanel();
      return;
    }
  }

  if (!sourceContextMenuVisible.value && event.key === 'Escape' && previewInlineEditVisible.value) {
    event.preventDefault();
    cancelPreviewInlineEdit();
    return;
  }

  if (!sourceContextMenuVisible.value && event.key === 'Escape' && previewTableCellEditVisible.value) {
    event.preventDefault();
    closePreviewTableCellEditor();
    return;
  }

  if (!(event.ctrlKey || event.metaKey)) {
    if (event.key === 'F3') {
      event.preventDefault();
      goToNextDocMatch();
    }
    return;
  }

  if (event.altKey && event.key.toLowerCase() === 'f') {
    event.preventDefault();
    toggleFocusMode();
    return;
  }

  if (event.altKey && event.key.toLowerCase() === 't') {
    event.preventDefault();
    toggleTypewriterMode();
    return;
  }

  if (event.altKey && event.key.toLowerCase() === 'l') {
    event.preventDefault();
    toggleSpellPanel();
    return;
  }

  if (event.key.toLowerCase() === 'f') {
    event.preventDefault();
    openDocSearchPanel(false);
    return;
  }

  if (event.key.toLowerCase() === 'h') {
    event.preventDefault();
    openDocSearchPanel(true);
    return;
  }

  if (event.key.toLowerCase() === 'o') {
    event.preventDefault();
    void openFile();
    return;
  }

  if (event.key.toLowerCase() === 's') {
    event.preventDefault();
    if (event.shiftKey) {
      void saveAsFile();
      return;
    }

    void saveFile();
    return;
  }

  if (event.key.toLowerCase() === 'e') {
    event.preventDefault();
    void exportHtml();
  }
}

function jumpToHeading(line: number) {
  currentCursorLine.value = line;
  store.setCursorHint(line);
  scrollEditorToLine(line);
  scrollPreviewToLine(line);
}

async function runTask<T>(
  taskName: string,
  executor: (traceId: string) => Promise<T>,
  options?: {
    fallbackHint?: string;
  },
): Promise<T | null> {
  const traceId = createTraceId(taskName);

  try {
    await logBridge('info', `${taskName}:start`, traceId);
    const result = await executor(traceId);
    await logBridge('info', `${taskName}:ok`, traceId);
    errorMessage.value = '';

    return result;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const fallbackText = options?.fallbackHint
      ? `；建议：${options.fallbackHint}`
      : '';
    errorMessage.value = `${taskName} 失败：${message}（追踪ID: ${traceId}）${fallbackText}`;
    await logBridge('error', `${taskName}:fail`, traceId, {
      message,
      fallbackHint: options?.fallbackHint || '',
    });

    return null;
  }
}

async function runDocumentLoadingTask<T>(
  taskName: string,
  executor: (traceId: string) => Promise<T>,
  options?: {
    fallbackHint?: string;
  },
): Promise<T | null> {
  isDocumentLoading.value = true;
  await nextTick();

  try {
    return await runTask(taskName, executor, options);
  } finally {
    isDocumentLoading.value = false;
  }
}

async function openFile() {
  const result = await runDocumentLoadingTask('open-file', (traceId) => {
    return bridgeService.openMarkdown(traceId);
  });

  if (!result) {
    return;
  }

  store.setFile(result.filePath, result.content);
  pushRecentFile(result.filePath);
  activeTab.value = 'outline';
}

async function saveFile() {
  const result = await runTask('save-file', (traceId) => {
    return bridgeService.saveMarkdown({
      filePath: store.filePath || undefined,
      content: store.content,
      traceId,
    });
  });

  if (!result) {
    return;
  }

  store.markSaved(result.filePath);
  pushRecentFile(result.filePath);
}

async function saveAsFile() {
  const result = await runTask('save-as-file', (traceId) => {
    return bridgeService.saveMarkdown({
      content: store.content,
      traceId,
    });
  });

  if (!result) {
    return;
  }

  store.markSaved(result.filePath);
  pushRecentFile(result.filePath);
}

async function openRecentFile(filePath: string) {
  const result = await runDocumentLoadingTask('open-recent-file', (traceId) => {
    return bridgeService.openFromPath(filePath, traceId);
  });

  if (!result) {
    return;
  }

  store.setFile(result.filePath, result.content);
  pushRecentFile(result.filePath);
  activeTab.value = 'outline';
}

async function tryOpenLaunchFileFromNative() {
  if (!isNativeWindow.value) {
    return;
  }

  const result = await runTask('consume-launch-file', (traceId) => {
    return bridgeService.consumeLaunchFilePath(traceId);
  });
  const filePath = result?.filePath || '';
  if (!filePath) {
    return;
  }

  await openRecentFile(filePath);
}

async function handleNativeLaunchFile(filePath: string) {
  if (!filePath) {
    return;
  }

  await openRecentFile(filePath);
}

function buildExportHtmlDocument() {
  const titleText = store.filePath || 'UniReadMD Export';

  return `<!doctype html>
<html lang='zh-CN'>
  <head>
    <meta charset='UTF-8' />
    <meta name='viewport' content='width=device-width, initial-scale=1.0' />
    <title>${titleText}</title>
  </head>
  <body>
    ${renderedHtml.value}
  </body>
</html>`;
}

function getExportFallbackHint(format: ExportFormat) {
  if (format === 'docx' || format === 'odt' || format === 'rtf') {
    return '请确认已安装 Pandoc，或先导出 HTML/PDF/Image。';
  }

  if (format === 'pdf' || format === 'png' || format === 'jpeg') {
    return '可先导出 HTML，再通过系统打印或截图作为临时回退。';
  }

  return '可检查目标文件路径权限，或改用“另存为”后重试。';
}

function buildExportMarkdownDocument() {
  return store.content.replace(/\r\n/g, '\n');
}

async function refreshExportCapabilities() {
  const result = await runTask('export-capabilities', (traceId) => {
    return bridgeService.getExportCapabilities(traceId);
  }, {
    fallbackHint: '导出能力探测失败时将仅保留 HTML 导出。',
  });

  if (result) {
    exportCapabilities.value = result;
    return;
  }

  exportCapabilities.value = {
    builtinFormats: [...BUILTIN_EXPORT_FORMATS],
    pandocAvailable: false,
    pandocVersion: '',
    pandocFormats: [],
  };
}

async function exportDocumentByFormat(format: ExportFormat) {
  if (!availableExportFormats.value.has(format)) {
    const taskName = `export-${format}`;
    const traceId = createTraceId(taskName);
    const hint = getExportFallbackHint(format);
    errorMessage.value = `当前环境暂不支持导出 ${EXPORT_LABELS[format]}（追踪ID: ${traceId}）；建议：${hint}`;
    await logBridge('error', `${taskName}:unsupported`, traceId, {
      format,
      capability: exportCapabilities.value,
      hint,
    });
    return;
  }

  await runTask(`export-${format}`, (traceId) => {
    return bridgeService.exportDocument({
      format,
      html: buildExportHtmlDocument(),
      markdown: buildExportMarkdownDocument(),
      sourceFilePath: store.filePath || undefined,
      traceId,
    });
  }, {
    fallbackHint: getExportFallbackHint(format),
  });
}

async function exportHtml() {
  await exportDocumentByFormat('html');
}

async function exportPdf() {
  await exportDocumentByFormat('pdf');
}

async function exportPng() {
  await exportDocumentByFormat('png');
}

async function exportJpeg() {
  await exportDocumentByFormat('jpeg');
}

async function exportDocx() {
  await exportDocumentByFormat('docx');
}

async function changeTheme() {
  applyTheme(theme.value);

  await runTask('settings-theme', (traceId) => {
    return bridgeService.setSetting('theme', theme.value, traceId);
  });

  await nextTick();
  await renderMermaidBlocks();
}

async function minimizeWindow() {
  await runTask('window-minimize', (traceId) => bridgeService.minimizeWindow(traceId));
}

async function toggleMaximizeWindow() {
  await runTask('window-toggle-maximize', (traceId) => {
    return bridgeService.toggleMaximizeWindow(traceId);
  });
}

async function closeWindow() {
  await runTask('window-close', (traceId) => bridgeService.closeWindow(traceId));
}

async function handleMenuCommand(command: string) {
  if (command === 'file-open') {
    await openFile();
    return;
  }

  if (command === 'file-save') {
    await saveFile();
    return;
  }

  if (command === 'file-save-as') {
    await saveAsFile();
    return;
  }

  if (command === 'file-export-html') {
    await exportHtml();
    return;
  }

  if (command === 'file-export-pdf') {
    await exportPdf();
    return;
  }

  if (command === 'file-export-png') {
    await exportPng();
    return;
  }

  if (command === 'file-export-jpeg') {
    await exportJpeg();
    return;
  }

  if (command === 'file-export-docx') {
    await exportDocx();
    return;
  }

  if (command === 'edit-search') {
    openDocSearchPanel(false);
    return;
  }

  if (command === 'edit-focus') {
    toggleFocusMode();
    return;
  }

  if (command === 'edit-typewriter') {
    toggleTypewriterMode();
    return;
  }

  if (command === 'help-spell') {
    toggleSpellPanel();
    return;
  }

  if (command === 'theme-light' || command === 'theme-dark' || command === 'theme-system') {
    const nextTheme = command.replace('theme-', '') as Theme;
    if (theme.value === nextTheme) {
      return;
    }

    theme.value = nextTheme;
    await changeTheme();
  }
}

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
    if (!editorView) {
      return;
    }

    editorView.dispatch({
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
        if (typewriterMode.value && editorView) {
          scheduleTypewriterScroll(editorView);
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

watch(
  () => [
    docSearchQuery.value,
    docSearchCaseSensitive.value,
    docSearchWholeWord.value,
    store.content,
  ],
  () => {
    docSearchMatchIndex.value = 0;

    if (docSearchMatches.value.length > 0 && showDocSearchPanel.value) {
      nextTick(() => {
        selectDocMatch(0, {
          focusEditor: false,
        });
      });
    }
  },
);

watch(
  () => showDocSearchPanel.value,
  (visible) => {
    if (visible && docSearchMatches.value.length > 0) {
      nextTick(() => {
        selectDocMatch(docSearchMatchIndex.value, {
          focusEditor: false,
        });
      });
    }
  },
);

watch(
  () => windowTitle.value,
  (nextTitle) => {
    document.title = nextTitle;
  },
  {
    immediate: true,
  },
);

watch(
  () => [
    store.filePath,
    store.content,
    store.isDirty,
    store.cursorHint,
    recentFiles.value.join('\n'),
    pinOutline.value,
    sourceMode.value,
    focusMode.value,
    typewriterMode.value,
    contentZoom.value,
    sidebarWidth.value,
    mainPreviewWidth.value,
    fileViewMode.value,
    fileSortMode.value,
    filesFilter.value,
    showDocSearchPanel.value,
    showDocReplacePanel.value,
    docSearchQuery.value,
    docReplaceQuery.value,
    docSearchCaseSensitive.value,
    docSearchWholeWord.value,
    docSearchMatchIndex.value,
    JSON.stringify(fileTreeExpanded.value),
  ],
  () => {
    persistSessionSnapshot();
  },
);

onMounted(async () => {
  Object.defineProperty(window, '__unireadmdMounted', {
    configurable: true,
    writable: true,
    value: true,
  });

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
  menuCommandCleanup = window.nativeBridge.menu?.onCommand((command) => {
    void handleMenuCommand(command);
  }) || null;
  launchFileCleanup = bridgeService.onLaunchFile((payload) => {
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
  if (sidebarResizeCleanup) {
    sidebarResizeCleanup();
  }
  if (mainPreviewResizeCleanup) {
    mainPreviewResizeCleanup();
  }
  if (menuCommandCleanup) {
    menuCommandCleanup();
  }
  if (launchFileCleanup) {
    launchFileCleanup();
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
</script>

<template>
  <div
    class='layout'
    :class='layoutClassName'
    :style='layoutStyle'
    data-testid='app-layout'
    @dragenter='handleWindowDragEnter'
    @dragleave='handleWindowDragLeave'
    @dragover='handleWindowDragOver'
    @drop='handleWindowDrop'
  >
    <header class='top-titlebar'>
      <div class='title-left'>
        <button type='button' class='menu-icon' @click='setTab("files")'>
          <span></span>
          <span></span>
          <span></span>
        </button>
        <span class='title-name'>{{ windowTitle }}</span>
        <span v-if='store.isDirty' class='dirty-dot'>●</span>
      </div>
      <div class='title-actions'></div>
      <div class='window-actions'>
        <button type='button' aria-label='最小化' @click='minimizeWindow'>_</button>
        <button type='button' aria-label='最大化' @click='toggleMaximizeWindow'>□</button>
        <button type='button' aria-label='关闭' @click='closeWindow'>×</button>
      </div>
    </header>

    <div class='layout-body'>
      <aside class='sidebar' data-testid='sidebar'>
        <div class='sidebar-search-top'>
          <input
            v-model='sidebarSearchText'
            class='search-input'
            data-testid='sidebar-search-input'
            :placeholder='sidebarSearchPlaceholder'
          />
        </div>
        <header class='sidebar-header'>
          <button
            type='button'
            data-testid='tab-files'
            :class='{ active: activeTab === "files" }'
            @click='setTab("files")'
          >
            文件
          </button>
          <button
            type='button'
            data-testid='tab-outline'
            :class='{ active: activeTab === "outline" }'
            @click='setTab("outline")'
          >
            大纲
          </button>
        </header>

        <section v-if='activeTab === "files"' class='panel panel-files'>
          <div class='files-workbench'>
            <div class='files-tree' data-testid='files-tree-view'>
              <div
                v-for='row in fileTreeRows'
                :key='row.id'
                class='files-tree-row'
                :class='{ group: row.type === "dir" && row.depth === 0 }'
                :style='{ paddingLeft: `${row.depth * 16 + 8}px` }'
              >
                <button
                  type='button'
                  class='files-tree-toggle'
                  :class='{
                    group: row.type === "dir" && row.depth === 0,
                    hidden: !row.expandable,
                  }'
                  @click='toggleTreeRow(row)'
                >
                  {{ row.expanded ? '▾' : '▸' }}
                </button>
                <button
                  type='button'
                  class='files-tree-entry'
                  :title='row.label'
                  :class='{
                    dir: row.type === "dir",
                    file: row.type === "file",
                    group: row.type === "dir" && row.depth === 0,
                    active: row.filePath === store.filePath,
                  }'
                  @contextmenu.prevent='handleFileTreeContextMenu($event, row)'
                  @click='
                    row.type === "file" && row.filePath
                      ? openFileFromList(row.filePath)
                      : toggleTreeRow(row)
                  '
                >
                  {{ row.label }}
                </button>
              </div>
              <div v-if='fileTreeRows.length === 0' class='empty'>
                {{ filesFilter.trim().length > 0 ? '未找到匹配文档' : '暂无 Markdown 文件' }}
              </div>
            </div>
          </div>
        </section>

        <section v-else-if='activeTab === "outline"' class='panel'>
          <div v-if='outlineDisplayItems.length === 0' class='empty'>
            {{ searchText.trim().length > 0 ? '未找到匹配标题' : '当前文档无标题' }}
          </div>
          <button
            v-for='item in outlineDisplayItems'
            :key='`${item.line}-${item.text}`'
            type='button'
            data-testid='outline-item'
            class='outline-item'
            :class='[
              `outline-level-${item.level}`,
              { active: item.line === activeHeadingLine },
            ]'
            :style='{ paddingLeft: `${(item.level - 1) * 14 + 10}px` }'
            @click='jumpToHeading(item.line)'
          >
            <span class='outline-item-text'>{{ item.text }}</span>
          </button>
        </section>
        <footer class='sidebar-footer' aria-label='sidebar-footer'></footer>
      </aside>
      <div
        class='sidebar-resizer'
        data-testid='sidebar-resizer'
        role='separator'
        aria-orientation='vertical'
        @mousedown='startSidebarResize'
      ></div>

      <main class='main'>
        <section v-if='showDocSearchPanel' class='doc-search-panel' data-testid='doc-search-panel'>
          <div class='doc-search-row'>
            <input
              ref='docSearchInput'
              v-model='docSearchQuery'
              class='doc-search-input'
              data-testid='doc-search-query'
              placeholder='Search text'
              @keydown.enter.prevent='handleDocSearchEnter($event.shiftKey)'
            />
            <button
              type='button'
              class='doc-search-option'
              :class='{ active: docSearchCaseSensitive }'
              @click='docSearchCaseSensitive = !docSearchCaseSensitive'
            >
              Aa
            </button>
            <button
              type='button'
              class='doc-search-option'
              :class='{ active: docSearchWholeWord }'
              @click='docSearchWholeWord = !docSearchWholeWord'
            >
              W
            </button>
            <button type='button' class='doc-search-btn' @click='goToPrevDocMatch'>↑</button>
            <button type='button' class='doc-search-btn' @click='goToNextDocMatch'>↓</button>
            <button
              v-if='!showDocReplacePanel'
              type='button'
              class='doc-search-btn'
              data-testid='doc-open-replace'
              @click='showDocReplacePanel = true'
            >
              Replace
            </button>
            <button
              v-else
              type='button'
              class='doc-search-btn'
              data-testid='doc-close-replace'
              @click='showDocReplacePanel = false'
            >
              Search Only
            </button>
            <button
              type='button'
              class='doc-search-btn close'
              @click='closeDocSearchPanel'
            >
              ✕
            </button>
          </div>
          <div v-if='showDocReplacePanel' class='doc-search-row'>
            <input
              v-model='docReplaceQuery'
              class='doc-search-input'
              data-testid='doc-replace-query'
              placeholder='Replace with'
              @keydown.enter.prevent='replaceCurrentDocMatch'
            />
            <button type='button' class='doc-search-btn' @click='replaceCurrentDocMatch'>
              Replace
            </button>
            <button type='button' class='doc-search-btn' @click='replaceAllDocMatch'>
              Replace All
            </button>
          </div>
          <div class='doc-search-summary' data-testid='doc-search-summary'>
            {{ docSearchSummary }}
          </div>
        </section>
        <div v-if='errorMessage' class='error-banner'>{{ errorMessage }}</div>
        <section
          ref='editorPreviewHost'
          class='editor-preview'
          data-testid='editor-preview'
        >
          <section
            v-if='isDocumentLoading'
            class='document-loading-overlay'
            data-testid='document-loading-overlay'
          >
            <div class='document-loading-card'>加载中...</div>
          </section>
          <article
            ref='previewHost'
            id='write'
            class='preview'
            data-testid='preview-host'
            v-html='renderedHtml'
            @click='handlePreviewClick'
            @dblclick='handlePreviewDoubleClick'
            @scroll='handlePreviewScroll'
            @contextmenu='handlePreviewContextMenu'
          ></article>
          <section
            v-if='previewInlineEditVisible'
            class='preview-inline-editor'
            :style='previewInlineEditStyle'
            data-testid='preview-inline-editor'
          >
            <textarea
              ref='previewInlineEditTextarea'
              v-model='previewInlineEditValue'
              class='preview-inline-editor-input'
              @input='handlePreviewInlineEditInput'
              @compositionstart='handlePreviewInlineEditCompositionStart'
              @compositionend='handlePreviewInlineEditCompositionEnd'
              @blur='handlePreviewInlineEditBlur'
              @keydown='handlePreviewInlineEditKeydown'
              @wheel='handlePreviewInlineEditorWheel'
            ></textarea>
          </section>
          <section
            v-if='previewTableCellEditVisible'
            class='preview-table-cell-editor'
            :style='previewTableCellEditStyle'
            data-testid='preview-table-cell-editor'
          >
            <input
              ref='previewTableCellEditInput'
              v-model='previewTableCellEditValue'
              class='preview-table-cell-editor-input'
              data-testid='preview-table-cell-editor-input'
              @blur='handlePreviewTableCellEditBlur'
              @keydown='handlePreviewTableCellEditKeydown'
            />
          </section>
          <section
            v-if='previewTableMenuVisible'
            class='preview-table-menu'
            :style='previewTableMenuStyle'
            data-testid='preview-table-menu'
          >
            <button type='button' @click='applyPreviewTableAction("insert-row-above")'>
              在上方插入行
            </button>
            <button type='button' @click='applyPreviewTableAction("insert-row-below")'>
              在下方插入行
            </button>
            <button type='button' @click='applyPreviewTableAction("delete-row")'>删除当前行</button>
            <button type='button' @click='applyPreviewTableAction("move-row-up")'>上移当前行</button>
            <button type='button' @click='applyPreviewTableAction("move-row-down")'>下移当前行</button>
            <div class='preview-table-menu-divider'></div>
            <button type='button' @click='applyPreviewTableAction("insert-col-left")'>
              在左侧插入列
            </button>
            <button type='button' @click='applyPreviewTableAction("insert-col-right")'>
              在右侧插入列
            </button>
            <button type='button' @click='applyPreviewTableAction("delete-col")'>删除当前列</button>
            <button type='button' @click='applyPreviewTableAction("move-col-left")'>左移当前列</button>
            <button type='button' @click='applyPreviewTableAction("move-col-right")'>右移当前列</button>
          </section>
          <section
            v-if='sourceContextMenuVisible'
            class='source-context-menu'
            :style='sourceContextMenuStyle'
            data-testid='source-context-menu'
          >
            <button type='button' @click='handleSourceContextAction("cut")'>剪切</button>
            <button type='button' @click='handleSourceContextAction("copy")'>复制</button>
            <button type='button' @click='handleSourceContextAction("paste")'>粘贴</button>
            <div class='source-context-menu-divider'></div>
            <button type='button' @click='handleSourceContextAction("select-all")'>全选</button>
            <button type='button' @click='handleSourceContextAction("search")'>查找</button>
          </section>
          <div
            class='main-preview-resizer'
            data-testid='main-preview-resizer'
            role='separator'
            aria-orientation='vertical'
            @mousedown='startMainPreviewResize'
          ></div>
          <div
            ref='editorHost'
            id='typora-source'
            class='editor-host'
            data-testid='editor-host'
            @contextmenu='handleSourceContextMenu'
          ></div>
        </section>
      </main>
    </div>

    <footer class='status-bar'>
      <div class='status-left'>
        <button
          id='outline-btn'
          type='button'
          class='status-action'
          :class='{ active: pinOutline }'
          @click='togglePinOutline'
        >
          {{ pinOutline ? '隐藏侧栏' : '显示侧栏' }}
        </button>
        <button
          id='toggle-sourceview-btn'
          type='button'
          class='status-action'
          :class='{ active: sourceMode }'
          @click='toggleSourceMode'
        >
          {{ sourceToggleLabel }}
        </button>
      </div>
      <div class='status-right'>
        <button
          id='footer-content-zoom'
          type='button'
          class='status-action status-zoom'
          title='按住 Ctrl + 鼠标滚轮缩放，单击重置为 100%'
          @click='resetContentZoom'
        >
          {{ contentZoomLabel }}
        </button>
        <button id='footer-word-count' type='button' class='status-action status-word'>
          {{ statusWordLabel }}
        </button>
        <span class='status-metric'>{{ contentReadMinutes }} min</span>
        <span class='status-metric'>{{ contentCharCount }} chars</span>
        <span class='status-metric'>{{ statusLineLabel }}</span>
      </div>
    </footer>
    <div
      v-if='showSpellPanel'
      class='spell-panel spell-panel-global'
      data-testid='spell-panel'
    >
      <label class='spell-option'>
        <input
          v-model='spellCheckEnabled'
          type='checkbox'
          @change='changeSpellCheckEnabled'
        />
        Enable
      </label>
      <label class='spell-option'>
        Language
        <select v-model='spellCheckLanguage' @change='changeSpellCheckLanguage'>
          <option
            v-for='item in spellCheckLanguages'
            :key='item'
            :value='item'
          >
            {{ item }}
          </option>
        </select>
      </label>
      <div class='spell-dictionary-status' data-testid='spell-dictionary-status'>
        {{ spellDictionaryLabel }}
      </div>
    </div>
    <div v-if='isDragOverFile' class='drag-drop-overlay' aria-hidden='true'>
      <div class='drag-drop-overlay-card'>
        <div class='drag-drop-overlay-title'>松开以打开 Markdown</div>
        <div class='drag-drop-overlay-subtitle'>支持 .md / .markdown 文件</div>
      </div>
    </div>
  </div>
</template>
