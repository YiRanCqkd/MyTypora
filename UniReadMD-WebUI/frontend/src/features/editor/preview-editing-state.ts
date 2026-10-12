import type {
  PreviewInlineEditSelection,
  PreviewInlineEditSession,
  PreviewTableSelection,
} from './editor-contracts';
import { Compartment } from '@codemirror/state';
import { extractHeadings, findActiveHeadingLine } from '../../utils/markdown';
import { ref, computed } from 'vue';
import type { useCursorPosition } from './cursor-position';
import type { createMarkdownRenderer } from '../markdown/renderer';
import type { createWorkbenchRuntime } from './workbench-runtime';
import type { useWorkbenchStore } from './editor-store';
import { sanitizeMarkdownHtml } from '../markdown/safe-html';
import { createPlainRenderCache } from '../markdown/plain-render-cache';
type UsePreviewEditingStateDependencies =
  ReturnType<typeof useCursorPosition> &
  ReturnType<typeof createMarkdownRenderer> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof useWorkbenchStore>;

export function usePreviewEditingState(input: UsePreviewEditingStateDependencies) {
  const { currentCursorLine,markdown,runtime,store,
  } = input;

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

  const mathPluginRevision = ref(0);
  const renderPlain = createPlainRenderCache(markdown);
  const previewVisualEditing = ref(false);
  const previewRenderRevision = ref(0);
  let cachedHtml = '';

  const renderedHtml = computed(() => {
    const revision = previewRenderRevision.value;
    if (previewVisualEditing.value && cachedHtml) return cachedHtml;
    void mathPluginRevision.value;
    void runtime.latestFrontMatter;
    const normalizedContent = store.content.replace(/^\uFEFF/, '');
    const normalized = normalizedContent.replace(/^\[toc\]$/gim, '[[toc]]');
    if (!normalized.trim()) {
      const endLine = normalized.split('\n').length;
      cachedHtml = '<p class="preview-empty-line" data-source-start="1" '
        + `data-source-end="${endLine}" data-source-token="empty_document"><br></p>`
        + `<!-- preview:${revision} -->`;
      return cachedHtml;
    }
    const plainHtml = renderPlain(normalized);
    cachedHtml = (plainHtml ?? sanitizeMarkdownHtml(markdown.render(normalized), {
      resourceBaseUrl: store.resourceBaseUrl,
    })) + `<!-- preview:${revision} -->`;
    return cachedHtml;
  });

  const headings = computed(() => {
    return extractHeadings(store.content);
  });

  const activeHeadingLine = computed(() => {
    return findActiveHeadingLine(headings.value, currentCursorLine.value);
  });

  return { previewInlineEditVisible,previewInlineEditValue,previewInlineEditStyle,
    previewInlineEditSelection,previewInlineEditSession,previewInlineEditComposing,
    previewInlineEditTextarea,previewTableMenuVisible,previewTableMenuStyle,
    previewTableSelection,previewTableCellEditVisible,previewTableCellEditValue,
    previewTableCellEditStyle,previewTableCellEditSelection,previewTableCellEditInput,
    sourceContextMenuVisible,sourceContextMenuStyle,fontSizeCompartment,mathPluginRevision,
    renderedHtml,headings,activeHeadingLine,previewVisualEditing,previewRenderRevision,
  };
}
