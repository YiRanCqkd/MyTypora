import { computed } from 'vue';
import { highlightPreviewCode } from './preview-code-presentation';
import type { createMarkdownRenderer } from '../markdown/renderer';
import type { usePreviewEditingState } from './preview-editing-state';
type UsePreviewEditingLabelDependencies =
  ReturnType<typeof createMarkdownRenderer> &
  ReturnType<typeof usePreviewEditingState>;

export function usePreviewEditingLabel(input: UsePreviewEditingLabelDependencies) {
  const { PREVIEW_TOKEN_LABELS,previewInlineEditSelection,
  } = input;

  const previewInlineEditTitle = computed(() => {
    const selection = previewInlineEditSelection.value;
    if (!selection) {
      return '渲染态编辑';
    }

    return PREVIEW_TOKEN_LABELS[selection.token] || '渲染块';
  });

  const previewInlineEditCode = computed(() => ['fence', 'code_block']
    .includes(previewInlineEditSelection.value?.token || ''));
  const previewInlineEditHighlight = computed(() => highlightPreviewCode(
    input.previewInlineEditValue.value, previewInlineEditSelection.value?.token || '',
  ));

  return { previewInlineEditTitle, previewInlineEditCode, previewInlineEditHighlight,
  };
}
