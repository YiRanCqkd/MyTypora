import type { useMarkdownFeature } from '../markdown/markdown-feature';
import type { useSettingsFoundation } from '../settings/settings-foundation';
import type { useWorkspaceFoundation } from '../workspace/workspace-foundation';
import { useDocumentSearchMatches } from './document-search-matches';
import { useDocumentStatistics } from './document-statistics';
import { usePreviewEditingLabel } from './preview-editing-label';
import { usePreviewEditingState } from './preview-editing-state';
import { createPreviewSearchDom } from './preview-search-dom';
import { usePreviewSearchHighlighting } from './preview-search-highlighting';
import { useStatusLine } from './status-line';
import type { useWorkbenchFoundation } from './workbench-foundation';

type EditorPresentationDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  markdown: ReturnType<typeof useMarkdownFeature>;
  settings: ReturnType<typeof useSettingsFoundation>;
  workspace: ReturnType<typeof useWorkspaceFoundation>;
}>;

export function useEditorPresentation(input: EditorPresentationDependencies) {
  const preview = usePreviewEditingState({ ...input.editor, ...input.markdown });
  const statistics = useDocumentStatistics(input.editor);
  const status = useStatusLine(input.editor);
  const editLabel = usePreviewEditingLabel({ ...input.markdown, ...preview });
  const matches = useDocumentSearchMatches(input.editor);
  const searchDom = createPreviewSearchDom(input.editor);
  const highlights = usePreviewSearchHighlighting({
    ...matches,
    ...searchDom,
    ...input.editor,
  });

  return {
    ...preview,
    ...statistics,
    ...status,
    ...editLabel,
    ...matches,
    ...searchDom,
    ...highlights,
  };
}
