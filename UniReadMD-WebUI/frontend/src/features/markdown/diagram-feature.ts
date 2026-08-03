import type { useEditorPresentation } from '../editor/editor-presentation';
import type { useWorkbenchFoundation } from '../editor/workbench-foundation';
import type { useSettingsFoundation } from '../settings/settings-foundation';
import { useDiagramBlocks } from './diagram-blocks';
import { createDiagramRenderer } from './diagram-renderer';
import type { useMarkdownFeature } from './markdown-feature';

type DiagramFeatureDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
  editorPresentation: ReturnType<typeof useEditorPresentation>;
  markdown: ReturnType<typeof useMarkdownFeature>;
  settings: ReturnType<typeof useSettingsFoundation>;
}>;

export function useDiagramFeature(input: DiagramFeatureDependencies) {
  const renderer = createDiagramRenderer({
    ...input.editor,
    ...input.editorPresentation,
    ...input.markdown,
  });
  const blocks = useDiagramBlocks({
    ...input.editor,
    ...input.markdown,
    ...input.settings,
    ...renderer,
  });

  return {
    ...renderer,
    ...blocks,
  };
}
