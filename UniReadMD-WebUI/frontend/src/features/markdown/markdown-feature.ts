import type { useWorkbenchFoundation } from '../editor/workbench-foundation';
import { createMarkdownRenderer } from './renderer';
import { createMarkdownSourceMapping } from './source-mapping';

type MarkdownFeatureDependencies = Readonly<{
  editor: ReturnType<typeof useWorkbenchFoundation>;
}>;

export function useMarkdownFeature(input: MarkdownFeatureDependencies) {
  const renderer = createMarkdownRenderer(input.editor);
  const sourceMapping = createMarkdownSourceMapping(renderer);

  return {
    ...renderer,
    ...sourceMapping,
  };
}
