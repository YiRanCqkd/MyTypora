import type { useMarkdownFeature } from '../markdown/markdown-feature';
import { useAppearance } from './appearance';
import { usePreferences } from './preferences';
import { useStatusLabels } from './status-labels';
import { useViewOptions } from './view-options';
import { useZoomLabel } from './zoom-label';

type SettingsFoundationDependencies = Readonly<{
  markdown: ReturnType<typeof useMarkdownFeature>;
}>;

export function useSettingsFoundation(input: SettingsFoundationDependencies) {
  const preferences = usePreferences(input.markdown);
  const view = useViewOptions();
  const zoom = useZoomLabel(preferences);
  const status = useStatusLabels(view);
  const appearance = useAppearance();

  return {
    ...preferences,
    ...view,
    ...zoom,
    ...status,
    ...appearance,
  };
}
