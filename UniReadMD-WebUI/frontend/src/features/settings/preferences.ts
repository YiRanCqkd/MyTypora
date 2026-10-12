import type { Theme } from './preferences-contracts';
import { ref } from 'vue';
import type { createMarkdownRenderer } from '../markdown/renderer';
type UsePreferencesDependencies =
  ReturnType<typeof createMarkdownRenderer>;

export function usePreferences(input: UsePreferencesDependencies) {
  const { CONTENT_ZOOM_DEFAULT,
  } = input;

  const theme = ref<Theme>('light');

  const fontSize = ref(14);

  const contentZoom = ref(CONTENT_ZOOM_DEFAULT);

  const userCssEnabled = ref(false);

  const userCssPath = ref('');
  const autoSave = ref(false);
  const autoSaveInterval = ref(300);

  return { theme,fontSize,contentZoom,userCssEnabled,userCssPath,autoSave,autoSaveInterval,
  };
}
