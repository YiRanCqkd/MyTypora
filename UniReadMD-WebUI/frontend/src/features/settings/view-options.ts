import type { SpellDictionaryStatus } from './preferences-contracts';
import { ref } from 'vue';

export function useViewOptions() {
  const documentToolsVisible = ref(false);
  const preferencesVisible = ref(false);
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

  return { documentToolsVisible,preferencesVisible,isNativeWindow,pinOutline,
    sourceMode,focusMode,typewriterMode,
    showSpellPanel,spellCheckEnabled,spellCheckLanguage,spellCheckLanguages,
    spellDictionaryStatus,
  };
}
