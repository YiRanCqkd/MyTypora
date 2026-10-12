import type { Theme } from './preferences-contracts';
import { onBeforeUnmount } from 'vue';
export function useAppearance() {
  let requestedTheme: Theme = 'light';
  const media = window.matchMedia?.('(prefers-color-scheme: dark)');
  const updateTheme = () => {
    document.documentElement.dataset.theme = requestedTheme === 'system'
      ? media?.matches ? 'dark' : 'light' : requestedTheme;
  };
  media?.addEventListener('change', updateTheme);
  onBeforeUnmount(() => media?.removeEventListener('change', updateTheme));
  function applyTheme(mode: Theme) {
    requestedTheme = mode;
    updateTheme();
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

  function applyUserCss(css = '') {
    getUserCssLinkElement().remove();
    let style = document.getElementById('unireadmd-user-style');
    if (!style) {
      style = document.createElement('style');
      style.id = 'unireadmd-user-style';
      document.head.appendChild(style);
    }
    style.textContent = css;
  }

  return { applyTheme,getUserCssLinkElement,applyUserCss,
  };
}
