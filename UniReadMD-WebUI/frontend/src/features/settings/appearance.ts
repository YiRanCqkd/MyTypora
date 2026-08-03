import type { Theme } from './preferences-contracts';
export function useAppearance() {
  function applyTheme(mode: Theme) {
    document.documentElement.dataset.theme = mode;
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

  return { applyTheme,getUserCssLinkElement,applyUserCss,
  };
}
