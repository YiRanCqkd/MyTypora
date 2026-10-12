import { test as base, expect } from '@playwright/test';

// 旧流程和视觉基线显式使用展开的工作区；首次启动默认状态由专门用例覆盖。
export const test = base.extend({
  page: async ({ page }, use) => {
    await page.addInitScript(() => {
      if (!localStorage.getItem('unireadmd.view-visibility.v1')) {
        localStorage.setItem('unireadmd.view-visibility.v1', JSON.stringify({
          documentToolsVisible: true, preferencesVisible: true,
        }));
      }
    });
    await use(page);
  },
});
export { expect };
export type { Page, Locator } from '@playwright/test';
