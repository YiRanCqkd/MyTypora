import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const { expect } = require('../frontend/node_modules/@playwright/test');

// 原有工具流程显式展开区域，首次启动默认隐藏由独立原生用例验证。
export async function showWorkspaceControls(application, page) {
  for (const [id, selector] of [['view-document-tools', '.document-tools'],
    ['view-preferences', '.preferences-panel']]) {
    if (!await page.locator(selector).isVisible()) {
      await application.evaluate(({ Menu }, id) => {
        Menu.getApplicationMenu().getMenuItemById(id).click();
      }, id);
      await expect(page.locator(selector)).toBeVisible();
    }
  }
}
