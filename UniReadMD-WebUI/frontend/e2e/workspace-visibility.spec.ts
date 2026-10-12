import { expect, test } from '@playwright/test';

test('首次启动隐藏快捷操作栏与偏好设置，恢复已保存的显示选择', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('.preview')).toBeVisible();
  await expect(page.getByRole('toolbar', { name: '文档工具' })).toBeHidden();
  await expect(page.locator('.preferences-panel')).toHaveCount(0);
  await page.evaluate(async () => {
    await window.nativeBridge.settings.set('documentToolsVisible', true);
  });
  await page.reload();
  await expect(page.getByRole('toolbar', { name: '文档工具' })).toBeVisible();
  await expect(page.locator('.preferences-panel')).toHaveCount(0);
  await page.evaluate(async () => {
    await window.nativeBridge.settings.set('documentToolsVisible', false);
    await window.nativeBridge.settings.set('preferencesVisible', true);
  });
  await page.reload();
  await expect(page.getByRole('toolbar', { name: '文档工具' })).toBeHidden();
  await expect(page.locator('.preferences-panel')).toBeVisible();
});

test('圆形源码选择可取消，侧栏入口始终显示且与模式独立', async ({ page }) => {
  await page.goto('/');
  const source = page.getByRole('checkbox', { name: '源码模式' });
  const sidebar = page.locator('#outline-btn');
  await expect(source).not.toBeChecked();
  await expect(source).toHaveCSS('border-radius', '50%');
  await source.check();
  await expect(page.locator('.editor-host')).toBeVisible();
  await expect(page.locator('.preview')).toBeHidden();
  await expect(sidebar).toBeVisible();
  await expect(sidebar).toHaveText('隐藏侧栏');
  await sidebar.click();
  await expect(page.getByTestId('sidebar')).toBeHidden();
  await expect(sidebar).toHaveText('显示侧栏');
  await source.uncheck();
  await expect(page.locator('.preview')).toBeVisible();
  await expect(page.getByTestId('sidebar')).toBeHidden();
  await expect(sidebar).toBeVisible();
  await sidebar.click();
  await expect(page.getByTestId('sidebar')).toBeVisible();
  await source.focus();
  await page.keyboard.press('Space');
  await expect(source).toBeChecked();
  await expect(sidebar).toBeVisible();
  await expect(page.getByTestId('sidebar')).toBeVisible();
});
