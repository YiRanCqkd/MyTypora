import path from 'node:path';
import { expect, test } from '@playwright/test';

const SCREENSHOT_VIEWPORTS = [375, 768, 1280] as const;

async function openSource(page: Parameters<typeof test>[0]['page']) {
  await page.locator('#toggle-sourceview-btn').click();
  await expect(page.getByTestId('app-layout')).toHaveClass(/typora-sourceview-on/);
  await expect(page.getByTestId('editor-host')).toBeVisible();
}

async function replaceSource(
  page: Parameters<typeof test>[0]['page'],
  markdown: string,
) {
  const editor = page.locator('.cm-content').first();
  await editor.click();
  await page.keyboard.press('Control+A');
  await page.keyboard.insertText(markdown);
}

test.describe('Task 2 前端公开行为特征', () => {
  test('Given 默认文档，When 打开工作台，Then 渲染内容与响应式外观保持', async ({ page }) => {
    await page.goto('/');

    await expect(page.getByTestId('preview-host')).toContainText('UniReadMD');
    await expect(page.getByTestId('preview-host')).toContainText('在这里输入 Markdown 内容。');

    const screenshotDirectory = process.env.TASK2_SCREENSHOT_DIR;
    if (!screenshotDirectory) {
      return;
    }

    for (const width of SCREENSHOT_VIEWPORTS) {
      await page.setViewportSize({ height: 900, width });
      await expect(page.getByTestId('app-layout')).toBeVisible();
      await page.screenshot({
        animations: 'disabled',
        path: path.join(screenshotDirectory, `default-${width}.png`),
      });
    }
  });

  test('Given Markdown 输入，When 在 Source 编辑并切回，Then 渲染结果保持', async ({ page }) => {
    await page.goto('/');
    await openSource(page);
    await replaceSource(page, '# 特征标题\n\n这是可观察的编辑结果。');

    await page.locator('#toggle-sourceview-btn').click();

    await expect(page.getByTestId('preview-host').locator('h1')).toHaveText('特征标题');
    await expect(page.getByTestId('preview-host')).toContainText('这是可观察的编辑结果。');
  });

  test('Given 未保存内容，When 刷新会话，Then Markdown 与脏状态恢复', async ({ page }) => {
    await page.goto('/');
    await expect(page.getByTestId('preview-host')).toContainText('UniReadMD');
    await openSource(page);
    await expect(page.locator('.cm-content').first()).toContainText('# UniReadMD');
    await replaceSource(page, '# 会话恢复\n\n刷新后仍然存在。');
    await expect(page.locator('.cm-content').first()).toContainText('刷新后仍然存在。');
    await expect(page.locator('.dirty-dot')).toBeVisible();

    await page.reload();

    await expect(page.getByTestId('preview-host').locator('h1')).toHaveText('会话恢复');
    await expect(page.getByTestId('preview-host')).toContainText('刷新后仍然存在。');
    await expect(page.locator('.dirty-dot')).toBeVisible();

    const screenshotDirectory = process.env.TASK2_SCREENSHOT_DIR;
    if (!screenshotDirectory) {
      return;
    }

    for (const width of SCREENSHOT_VIEWPORTS) {
      await page.setViewportSize({ height: 900, width });
      await expect(page.getByTestId('app-layout')).toBeVisible();
      await page.screenshot({
        animations: 'disabled',
        path: path.join(screenshotDirectory, `session-recovered-${width}.png`),
      });
    }
  });

  test('Given 已编辑文档，When 保存到工作区，Then 文件树公开该文档', async ({ page }) => {
    await page.goto('/');
    await openSource(page);
    await replaceSource(page, '# 工作区文档');
    await page.getByTestId('tab-files').click();

    await page.keyboard.press('Control+S');

    await expect(page.locator('.dirty-dot')).toHaveCount(0);
    await expect(page.getByTestId('files-tree-view')).toContainText('virtual');
    await expect(page.getByTestId('files-tree-view')).toContainText(/document-\d+\.md/);
  });

  test('Given 拼写设置面板，When 修改设置，Then 状态与桥接读取一致', async ({ page }) => {
    await page.goto('/');
    await page.keyboard.press('Control+Alt+L');
    const panel = page.getByTestId('spell-panel');
    await expect(panel).toBeVisible();

    await panel.locator('input[type="checkbox"]').uncheck();
    await panel.locator('select').selectOption('zh-CN');

    await expect(panel.locator('input[type="checkbox"]')).not.toBeChecked();
    await expect(page.getByTestId('spell-dictionary-status')).toContainText('Dictionary Ready');
    const settings = await page.evaluate(async () => window.nativeBridge?.settings.getAll());
    expect(settings?.settings.spellCheckEnabled).toBe(false);
    expect(settings?.settings.spellCheckLanguage).toBe('zh-CN');
  });

  test('Given 当前 Markdown，When 使用导出快捷键，Then HTML 产物包含渲染内容', async ({ page }) => {
    await page.goto('/');
    await openSource(page);
    await replaceSource(page, '# 导出特征\n\n导出正文。');
    await page.locator('#toggle-sourceview-btn').click();

    await page.keyboard.press('Control+E');

    await expect.poll(async () => page.evaluate(async () => {
      const result = await window.nativeBridge?.file.openFromPath({
        filePath: '/virtual/document.html',
      });
      return result?.content ?? '';
    })).toContain('>导出特征</h1>');
  });
});
