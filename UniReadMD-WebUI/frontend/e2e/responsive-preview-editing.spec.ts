import { expect, test, type Page } from './fixtures';

async function openDocument(page: Page, content: string) {
  await page.addInitScript((content) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content, sourceMode: false, filePath: '', isDirty: true, cursorHint: 1,
    }));
  }, content);
  await page.goto('/');
  await expect(page.getByTestId('preview-host')).toBeVisible();
}

test('窗口放大后正文和编辑框充分展开，专注与源码视图保持可用宽度', async ({ page }) => {
  await page.setViewportSize({ width: 2000, height: 950 });
  await openDocument(page, '宽屏正文。'.repeat(80));
  const host = page.getByTestId('editor-preview');
  const preview = page.getByTestId('preview-host');
  await expect.poll(async () => Math.abs((await host.boundingBox())!.width
    - (await preview.boundingBox())!.width)).toBeLessThanOrEqual(1);
  await preview.locator('p').click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  expect((await input.boundingBox())!.width).toBeGreaterThan(1200);
  await page.screenshot({ path: test.info().outputPath('wide-preview-editing.png') });
  await input.press('Control+Alt+f');
  await expect.poll(async () => Math.abs((await host.boundingBox())!.width
    - (await preview.boundingBox())!.width)).toBeLessThanOrEqual(1);
  await expect(input).toBeFocused();
  await input.press('Control+Enter');
  await page.locator('#toggle-sourceview-btn').click();
  expect((await page.locator('.cm-content').boundingBox())!.width).toBeGreaterThan(1200);
});

test('容器单独缩放后编辑框重新对齐并保留内容、焦点和反向选区', async ({ page }) => {
  await openDocument(page, '容器布局正文。'.repeat(50) + '\n\n后续正文');
  const paragraph = page.locator('.preview p').first();
  await paragraph.click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await input.evaluate((node: HTMLTextAreaElement) => node.setSelectionRange(4, 12, 'backward'));
  const content = await input.inputValue();
  await page.getByTestId('editor-preview').evaluate((node) => { node.style.width = '580px'; });
  await expect.poll(async () => {
    const editor = (await input.boundingBox())!;
    const anchor = (await paragraph.boundingBox())!;
    return Math.abs(editor.width - anchor.width) + Math.abs(editor.x - anchor.x)
      + Math.abs(editor.y - anchor.y);
  }).toBeLessThanOrEqual(2);
  await expect(input).toBeFocused();
  await expect(input).toHaveValue(content);
  expect(await input.evaluate((node: HTMLTextAreaElement) => ({
    start: node.selectionStart, end: node.selectionEnd, direction: node.selectionDirection,
  }))).toEqual({ start: 4, end: 12, direction: 'backward' });
  await page.screenshot({ path: test.info().outputPath('resized-preview-editing.png') });
  await input.press('Control+Enter');
  await expect(page.locator('.preview p').last()).toHaveText('后续正文');
});

test('正向扩展长选区跟随末端，反向扩展跟随开头', async ({ page }) => {
  await openDocument(page, '起始正文\n\n后续正文');
  await page.locator('.preview p').first().click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await input.fill('选区内容\n'.repeat(80));
  await input.press('Control+Home');
  await input.press('Control+Shift+End');
  await expect.poll(() => page.getByTestId('preview-host').evaluate((node) => node.scrollTop))
    .toBeGreaterThan(1000).catch(async (error) => {
      const geometry = await input.evaluate((node: HTMLTextAreaElement) => ({
        selection: [node.selectionStart, node.selectionEnd, node.selectionDirection],
        input: node.getBoundingClientRect().toJSON(), scrollTop: node.scrollTop,
        scrollHeight: node.scrollHeight, style: node.getAttribute('style'),
        containers: ['.preview', '.editor-preview', '.main'].map((selector) => {
          const host = document.querySelector<HTMLElement>(selector)!;
          return { selector, rect: host.getBoundingClientRect().toJSON(),
            scrollTop: host.scrollTop, scrollHeight: host.scrollHeight };
        }),
      }));
      await test.info().attach('选区失败布局', {
        body: JSON.stringify(geometry, null, 2), contentType: 'application/json',
      });
      throw error;
    });
  await input.press('Control+End');
  await input.press('Control+Shift+Home');
  await expect.poll(() => page.getByTestId('preview-host').evaluate((node) => node.scrollTop))
    .toBeLessThan(100);
});

test('编辑框一次滚轮只滚动一次，Ctrl 滚轮只执行内容缩放', async ({ page }) => {
  await openDocument(page, Array.from({ length: 80 }, (_, i) => `段落 ${i}`).join('\n\n'));
  const preview = page.getByTestId('preview-host');
  await preview.locator('p').nth(20).click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  const before = await preview.evaluate((node) => node.scrollTop);
  await input.dispatchEvent('wheel', { deltaY: 60, bubbles: true, cancelable: true });
  await expect.poll(() => preview.evaluate((node) => node.scrollTop)).toBe(before + 60);
  const zoomScroll = await preview.evaluate((node) => node.scrollTop);
  await input.dispatchEvent('wheel', {
    deltaY: 60, ctrlKey: true, bubbles: true, cancelable: true,
  });
  await expect(page.locator('#footer-content-zoom')).toHaveText('90%');
  // 缩放的排版更新另行定位，事件处理本身不能叠加一次普通滚动。
  expect(await preview.evaluate((node) => node.scrollTop)).toBeLessThanOrEqual(zoomScroll);
});

test('直接编辑沿用写作主题背景与自定义正文文字颜色', async ({ page }) => {
  await page.addInitScript(() => {
    localStorage.setItem('unireadmd.writing.v1', JSON.stringify({ lightTheme: 'sepia' }));
  });
  await openDocument(page, '主题中的正文');
  await expect(page.locator('html')).toHaveAttribute('data-writing-theme', 'sepia');
  await page.addStyleTag({ content: '.preview p { color: rgb(18, 52, 86); }' });
  const background = await page.locator('.preview').evaluate((node) =>
    getComputedStyle(node).backgroundColor);
  await page.locator('.preview p').click({ modifiers: ['Alt'] });
  await expect(page.locator('.preview-inline-editor-input'))
    .toHaveCSS('color', 'rgb(18, 52, 86)');
  await expect(page.locator('.preview-inline-editor')).toHaveCSS('background-color', background);
});
