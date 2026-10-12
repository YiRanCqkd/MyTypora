import { expect, test } from './fixtures';
import AxeBuilder from '@axe-core/playwright';

async function openDocument(page: Parameters<typeof test>[0]['page'], content: string) {
  await page.addInitScript((content) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content, sourceMode: false, filePath: '', isDirty: true, cursorHint: 1,
    }));
  }, content);
  await page.goto('/');
  await expect(page.getByTestId('preview-host')).toBeVisible();
}

test('标题编辑继承原字号和行高，退出后内容与位置保留', async ({ page }) => {
  await openDocument(page, '# 连续写作\n\n正文');
  const heading = page.locator('.preview h1');
  const typography = await heading.evaluate((node) => ({
    size: getComputedStyle(node).fontSize, height: getComputedStyle(node).lineHeight,
  }));
  await heading.click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await expect(input).toHaveCSS('font-size', typography.size);
  await expect(input).toHaveCSS('line-height', typography.height);
  await input.fill('# 标题编辑完成');
  await input.press('Control+Enter');
  await expect(heading).toHaveText('标题编辑完成');
});

test('长段落编辑框完整展示并增高，不遮挡后续正文，缩窄窗口仍可写作', async ({ page }) => {
  await openDocument(page, '长段落内容。'.repeat(90) + '\n\n后续正文');
  await page.locator('.preview p').first().click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  const complete = async () => input.evaluate((node: HTMLTextAreaElement) =>
    node.clientHeight >= node.scrollHeight - 1);
  await expect.poll(complete).toBe(true);
  await input.fill('扩展内容。'.repeat(120));
  await expect.poll(complete).toBe(true);
  await page.setViewportSize({ width: 900, height: 900 });
  await expect.poll(complete).toBe(true);
  await expect.poll(async () => {
    const editor = await input.boundingBox();
    const following = await page.locator('.preview').getByText('后续正文', { exact: true })
      .boundingBox();
    return Boolean(editor && following && editor.y + editor.height <= following.y + 1);
  }).toBe(true);
  await input.press('Control+Enter');
  await expect(page.locator('.preview p').first()).toHaveText('扩展内容。'.repeat(120));
  await expect(page.locator('.preview').getByText('后续正文', { exact: true })).toBeAttached();
});

test('短格式源码不遮挡相邻正文，取消清除临时占位', async ({ page }) => {
  await openDocument(page, '前文 **粗体** 后文保持可见。');
  await page.locator('.preview strong').click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  const rect = await input.boundingBox();
  const following = await page.locator('.preview p').evaluate((node) => {
    const text = node.lastChild!;
    const range = document.createRange();
    range.selectNodeContents(text);
    return range.getBoundingClientRect().left;
  });
  expect(rect!.x + rect!.width).toBeLessThanOrEqual(following + 1);
  await input.press('Escape');
  await expect(page.locator('.preview p')).toHaveText('前文 粗体 后文保持可见。');
  await expect(page.locator('[style*="--preview-edit-"]')).toHaveCount(0);
});

test('空白预览可以直接开始写作并撤销，提示不写入源码', async ({ page }) => {
  await openDocument(page, '');
  await page.locator('.preview-empty-line').click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await expect(input).toBeFocused();
  await input.fill('从空白开始写作');
  await input.press('Control+Enter');
  await expect(page.locator('.preview p')).toHaveText('从空白开始写作');
  await page.locator('#toggle-sourceview-btn').click();
  await expect(page.locator('.cm-content')).not.toContainText('点击此处开始写作');
  await page.locator('.cm-content').focus();
  await page.keyboard.press('Control+z');
  await expect(page.locator('.cm-content')).toHaveText('');
});

test('窄窗口工具栏保持一行，更多工具可操作且 Escape 恢复焦点', async ({ page }) => {
  await page.setViewportSize({ width: 820, height: 700 });
  await openDocument(page, '# 紧凑工具栏');
  const tools = page.getByRole('toolbar', { name: '文档工具' });
  expect((await tools.boundingBox())!.height).toBeLessThanOrEqual(42);
  const more = page.getByRole('button', { name: '更多文档工具' });
  await more.click();
  await expect(page.getByRole('button', { name: '历史草稿', exact: true })).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(more).toBeFocused();
  await expect(more).toHaveAttribute('aria-expanded', 'false');
  await more.click();
  await page.getByRole('button', { name: '导出设置', exact: true }).click();
  await expect(page.getByRole('dialog', { name: '导出设置' })).toBeVisible();
});

test('连续多行输入时预览跟随光标，编辑区保持在可见范围', async ({ page }) => {
  await openDocument(page, '连续输入\n\n下一段');
  await page.locator('.preview p').first().click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await input.fill('连续输入行\n'.repeat(60));
  await expect.poll(() => page.getByTestId('preview-host').evaluate((node) => node.scrollTop))
    .toBeGreaterThan(0);
  await expect.poll(async () => {
    const editor = await input.boundingBox();
    const preview = await page.getByTestId('preview-host').boundingBox();
    return Boolean(editor && preview && editor.y + editor.height <= preview.y + preview.height - 4);
  }).toBe(true);
  await input.press('Enter');
  await expect(input).toBeFocused();
  await input.press('Control+Enter');
  await expect(page.locator('.preview')).toContainText('下一段');
});

test('HTML 导出不含空白提示和编辑占位，编辑中的正文保持完整', async ({ page }) => {
  await openDocument(page, '');
  await page.evaluate(() => {
    // 此用例检查正文节点清理；带字体和样式的完整导出由原生回归覆盖。
    localStorage.setItem('unireadmd.export-options', JSON.stringify({ includeStyles: false }));
    window.nativeBridge.exporter.exportDocument = async (payload) => {
      Reflect.set(window, '__experienceExport', payload.html);
      return { filePath: '/virtual/experience.html', format: 'html', engine: 'builtin' };
    };
  });
  await page.getByRole('combobox', { name: '导出格式' }).selectOption('html');
  const exportButton = page.getByRole('button', { name: '导出', exact: true });
  await exportButton.click();
  await expect.poll(() => page.evaluate(() => {
    const html = Reflect.get(window, '__experienceExport');
    if (!html) return -1;
    return new DOMParser().parseFromString(html, 'text/html').body.textContent?.trim().length;
  })).toBe(0);
  await page.locator('.preview-empty-line').click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await input.fill('完整正文 **加粗**');
  await expect(page.locator('.preview strong')).toHaveText('加粗');
  await page.keyboard.press('Control+e');
  await expect.poll(() => page.evaluate(() => {
    const html = String(Reflect.get(window, '__experienceExport') || '');
    const documentCopy = new DOMParser().parseFromString(html, 'text/html');
    const article = documentCopy.querySelector('article');
    return article?.textContent?.trim();
  })).toBe('完整正文 加粗');
  const clean = await page.evaluate(() => {
    const documentCopy = new DOMParser().parseFromString(
      String(Reflect.get(window, '__experienceExport')), 'text/html',
    );
    return documentCopy.body.querySelectorAll('.preview-empty-line, .preview-editing-block, '
      + '.preview-editing-fragment, [style*="--preview-edit-"]').length;
  });
  expect(clean).toBe(0);
});

test('深色主题状态栏的选中和悬停按钮保持可读对比度', async ({ page }) => {
  await page.addInitScript(() => {
    localStorage.setItem('unireadmd.writing.v1', JSON.stringify({ darkTheme: 'midnight' }));
  });
  await openDocument(page, '# 深色写作');
  await page.getByText('偏好设置', { exact: true }).click();
  await page.getByRole('combobox', { name: '主题', exact: true }).selectOption('dark');
  await expect(page.locator('html')).toHaveAttribute('data-writing-theme', 'midnight');
  await page.locator('#toggle-sourceview-btn').hover();
  await expect(page.locator('#outline-btn')).toHaveClass(/active/);
  const results = await new AxeBuilder({ page }).include('.status-bar')
    .withRules(['color-contrast']).analyze();
  expect(results.violations).toEqual([]);
  expect(results.incomplete).toEqual([]);
});

test('预览编辑切换专注模式时保持阅读行宽和输入焦点', async ({ page }) => {
  await page.setViewportSize({ width: 2400, height: 900 });
  await openDocument(page, '# 专注写作\n\n保持正文宽度和编辑位置');
  const preview = page.getByTestId('preview-host');
  const width = (await preview.boundingBox())!.width;
  expect(width).toBe((await page.getByTestId('editor-preview').boundingBox())!.width);
  await preview.locator('p').click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  const inputWidth = (await input.boundingBox())!.width;
  await input.press('Control+Alt+f');
  await expect(page.getByTestId('app-layout')).toHaveClass(/on-focus-mode/);
  await expect.poll(async () => (await preview.boundingBox())!.width).toBe(width);
  await expect.poll(async () => (await input.boundingBox())!.width).toBe(inputWidth);
  await expect(input).toBeFocused();
  await input.press('Control+Alt+f');
  await expect(page.getByTestId('app-layout')).not.toHaveClass(/on-focus-mode/);
  await expect.poll(async () => (await preview.boundingBox())!.width).toBe(width);
  await input.press('Escape');
  await expect(preview.locator('p')).toHaveText('保持正文宽度和编辑位置');
});
