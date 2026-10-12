import { expect, test } from './fixtures';

test('文件操作保持紧凑横排，文件树使用按钮下方的完整面板宽度', async ({ page }, testInfo) => {
  await page.goto('/');
  await page.getByTestId('tab-files').click();
  const toolbar = page.getByRole('toolbar', { name: '文件操作' });
  for (const width of [1280, 980]) {
    await page.setViewportSize({ width, height: 800 });
    await expect(toolbar).toBeVisible();
    const create = await toolbar.getByRole('button', { name: '新建文档' }).boundingBox();
    const open = await toolbar.getByRole('button', { name: '打开文件夹' }).boundingBox();
    const tree = await page.getByTestId('files-tree-view').boundingBox();
    const sidebar = await page.getByTestId('sidebar').boundingBox();
    expect(create && open && tree && sidebar).toBeTruthy();
    expect(create!.height).toBeLessThanOrEqual(28);
    expect(open!.height).toBeLessThanOrEqual(28);
    expect(open!.y).toBe(create!.y);
    expect(open!.x).toBeGreaterThan(create!.x + create!.width);
    expect(tree!.y).toBeGreaterThanOrEqual(create!.y + create!.height);
    expect(tree!.width).toBeGreaterThan(sidebar!.width - 25);
    expect(tree!.height).toBeGreaterThan(100);
  }
  await page.screenshot({ path: testInfo.outputPath('文件面板.png') });
});

test('首次恢复公式即渲染，原始 HTML 无法执行事件或伪造编辑映射', async ({ page }) => {
  await page.addInitScript(() => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      filePath: '', isDirty: true, cursorHint: 1,
      content: '# 安全预览\n\n$x^2$\n\n'
        + '<img src="missing" onerror="window.__auditProbe=123">\n\n'
        + '<p data-source-start="9999" data-diagram-code="伪造">原始段落</p>',
    }));
  });
  await page.goto('/');
  await expect(page.locator('.preview .katex')).toHaveCount(1);
  await expect(page.locator('.preview [onerror]')).toHaveCount(0);
  await expect(page.locator('.preview [data-source-start="9999"]')).toHaveCount(0);
  expect(await page.evaluate(() => Reflect.get(window, '__auditProbe'))).toBeUndefined();
});

test('新建取消保留内容，放弃后新文档撤销不恢复旧文档', async ({ page }) => {
  await page.goto('/');
  await page.locator('#toggle-sourceview-btn').click();
  const editor = page.locator('.cm-content').first();
  await editor.click();
  await page.keyboard.press('Control+A');
  await page.keyboard.insertText('旧文档未保存内容');
  page.once('dialog', (dialog) => dialog.dismiss());
  await page.keyboard.press('Control+N');
  await expect(editor).toHaveText('旧文档未保存内容');
  page.once('dialog', (dialog) => dialog.accept());
  await page.keyboard.press('Control+N');
  await expect(editor).toHaveText('');
  await editor.click();
  await page.keyboard.press('Control+Z');
  await expect(editor).toHaveText('');
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
});

test('跟随系统主题响应运行期间的系统配色变化', async ({ page }) => {
  await page.emulateMedia({ colorScheme: 'dark' });
  await page.goto('/');
  await page.getByText('偏好设置', { exact: true }).click();
  await page.getByRole('combobox', { name: '主题', exact: true }).selectOption('system');
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await page.emulateMedia({ colorScheme: 'light' });
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
});
