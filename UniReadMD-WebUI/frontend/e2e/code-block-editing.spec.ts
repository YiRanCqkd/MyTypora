import { expect, test, type Page } from './fixtures';

async function open(page: Page, content: string, theme = 'light') {
  await page.addInitScript(({ content, theme }) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content, filePath: '', isDirty: false, sourceMode: false,
    }));
    localStorage.setItem('unireadmd.web-theme.v1', theme);
  }, { content, theme });
  await page.goto('/');
  await expect(page.locator('.preview')).toBeVisible();
}

for (const theme of ['light', 'dark']) {
  test(`${theme}：代码块源码保留背景、内边距与高亮，长行不折行且同步滚动`, async ({ page }) => {
    const source = '```js\nconst text = "' + '长行abc'.repeat(100) + '";\n\n```';
    await open(page, source, theme);
    const before = await page.locator('.preview pre').boundingBox();
    await page.locator('.preview pre').click({ position: { x: 40, y: 20 } });
    const editor = page.locator('.preview-inline-code');
    const input = editor.locator('textarea');
    await expect(input).toHaveValue(source);
    await expect(input).toBeFocused();
    await expect(editor.locator('.hljs-keyword')).toHaveText('const');
    await expect(input).toHaveCSS('white-space', 'pre');
    await expect(input).toHaveCSS('padding-left', '13px');
    const after = (await editor.boundingBox())!;
    expect(Math.abs(after.x - before!.x)).toBeLessThan(1);
    expect(Math.abs(after.width - before!.width)).toBeLessThan(1);
    const data = await editor.evaluate((node) => {
      const input = node.querySelector('textarea')!;
      const mirror = node.querySelector('pre')!;
      return { background: getComputedStyle(node).backgroundColor,
        font: getComputedStyle(input).fontFamily, height: input.clientHeight,
        line: parseFloat(getComputedStyle(input).lineHeight),
        mirror: mirror.textContent, scrollable: input.scrollWidth > input.clientWidth };
    });
    expect(data.background).not.toBe('rgba(0, 0, 0, 0)');
    expect(data.font).toContain('Consolas');
    expect(data.height).toBeLessThan(data.line * 6 + 24);
    expect(data.mirror).toBe(source + '\u200b');
    expect(data.scrollable).toBe(true);
    await page.screenshot({ path: test.info().outputPath(`code-editor-${theme}.png`) });
    await input.evaluate((node: HTMLTextAreaElement) => { node.scrollLeft = 320; });
    await expect.poll(() => editor.locator('pre').evaluate((node) => node.scrollLeft)).toBe(320);
    await page.setViewportSize({ width: 1600, height: 1000 });
    await expect.poll(async () => {
      const anchor = (await page.locator('.preview pre').boundingBox())!;
      const edit = (await editor.boundingBox())!;
      return Math.abs(anchor.width - edit.width);
    }).toBeLessThan(1);
    await input.press('Control+Enter');
    await expect(editor).toHaveCount(0);
    await page.locator('#toggle-sourceview-btn').check();
    expect((await page.locator('.cm-line').allTextContents()).join('\n')).toBe(source);
  });
}

test('代码输入、中文组合、取消与保存不会丢失围栏或空行', async ({ page }) => {
  const source = '~~~js\nconst value = 1;\n\n~~~~';
  await open(page, source);
  await page.locator('.preview pre').click();
  const input = page.locator('.preview-inline-code textarea');
  await input.evaluate((node: HTMLTextAreaElement) => {
    node.setSelectionRange(6, 6);
    node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
  });
  const color = await input.evaluate((node) => getComputedStyle(node).color);
  await expect(input).toHaveCSS('-webkit-text-fill-color', color);
  await expect(page.locator('.preview-code-highlight')).toHaveCSS('visibility', 'hidden');
  await input.evaluate((node) => {
    node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
  });
  await input.fill('~~~js\nconst value = "中文";\n\n~~~~');
  await page.waitForTimeout(200);
  await input.press('Escape');
  await expect(page.locator('.preview pre code')).toHaveText('const value = 1;\n\n');
  await page.locator('.preview pre').click();
  await input.fill('~~~js\nconst value = "中文";\n\n~~~~');
  await input.press('Control+s');
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  await page.locator('#toggle-sourceview-btn').check();
  await expect(page.locator('.cm-content')).toContainText('const value = "中文";');
  await expect(page.locator('.cm-content')).toContainText('~~~~');
});

test('源码高亮安全回退，切回正文恢复原位编辑布局', async ({ page }) => {
  await open(page, '```unknown\n<script>alert(1)</script>\n```\n\n正文');
  await page.locator('.preview pre').click();
  await expect(page.locator('.preview-code-highlight script')).toHaveCount(0);
  await expect(page.locator('.preview-code-highlight')).toContainText('<script>');
  await page.locator('.preview-inline-code textarea').press('Control+Enter');
  await page.locator('.preview p').click();
  await expect(page.locator('.preview p[contenteditable="true"]')).toBeFocused();
  await expect(page.locator('.preview-inline-code')).toHaveCount(0);
  await expect(page.locator('.preview p')).toHaveCSS('padding-top', '0px');
});
