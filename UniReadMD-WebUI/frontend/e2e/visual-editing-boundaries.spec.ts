import { expect, test, type Page } from './fixtures';

async function open(page: Page, content: string) {
  await page.addInitScript((content) => localStorage.setItem('unireadmd.phase2.session.v1',
    JSON.stringify({ content, filePath: '', isDirty: true, sourceMode: false })), content);
  await page.goto('/');
  await expect(page.locator('.preview')).toBeVisible();
}

test('仅点击阅读和退出编辑不改写原有 Markdown 语法', async ({ page }) => {
  const source = '# 标题\n\n正文 __加粗__ 与 [链接][id]\n\n[id]: https://example.com';
  await open(page, source);
  await page.locator('.preview p').click();
  await page.locator('#toggle-sourceview-btn').check();
  await expect.poll(async () => (await page.locator('.cm-line').allTextContents()).join('\n'))
    .toBe(source);
});

test('文末 Enter 后可继续输入，段首退格合并正文并保留光标位置', async ({ page }) => {
  await open(page, '第一段');
  await page.locator('.preview p').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.insertText('第二段');
  await page.keyboard.press('Home');
  await page.keyboard.press('Backspace');
  await expect(page.locator('.preview p')).toHaveCount(1);
  await expect(page.locator('.preview p')).toHaveText('第一段第二段');
  await page.keyboard.insertText('中间');
  await page.keyboard.press('Control+Enter');
  await expect(page.locator('.preview p')).toHaveText('第一段中间第二段');
});

test('引用中 Enter 连续输入，空列表项 Enter 回到正文', async ({ page }) => {
  await open(page, '> 引用\n\n- 第一项');
  await page.locator('.preview blockquote p').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.insertText('引用第二行');
  await page.keyboard.press('Control+Enter');
  await expect(page.locator('.preview blockquote')).toContainText('引用第二行');
  await page.locator('.preview li').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  await page.keyboard.insertText('列表后的正文');
  await page.keyboard.press('Control+Enter');
  await expect(page.locator('.preview li')).toHaveCount(1);
  await expect(page.locator('.preview > p')).toHaveText('列表后的正文');
});

test('上下边界移动到相邻块，长段落与最大化布局保持原位输入', async ({ page }) => {
  await open(page, '第一段\n\n第二段');
  await page.locator('.preview p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.press('ArrowDown');
  await expect(page.locator('.preview p').last()).toBeFocused();
  await page.keyboard.press('Home');
  await page.keyboard.press('ArrowUp');
  await expect(page.locator('.preview p').first()).toBeFocused();
  await page.setViewportSize({ width: 1920, height: 1080 });
  await expect(page.locator('.preview p').first()).toBeFocused();
  const preview = (await page.locator('.preview').boundingBox())!;
  const paragraph = (await page.locator('.preview p').first().boundingBox())!;
  expect(paragraph.width).toBeGreaterThan(preview.width * 0.8);
});

test('可视编辑安全粘贴保留格式，非法链接有明确提示', async ({ page }) => {
  await open(page, '原文');
  const paragraph = page.locator('.preview p');
  await paragraph.click();
  await paragraph.press('Control+a');
  await paragraph.evaluate((node) => {
    const data = new DataTransfer();
    data.setData('text/plain', '粘贴正文');
    data.setData('text/html', '<strong>粘贴正文</strong><script>window.__unsafe = true</script>');
    node.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true,
      clipboardData: data }));
  });
  await expect(paragraph.locator('strong, b')).toHaveText('粘贴正文');
  expect(await page.evaluate(() => Reflect.get(window, '__unsafe'))).toBeUndefined();
  await paragraph.press('Control+a');
  await paragraph.press('Control+k');
  await page.getByRole('textbox', { name: '链接地址' }).fill('javascript:alert(1)');
  await page.getByRole('button', { name: '确定', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('请输入');
  await page.getByRole('textbox', { name: '链接地址' }).fill('./其它文档.md');
  await page.getByRole('button', { name: '确定', exact: true }).click();
  await expect(paragraph.locator('a')).toHaveAttribute('href', './其它文档.md');
});

test('正在可视编辑时保存及导出会提交正文，导出不携带活动编辑属性', async ({ page }) => {
  await open(page, '初始正文');
  await page.evaluate(() => {
    window.nativeBridge.exporter.exportDocument = async (payload) => {
      Reflect.set(window, '__visualExport', payload.html);
      return { filePath: 'I:/virtual/export.html', format: 'html', engine: 'builtin' };
    };
  });
  await page.locator('.preview p').click();
  await page.keyboard.press('End');
  await page.keyboard.insertText('保存续写');
  await page.keyboard.press('Control+s');
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  await page.locator('.preview p').click();
  await page.keyboard.press('End');
  await page.keyboard.insertText('导出续写');
  await page.keyboard.press('Control+e');
  await expect.poll(() => page.evaluate(() => Reflect.get(window, '__visualExport')))
    .toContain('保存续写导出续写');
  const html = await page.evaluate(() => Reflect.get(window, '__visualExport'));
  expect(await page.evaluate((html) => new DOMParser().parseFromString(html, 'text/html')
    .querySelectorAll('[contenteditable="true"], .preview-visual-editing').length, html)).toBe(0);
});

test('刚输入正文立即右键，菜单保持打开并可插入表格', async ({ page }) => {
  await open(page, '初始正文');
  await page.locator('.preview p').click();
  await page.locator('.preview p').evaluate((node) => {
    node.textContent = '刚输入的正文';
    node.dispatchEvent(new InputEvent('input', { bubbles: true }));
    const host = node.closest('.preview')!;
    const bounds = host.getBoundingClientRect();
    host.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true,
      clientX: bounds.right - 20, clientY: bounds.bottom - 20 }));
  });
  await expect(page.getByRole('menu', { name: '工作区插入菜单' })).toBeVisible();
  await page.getByRole('menuitem', { name: '表格…' }).click();
  await page.getByRole('button', { name: '插入表格', exact: true }).click();
  await expect(page.locator('.preview table th[contenteditable="true"]')).toBeFocused();
  await expect(page.locator('.preview > p')).toHaveText('刚输入的正文');
});
