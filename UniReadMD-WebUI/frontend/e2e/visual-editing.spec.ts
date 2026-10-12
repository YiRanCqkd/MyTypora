import { expect, test, type Page } from './fixtures';

async function open(page: Page, content: string) {
  await page.addInitScript((content) => localStorage.setItem('unireadmd.phase2.session.v1',
    JSON.stringify({ content, filePath: '', isDirty: true, sourceMode: false })), content);
  await page.goto('/');
  await expect(page.locator('.preview')).toBeVisible();
}

test('正文和标题直接编辑保留样式与光标，源码同步且可以撤销重做', async ({ page }) => {
  await open(page, '# 标题\n\n正文 **加粗** 和 *斜体*');
  const paragraph = page.locator('.preview p');
  await paragraph.click();
  await expect(paragraph).toHaveAttribute('contenteditable', 'true');
  await expect(page.locator('.preview-inline-editor-input')).toHaveCount(0);
  await expect(paragraph.locator('strong')).toHaveText('加粗');
  await paragraph.press('End');
  await page.keyboard.insertText('，中文新增');
  await expect(paragraph).toBeFocused();
  await page.waitForTimeout(150);
  await expect(paragraph).toBeFocused();
  await paragraph.press('Control+z');
  await expect(page.locator('.preview p')).not.toContainText('中文新增');
  await page.locator('.preview p').press('Control+Shift+z');
  await expect(page.locator('.preview p')).toContainText('中文新增');
  await page.locator('#toggle-sourceview-btn').check();
  await expect(page.locator('.cm-content')).toContainText('**加粗**');
  await expect(page.locator('.cm-content')).toContainText('中文新增');
});

test('标题 Enter 新建正文并继续输入，列表与引用保持结构', async ({ page }) => {
  await open(page, '## 标题\n\n- 列表\n\n> 引用');
  await page.locator('.preview h2').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await expect(page.getByRole('textbox', { name: '原位可视编辑' })).toBeFocused();
  await page.keyboard.insertText('新段落');
  await page.keyboard.press('Control+Enter');
  await expect(page.locator('.preview p').filter({ hasText: '新段落' })).toBeVisible();
  await page.locator('.preview li').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.insertText('第二项');
  await page.keyboard.press('Control+Enter');
  await expect(page.locator('.preview li')).toHaveCount(2);
  await expect(page.locator('.preview li').last()).toHaveText('第二项');
  await page.locator('.preview blockquote p').click();
  await page.keyboard.press('End');
  await page.keyboard.insertText('新增');
  await page.keyboard.press('Control+Enter');
  await expect(page.locator('.preview blockquote')).toContainText('引用新增');
});

test('中文组合输入结束前保留活动 DOM，格式与链接快捷键写回 Markdown', async ({ page }) => {
  await open(page, '原始正文');
  const input = page.locator('.preview p');
  await input.click();
  await input.evaluate((node) => {
    node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    node.textContent = '中文输入完成';
    node.dispatchEvent(new InputEvent('input', { bubbles: true, isComposing: true }));
  });
  await page.waitForTimeout(160);
  await expect(input).toBeFocused();
  await input.evaluate((node) => node.dispatchEvent(new CompositionEvent('compositionend',
    { bubbles: true, data: '中文输入完成' })));
  await input.press('Control+a');
  await input.press('Control+b');
  await expect(input.locator('b, strong')).toHaveText('中文输入完成');
  await input.press('Control+k');
  await page.getByRole('textbox', { name: '链接地址' }).fill('https://example.com');
  await page.getByRole('button', { name: '确定', exact: true }).click();
  await expect(input.locator('a')).toHaveAttribute('href', 'https://example.com');
  await page.locator('#toggle-sourceview-btn').check();
  await expect(page.locator('.cm-content')).toContainText('**[中文输入完成]');
  await expect(page.locator('.cm-content')).toContainText('https://example.com');
});

test('Alt 单击保留 Markdown 精确编辑，图片块沿用原编辑方式', async ({ page }) => {
  await open(page, '# 标题\n\n![图](missing.png)');
  await page.locator('.preview h1').click({ modifiers: ['Alt'] });
  await expect(page.locator('.preview-inline-editor-input')).toHaveValue('# 标题');
  await page.locator('.preview-inline-editor-input').press('Escape');
  await page.locator('.preview img').click();
  await expect(page.locator('.preview-inline-editor-input')).toHaveValue('![图](missing.png)');
  await expect(page.locator('.preview [contenteditable="true"]')).toHaveCount(0);
});
