import { expect, test, type Page } from './fixtures';

async function openDocument(page: Page, content = '') {
  await page.addInitScript((content) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content, sourceMode: false, filePath: '', isDirty: true, cursorHint: 1,
    }));
  }, content);
  await page.goto('/');
  await expect(page.locator('.preview')).toBeVisible();
}
async function rightClickBlank(page: Page) {
  const preview = page.locator('.preview');
  const size = (await preview.boundingBox())!;
  await preview.click({ button: 'right', position: { x: size.width - 12, y: size.height - 20 } });
  await expect(page.getByRole('menu', { name: '工作区插入菜单' })).toBeVisible();
}

test('空白工作区右键插入标题并继续编辑，撤销保留原文', async ({ page }) => {
  await openDocument(page, '原有正文');
  await rightClickBlank(page);
  await page.getByRole('menuitem', { name: '标题', exact: true }).click();
  await page.getByRole('menuitem', { name: '二级标题' }).click();
  const input = page.locator('.preview h2[contenteditable="true"]');
  await expect(input).toBeFocused();
  await expect(input).toHaveText('标题');
  await input.press('Control+a');
  await page.keyboard.insertText('新标题');
  await input.press('Control+Enter');
  await expect(page.locator('.preview h2')).toHaveText('新标题');
  await expect(page.locator('.preview p')).toHaveText('原有正文');
  await page.locator('#toggle-sourceview-btn').click();
  await page.locator('.cm-content').focus();
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+z');
  await expect(page.locator('.cm-content')).not.toContainText('标题');
  await expect(page.locator('.cm-content')).toContainText('原有正文');
});

test('空白工作区右键插入指定行列表格，直接编辑表头', async ({ page }) => {
  await openDocument(page);
  await rightClickBlank(page);
  await page.getByRole('menuitem', { name: '表格…' }).click();
  await page.getByRole('spinbutton', { name: '数据行数' }).fill('3');
  await page.getByRole('spinbutton', { name: '列数' }).fill('4');
  await page.getByRole('button', { name: '插入表格', exact: true }).click();
  await expect(page.locator('.preview table tbody tr')).toHaveCount(3);
  await expect(page.locator('.preview table th')).toHaveCount(4);
  const cell = page.locator('.preview [contenteditable="true"]');
  await expect(cell).toBeFocused();
  await cell.fill('名称');
  await cell.press('Enter');
  await expect(page.locator('.preview table th').first()).toHaveText('名称');
  await cell.press('Escape');
  await page.locator('.preview table td').first().click({ button: 'right' });
  await expect(page.getByTestId('preview-table-menu')).toBeVisible();
  await expect(page.getByRole('menu', { name: '工作区插入菜单' })).toHaveCount(0);
});

test('空白文档菜单支持键盘导航、Escape 和外部点击关闭', async ({ page }) => {
  await openDocument(page);
  await rightClickBlank(page);
  await expect(page.getByRole('menuitem', { name: '标题', exact: true })).toBeFocused();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await expect(page.locator('.preview hr')).toHaveCount(1);
  await rightClickBlank(page);
  await page.keyboard.press('Escape');
  await expect(page.getByRole('menu', { name: '工作区插入菜单' })).toHaveCount(0);
  await rightClickBlank(page);
  await page.locator('#outline-btn').click();
  await expect(page.getByRole('menu', { name: '工作区插入菜单' })).toHaveCount(0);
});

test('长文档末尾插入表格，自动滚入后单元格持续获得编辑焦点', async ({ page }) => {
  await openDocument(page, Array.from({ length: 70 }, (_, index) => `正文 ${index}`).join('\n\n'));
  await page.locator('.preview').evaluate((node) => { node.scrollTop = node.scrollHeight; });
  await page.waitForTimeout(100);
  await rightClickBlank(page);
  await page.getByRole('menuitem', { name: '表格…' }).click();
  await page.getByRole('button', { name: '插入表格', exact: true }).click();
  const cell = page.locator('.preview [contenteditable="true"]');
  await expect(cell).toBeFocused();
  await page.waitForTimeout(150);
  await expect(cell).toBeFocused();
  await cell.fill('长文表头');
  await cell.press('Enter');
  await cell.press('Escape');
  await expect(page.locator('.preview table th').first()).toHaveText('长文表头');
});
