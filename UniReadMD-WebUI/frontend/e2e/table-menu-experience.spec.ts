import { expect, test, type Page } from './fixtures';

const source = '| 名称 | 内容 |\n| :---: | ---: |\n| **加粗** | 原内容 |';

async function openDocument(page: Page, content = source) {
  await page.addInitScript((content) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content, sourceMode: false, filePath: '', isDirty: false,
    }));
  }, content);
  await page.goto('/');
  await expect(page.locator('.preview table')).toBeVisible();
}

test('单击表格保留格式与尺寸，直接在单元格中输入', async ({ page }) => {
  await openDocument(page);
  const cell = page.locator('.preview td').first();
  const before = await cell.boundingBox();
  await cell.click();
  await expect(cell).toHaveAttribute('contenteditable', 'true');
  await expect(cell).toBeFocused();
  await expect(cell.locator('strong')).toHaveText('加粗');
  await expect(page.getByTestId('preview-table-cell-editor')).toHaveCount(0);
  expect(await cell.boundingBox()).toEqual(before);
  await expect(cell).toHaveCSS('text-align', 'center');
  await cell.press('Control+End');
  await page.keyboard.insertText('新增');
  await cell.press('Control+Enter');
  await expect(cell).toContainText('新增');
  await expect(cell.locator('strong')).toContainText('加粗');
});

test('单元格写回保留对齐和竖线，撤销重做有效', async ({ page }) => {
  await openDocument(page);
  const cell = page.locator('.preview td').last();
  await cell.click();
  await expect(cell).toHaveAttribute('contenteditable', 'true');
  await cell.press('Control+a');
  await page.keyboard.insertText('甲|乙');
  await cell.press('Control+Enter');
  await expect(cell).toHaveText('甲|乙');
  await cell.click();
  await cell.press('Control+z');
  await expect(cell).toHaveText('原内容');
  await cell.press('Control+Shift+z');
  await expect(cell).toHaveText('甲|乙');
  await page.locator('#toggle-sourceview-btn').check();
  await expect(page.locator('.cm-content')).toContainText('甲\\|乙');
  await expect(page.locator('.cm-content')).toContainText('| :---: | ---: |');
});

test('Tab 与 Enter 导航并追加末行，Escape 撤回当前编辑', async ({ page }) => {
  await openDocument(page);
  await page.locator('.preview th').first().click();
  await page.keyboard.press('Tab');
  await expect(page.locator('.preview th').last()).toBeFocused();
  await page.keyboard.press('Enter');
  const last = page.locator('.preview tbody tr').first().locator('td').last();
  await expect(last).toBeFocused();
  await last.press('Control+a');
  await page.keyboard.insertText('取消内容');
  await last.press('Escape');
  await expect(last).toHaveText('原内容');
  await last.click();
  await last.press('Tab');
  await expect(page.locator('.preview tbody tr')).toHaveCount(2);
  await expect(page.locator('.preview tbody tr').last().locator('td').first()).toBeFocused();
});

test('中文组合输入期间不导航，结束后保存编辑结果', async ({ page }) => {
  await openDocument(page);
  const cell = page.locator('.preview td').last();
  await cell.click();
  await expect(cell).toHaveAttribute('contenteditable', 'true');
  await cell.evaluate((node) => {
    node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    node.textContent = '中文组合';
    node.dispatchEvent(new InputEvent('input', { bubbles: true, isComposing: true }));
    node.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true,
      isComposing: true }));
  });
  await expect(cell).toBeFocused();
  await expect(page.locator('.preview tbody tr')).toHaveCount(1);
  await cell.evaluate((node) => node.dispatchEvent(new CompositionEvent('compositionend',
    { bubbles: true, data: '中文组合' })));
  await cell.press('Control+Enter');
  await expect(cell).toHaveText('中文组合');
  await page.locator('#toggle-sourceview-btn').check();
  await expect(page.locator('.cm-content')).toContainText('中文组合');
});

test('Alt 单击与复杂公式单元格保留源码入口', async ({ page }) => {
  await openDocument(page, '| 名称 | 公式 |\n| --- | --- |\n| 原内容 | $x^2$ |');
  await page.locator('.preview td').first().click({ modifiers: ['Alt'] });
  await expect(page.getByTestId('preview-table-cell-editor-input')).toHaveValue('原内容');
  await page.getByTestId('preview-table-cell-editor-input').press('Escape');
  await expect(page.locator('.preview .katex')).toBeVisible();
  await page.locator('.preview td').last().click();
  await expect(page.getByTestId('preview-table-cell-editor-input')).toHaveValue('$x^2$');
});

test('右键菜单采用统一样式，顶部弹出时靠近实际点击位置', async ({ page }) => {
  await openDocument(page);
  const cell = page.locator('.preview td').first();
  const bounds = (await cell.boundingBox())!;
  const point = { x: bounds.x + 10, y: bounds.y + 10 };
  await page.mouse.click(point.x, point.y, { button: 'right' });
  const menu = page.getByTestId('preview-table-menu');
  await expect(menu).toBeVisible();
  await expect(menu).toHaveCSS('position', 'fixed');
  await expect(menu).toHaveCSS('font-size', '13px');
  await expect(menu).toHaveCSS('width', '224px');
  await expect(menu.getByRole('button').first()).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(menu.getByRole('button').nth(1)).toBeFocused();
  const placed = (await menu.boundingBox())!;
  expect(Math.abs(placed.y - point.y)).toBeLessThanOrEqual(1);
  await page.keyboard.press('Escape');
  await expect(menu).toHaveCount(0);
});

test('表格行列快捷键仍可用，Shift Enter 保留单元格内换行', async ({ page }) => {
  await openDocument(page);
  const cell = page.locator('.preview [contenteditable="true"]');
  await page.locator('.preview td').last().click();
  await cell.press('Control+Shift+Enter');
  await expect(page.locator('.preview tbody tr')).toHaveCount(2);
  await expect(cell).toBeFocused();
  await cell.press('Control+Shift+Alt+Enter');
  await expect(page.locator('.preview th')).toHaveCount(3);
  await cell.press('Control+Shift+Alt+Delete');
  await expect(page.locator('.preview th')).toHaveCount(2);
  await cell.press('Control+Shift+Delete');
  await expect(page.locator('.preview tbody tr')).toHaveCount(1);
  await cell.fill('首行');
  await cell.press('Control+End');
  await cell.press('Shift+Enter');
  await page.keyboard.insertText('次行');
  await cell.press('Control+Enter');
  await page.locator('.preview td').last().click();
  await expect(cell).toBeFocused();
  await expect(page.getByTestId('preview-table-cell-editor')).toHaveCount(0);
  await page.locator('#toggle-sourceview-btn').check();
  await expect(page.locator('.cm-content')).toContainText('首行<br>次行');
});

test('单元格富文本粘贴保留格式并过滤危险链接', async ({ page }) => {
  await openDocument(page);
  const cell = page.locator('.preview td').last();
  await cell.click();
  await cell.press('Control+a');
  await cell.evaluate((node) => {
    const data = new DataTransfer();
    data.setData('text/html', '<b>粘贴格式</b><a href="javascript:alert(1)">安全文字</a>');
    node.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data,
      bubbles: true, cancelable: true }));
  });
  await expect(cell.locator('b, strong')).toHaveText('粘贴格式');
  await cell.press('Control+Enter');
  await page.locator('#toggle-sourceview-btn').check();
  await expect(page.locator('.cm-content')).toContainText('**粘贴格式**安全文字');
  await expect(page.locator('.cm-content')).not.toContainText('javascript:');
});

test('单元格刚输入即右键不会丢失内容，列对齐菜单仍可用', async ({ page }) => {
  await openDocument(page);
  const cell = page.locator('.preview td').last();
  await cell.click();
  await cell.fill('立即提交');
  await cell.click({ button: 'right' });
  const menu = page.getByTestId('preview-table-menu');
  await expect(menu).toBeVisible();
  await menu.getByRole('button', { name: '列居中', exact: true }).click();
  await expect(cell).toHaveText('立即提交');
  await expect(cell).toHaveCSS('text-align', 'center');
});

test('底部菜单完整显示，滚动与模式切换关闭菜单，深色主题保持可读', async ({ page }) => {
  await openDocument(page, Array.from({ length: 20 }, () => '正文').join('\n\n') + '\n\n' + source);
  await page.locator('.preview').evaluate((node) => { node.scrollTop = node.scrollHeight; });
  await page.locator('.preview td').first().click({ button: 'right' });
  const menu = page.getByTestId('preview-table-menu');
  await expect(menu).toBeVisible();
  const box = (await menu.boundingBox())!;
  const container = (await page.getByTestId('editor-preview').boundingBox())!;
  expect(box.y + box.height).toBeLessThanOrEqual(container.y + container.height);
  await page.screenshot({ path: test.info().outputPath('table-menu-light.png') });
  await page.locator('.preview').dispatchEvent('wheel', { deltaY: -100 });
  await expect(menu).toHaveCount(0);
  await page.evaluate(() => { document.documentElement.dataset.theme = 'dark'; });
  await page.locator('.preview td').first().click({ button: 'right' });
  await expect(menu).toBeVisible();
  await expect(menu).toHaveCSS('background-color', 'rgb(38, 40, 45)');
  await page.screenshot({ path: test.info().outputPath('table-menu-dark.png') });
  await page.locator('#toggle-sourceview-btn').check();
  await expect(menu).toHaveCount(0);
});
