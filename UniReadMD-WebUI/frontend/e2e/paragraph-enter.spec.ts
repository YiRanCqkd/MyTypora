import { expect, test, type Page } from './fixtures';

async function open(page: Page, content = '正文') {
  await page.addInitScript((content) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content, filePath: '', isDirty: false, sourceMode: false,
    }));
  }, content);
  await page.goto('/');
  await expect(page.locator('.preview p').first()).toBeVisible();
}

test('正文行末回车仅增加一行，光标在空白行首，继续输入和保存不增加空行', async ({ page }) => {
  await open(page, '正文\n\n后续段落');
  const paragraph = page.locator('.preview p').first();
  await paragraph.click();
  await paragraph.press('End');
  const before = (await paragraph.boundingBox())!;
  const lineHeight = await paragraph.evaluate((node) =>
    parseFloat(getComputedStyle(node).lineHeight));
  await paragraph.press('Enter');
  await expect(paragraph).toBeFocused();
  expect((await paragraph.boundingBox())!.height - before.height).toBeCloseTo(lineHeight, 0);
  expect(await paragraph.evaluate((node) => {
    const range = window.getSelection()!.getRangeAt(0).cloneRange();
    range.setEnd(node, node.childNodes.length);
    return range.toString();
  })).toBe('');
  await page.keyboard.insertText('下一行');
  await paragraph.press('Control+s');
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  await expect(page.locator('.preview > p')).toHaveCount(2);
  await expect(page.locator('.preview > p').first()).toHaveText('正文\n下一行');
  await page.screenshot({ path: test.info().outputPath('paragraph-single-enter.png') });
  await page.locator('#toggle-sourceview-btn').check();
  expect(await page.locator('.cm-line').allTextContents())
    .toEqual(['正文', '下一行', '', '后续段落']);
});

test('连续回车每次只新增一行，未输入文字时保存仍保留真实换行', async ({ page }) => {
  await open(page);
  const paragraph = page.locator('.preview p');
  await paragraph.click();
  await paragraph.press('End');
  const before = (await paragraph.boundingBox())!;
  const lineHeight = await paragraph.evaluate((node) =>
    parseFloat(getComputedStyle(node).lineHeight));
  await paragraph.press('Enter');
  await paragraph.press('Enter');
  expect((await paragraph.boundingBox())!.height - before.height).toBeCloseTo(lineHeight * 2, 0);
  await paragraph.press('Control+s');
  await expect(page.locator('.dirty-dot')).toHaveCount(0);
  await page.locator('#toggle-sourceview-btn').check();
  const lines = await page.locator('.cm-line').allTextContents();
  expect(lines.map((line) => line.replace(/\u200b/g, '')))
    .toEqual(['正文', '', '']);
});

test('换行后段首退格合并，撤销和重做仍恢复一行换行', async ({ page }) => {
  await open(page);
  const paragraph = page.locator('.preview p');
  await paragraph.click();
  await paragraph.press('End');
  await paragraph.press('Enter');
  await page.keyboard.insertText('第二行');
  await paragraph.press('Home');
  await paragraph.press('Backspace');
  await expect(paragraph).toHaveText('正文第二行');
  await paragraph.press('Control+z');
  await expect(paragraph).toContainText('第二行');
  await expect(paragraph.locator('br')).toHaveCount(1);
  await paragraph.press('Control+y');
  await expect(paragraph).toHaveText('正文第二行');
});

test('正文源码备用编辑回车也仅插入一个换行，标题和列表保留结构行为', async ({ page }) => {
  await open(page, '正文\n\n# 标题\n\n- 项目');
  await page.locator('.preview p').click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await input.press('End');
  await input.press('Enter');
  await expect(input).toHaveValue('正文\n');
  await page.keyboard.insertText('下一行');
  await input.press('Control+Enter');
  await page.locator('.preview h1').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.insertText('标题后的正文');
  await page.keyboard.press('Control+Enter');
  await expect(page.locator('.preview > p').last()).toHaveText('标题后的正文');
  await page.locator('.preview li').click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.insertText('下一项目');
  await page.keyboard.press('Control+Enter');
  await expect(page.locator('.preview li')).toHaveCount(2);
});
