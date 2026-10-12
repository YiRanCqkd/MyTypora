import { expect, test, type Page } from './fixtures';

async function openDocument(page: Page, content: string) {
  await page.addInitScript((content) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content, filePath: '', isDirty: false, sourceMode: false,
    }));
  }, content);
  await page.goto('/');
  await expect(page.locator('.preview')).toBeVisible();
}

async function sourceLines(page: Page) {
  await page.locator('#toggle-sourceview-btn').check();
  return (await page.locator('.cm-line').allTextContents())
    .map((line) => line.replace(/\u200b/g, ''));
}

for (const text of ['second', '**second**', '*second*', '`second`']) {
  test(`第二段开头回车再退格恢复原段落边界：${text}`, async ({ page }) => {
    await openDocument(page, `first\n\n${text}`);
    const paragraph = page.locator('.preview > p').nth(1);
    await paragraph.click();
    await paragraph.press('Control+Home');
    await paragraph.press('Enter');
    await paragraph.press('Backspace');
    await expect(page.locator('.preview > p')).toHaveCount(2);
    await paragraph.press('Control+s');
    expect(await sourceLines(page)).toEqual(['first', '', text]);
  });
}

test('第二段开头连续两个换行分别退格，不越过段落边界', async ({ page }) => {
  await openDocument(page, 'first\n\nsecond');
  const paragraph = page.locator('.preview > p').nth(1);
  await paragraph.click();
  await paragraph.press('Control+Home');
  await paragraph.press('Enter');
  await paragraph.press('Enter');
  await paragraph.press('Backspace');
  await paragraph.press('Backspace');
  await paragraph.press('Control+s');
  expect(await sourceLines(page)).toEqual(['first', '', 'second']);
});

test('真实段首退格仍可合并相邻正文段落', async ({ page }) => {
  await openDocument(page, 'first\n\nsecond');
  const paragraph = page.locator('.preview > p').nth(1);
  await paragraph.click();
  await paragraph.press('Control+Home');
  await paragraph.press('Backspace');
  await expect(page.locator('.preview > p')).toHaveCount(1);
  expect(await sourceLines(page)).toEqual(['firstsecond']);
});

const table = '| A | B |\n| --- | --- |\n| one | two |';

for (const committed of [false, true]) {
  test(`上方正文换行后右键表格重新定位，已提交=${committed}`, async ({ page }) => {
    await openDocument(page, `正文\n\n${table}\n\n${table}`);
    const paragraph = page.locator('.preview > p').first();
    await paragraph.click();
    await paragraph.press('End');
    await paragraph.press('Enter');
    if (committed) {
      await page.keyboard.insertText('新增正文');
      // 等待输入防抖提交，预览仍然冻结在原位编辑前的源码范围。
      await page.waitForTimeout(150);
      await page.locator('.preview table').last().locator('td').first()
        .click({ button: 'right' });
    } else {
      // 同一事件循环输入并右键，验证尚未提交的文字也不会使表格定位过期。
      await paragraph.evaluate((node) => {
        document.execCommand('insertText', false, '新增正文');
        const cell = node.closest('.preview')!.querySelectorAll('table')[1]
          .querySelector('td')!;
        cell.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true,
          cancelable: true, button: 2, clientX: 400, clientY: 250 }));
      });
    }
    const menu = page.getByTestId('preview-table-menu');
    await expect(menu).toBeVisible();
    await menu.getByRole('button', { name: '在下方插入行', exact: true }).click();
    await expect(page.locator('.preview table').first().locator('tbody tr')).toHaveCount(1);
    await expect(page.locator('.preview table').last().locator('tbody tr')).toHaveCount(2);
    expect(await sourceLines(page)).toEqual([
      '正文', '新增正文', '', ...table.split('\n'), '', ...table.split('\n'), '|  |  |',
    ]);
  });
}

test('正文删除换行后右键表格使用缩短后的范围', async ({ page }) => {
  await openDocument(page, `首行\n次行\n\n${table}`);
  const paragraph = page.locator('.preview > p');
  await paragraph.click();
  await paragraph.press('Control+End');
  await paragraph.press('Home');
  await paragraph.press('Backspace');
  await page.locator('.preview td').first().click({ button: 'right' });
  await page.getByTestId('preview-table-menu')
    .getByRole('button', { name: '在下方插入行', exact: true }).click();
  expect(await sourceLines(page)).toEqual(['首行次行', '', ...table.split('\n'), '|  |  |']);
});

test('省略竖线的合法表格可直接编辑，保存不丢失其他数据行', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await openDocument(page, '| A |\n| --- |\none\ntwo');
  const cell = page.locator('.preview td').last();
  await cell.click();
  await expect(cell).toHaveAttribute('contenteditable', 'true');
  await cell.fill('修改值');
  await cell.press('Control+s');
  await expect(page.locator('.preview tbody tr')).toHaveCount(2);
  expect(await sourceLines(page)).toEqual(['| A |', '| --- |', '| one |', '| 修改值 |']);
  expect(errors).toEqual([]);
});

test('无法解析目标单元格值时安全回退到源码编辑入口', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await openDocument(page, table);
  // 模拟预览映射无法覆盖数据行，验证异常路径仍可回退。
  await page.locator('.preview table').evaluate((node) => { node.dataset.sourceEnd = '2'; });
  await page.locator('.preview td').first().click();
  await expect(page.getByTestId('preview-table-cell-editor-input')).toBeVisible();
  expect(errors).toEqual([]);
});
