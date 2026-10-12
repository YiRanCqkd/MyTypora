import { expect, test, type Page } from './fixtures';

const reference = '/scaling-image.svg';
async function openImages(page: Page, content: string, theme = 'light') {
  await page.route('**/scaling-image.svg', (route) => route.fulfill({
    contentType: 'image/svg+xml',
    body: '<svg xmlns="http://www.w3.org/2000/svg" width="800" height="400">'
      + '<rect width="800" height="400" fill="#5a9aaf"/></svg>',
  }));
  await page.addInitScript(({ content, theme }) => {
    if (!localStorage.getItem('unireadmd.phase2.session.v1')) {
      localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
        content, sourceMode: false, filePath: '', isDirty: true, cursorHint: 1,
      }));
    }
    localStorage.setItem('unireadmd.web-theme.v1', theme);
  }, { content, theme });
  await page.goto('/');
  await expect(page.locator('.preview img').first()).toBeVisible();
  await expect.poll(() => page.locator('.preview img').first()
    .evaluate((image: HTMLImageElement) => image.naturalWidth)).toBe(800);
}
const menu = (page: Page) => page.getByRole('dialog', { name: '图片显示比例' });

for (const theme of ['light', 'dark']) {
  test(`${theme} 图片右键设置比例，保持纵横比并只修改选中图片`, async ({ page }) => {
    await openImages(page, `![同图](${reference}) 和 ![同图](${reference})`, theme);
    const images = page.locator('.preview img');
    await images.nth(1).click({ button: 'right' });
    await expect(menu(page)).toBeVisible();
    await expect(menu(page)).toContainText('800 × 400');
    await page.screenshot({ path: test.info().outputPath(`image-menu-${theme}.png`) });
    await menu(page).getByRole('button', { name: '50%', exact: true }).click();
    await expect(menu(page)).toHaveCount(0);
    await expect(images.nth(1)).toHaveAttribute('width', '400');
    await expect(images.first()).not.toHaveAttribute('width');
    const dimensions = await images.nth(1).boundingBox();
    expect(dimensions!.width / dimensions!.height).toBeCloseTo(2);
    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.locator('.cm-content')).toContainText(`![同图](${reference}) 和 <img`);
    await page.locator('.cm-content').focus();
    await page.keyboard.press('Control+z');
    await expect(page.locator('.cm-content')).toContainText(
      `![同图](${reference}) 和 ![同图](${reference})`,
    );
  });
}

test('自定义比例校验、持久化、恢复自动和窗口宽度约束', async ({ page }) => {
  await openImages(page, `![图](${reference})`);
  const image = page.locator('.preview img');
  await image.click({ button: 'right' });
  const custom = menu(page).getByRole('spinbutton', { name: '自定义比例（%）' });
  await custom.fill('0');
  await menu(page).getByRole('button', { name: '应用比例' }).click();
  await expect(menu(page)).toBeVisible();
  await expect(image).not.toHaveAttribute('width');
  await custom.fill('37.5');
  await custom.press('Enter');
  await expect(image).toHaveAttribute('width', '300');
  await page.waitForTimeout(1300);
  await page.reload();
  await expect(image).toHaveAttribute('width', '300');
  await image.click({ button: 'right' });
  await menu(page).getByRole('button', { name: '200%', exact: true }).click();
  await expect(image).toHaveAttribute('width', '1600');
  const bounds = await image.boundingBox();
  const preview = await page.locator('.preview').boundingBox();
  expect(bounds!.width).toBeLessThanOrEqual(preview!.width);
  expect(bounds!.width / bounds!.height).toBeCloseTo(2);
  await image.click({ button: 'right' });
  await menu(page).getByRole('button', { name: '恢复自动大小' }).click();
  await expect(image).not.toHaveAttribute('width');
  await expect(image).not.toHaveAttribute('style');
});

test('引用图片缩放后外层链接与标题仍存在，导出保留尺寸', async ({ page }) => {
  await openImages(page, `[![图][pic]](https://example.com)\n\n[pic]: ${reference} "标题"`);
  const image = page.locator('.preview img');
  await image.click({ button: 'right' });
  await menu(page).getByRole('button', { name: '25%', exact: true }).click();
  await expect(image).toHaveAttribute('title', '标题');
  await expect(image.locator('..')).toHaveAttribute('href', 'https://example.com');
  await page.evaluate(() => {
    window.nativeBridge.exporter.exportDocument = async (payload) => {
      Reflect.set(window, '__imageExport', payload.html);
      return { filePath: '/virtual/images.html', format: 'html', engine: 'builtin' };
    };
  });
  await page.getByRole('combobox', { name: '导出格式' }).selectOption('html');
  await page.getByRole('button', { name: '导出', exact: true }).click();
  await expect.poll(() => page.evaluate(() => {
    const html = String(Reflect.get(window, '__imageExport') || '');
    return new DOMParser().parseFromString(html, 'text/html').querySelector('img')
      ?.getAttribute('width');
  })).toBe('200');
});

test('菜单键盘操作、关闭行为及已有 HTML 图片属性保留', async ({ page }) => {
  await openImages(page, `<img src="${reference}" title="说明" style="border-radius: 4px">`);
  const image = page.locator('.preview img');
  await image.click({ button: 'right' });
  await expect(menu(page).getByRole('button', { name: '25%', exact: true })).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('Enter');
  await expect(image).toHaveAttribute('width', '400');
  await expect(image).toHaveCSS('border-radius', '4px');
  await image.click({ button: 'right' });
  await page.keyboard.press('Escape');
  await expect(menu(page)).toHaveCount(0);
  await image.click({ button: 'right' });
  await page.setViewportSize({ width: 900, height: 700 });
  await expect(menu(page)).toHaveCount(0);
  await image.click({ button: 'right' });
  await page.locator('#toggle-sourceview-btn').click();
  await expect(menu(page)).toHaveCount(0);
});

test('图片加载失败时禁用比例按钮，仍可恢复自动大小', async ({ page }) => {
  await openImages(page, `![图](${reference})\n\n<img src="/missing-image.png" width="200">`);
  const image = page.locator('.preview img').last();
  await image.click({ button: 'right' });
  await expect(menu(page).getByRole('button', { name: '50%', exact: true })).toBeDisabled();
  await menu(page).getByRole('button', { name: '恢复自动大小' }).click();
  await expect(image).not.toHaveAttribute('width');
});
